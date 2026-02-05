# -*- coding: utf-8 -*-
"""
Created on Tue Dec  2 12:57:17 2025

@author: ravis
"""
import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz, process
from docx import Document
from docx.oxml import OxmlElement
from io import BytesIO
import tempfile
from pathlib import Path
import zipfile
from docx.shared import RGBColor
from docxcompose.composer import Composer
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Pt

# import pypandoc
# =========================
#  YOUR HELPER FUNCTIONS
# =========================

# def docx_to_pdf(docx_path, pdf_path):
#     try:
#         pypandoc.convert_file(
#             source_file=str(docx_path),
#             to="pdf",
#             outputfile=str(pdf_path),
#             extra_args=[
#                 "--pdf-engine=xelatex",
#                 "-V", "mainfont=Arial",
#                 "-V", "sansfont=Arial",
#                 "-V", "monofont=Courier New",
#             ],
#         )
#         return True
#     except Exception as e:
#         # Optional: log or collect errors
#         print(f"PDF conversion failed for {docx_path}: {e}")
#         return False
#
# def create_reviewer_pdf_packets(assignments_df, processed_dir, out_zip_path):
#     import zipfile
#     reviewer_packets = {}
#     processed_dir = Path(processed_dir)

#     for reviewer, group in assignments_df.groupby("reviewer_name"):
#         abstract_nums = group["abstract_id"].dropna().astype(int).tolist()
#         abstract_nums_sorted = sorted(abstract_nums)

#         doc_paths = []
#         for num in abstract_nums_sorted:
#             filename = f"srd_abstract_{num}.docx"
#             path = processed_dir / filename
#             if path.exists():
#                 doc_paths.append(path)

#         if not doc_paths:
#             continue

#         # DOCX output name
#         nums_str = "-".join(str(x) for x in abstract_nums_sorted)
#         docx_out = processed_dir / f"{reviewer}_Abstracts_{nums_str}.docx"
#         pdf_out = processed_dir / f"{reviewer}_Abstracts_{nums_str}.pdf"

#         # Merge DOCX
#         merge_docx_files(doc_paths, docx_out)

#         # Convert to PDF
#         docx_to_pdf(docx_out, pdf_out)

#         reviewer_packets[reviewer] = pdf_out

#     # ZIP all PDFs
#     with zipfile.ZipFile(out_zip_path, "w") as zf:
#         for reviewer, pdf_file in reviewer_packets.items():
#             zf.write(pdf_file, arcname=pdf_file.name)

#     return out_zip_path



def recompress_docx_inplace(docx_path: str | Path, remove_thumbnail: bool = True) -> Path:
    """
    Re-compress a .docx file without changing its visible content.

    What it does:
      - Reads the DOCX as a zip
      - Writes a new zip with ZIP_DEFLATED compression (smaller)
      - Optionally removes docProps/thumbnail.jpeg (not part of the document content)

    What it does NOT do:
      - It does not edit any XML, text, styles, images, or relationships
      - It does not downsample images
      - It does not alter the document layout/content

    Returns:
      Path to the recompressed docx (same path, in-place).
    """
    docx_path = Path(docx_path)

    if docx_path.suffix.lower() != ".docx":
        raise ValueError(f"Expected a .docx file, got: {docx_path}")

    if not docx_path.exists():
        raise FileNotFoundError(docx_path)

    # Create temp output in same directory (safe replace)
    tmp_path = docx_path.with_suffix(".recompressed.tmp")

    # Choose compression
    compression = zipfile.ZIP_DEFLATED

    # compresslevel is available on Python 3.7+ for ZipFile
    zip_kwargs = {"compression": compression}
    try:
        zip_kwargs["compresslevel"] = 9
    except TypeError:
        # Older Python: ignore compresslevel
        pass

    with zipfile.ZipFile(docx_path, "r") as zin:
        # Validate it is a DOCX-like zip
        names = zin.namelist()
        if "[Content_Types].xml" not in names or "word/document.xml" not in names:
            raise ValueError("This file doesn't look like a valid DOCX.")

        with zipfile.ZipFile(tmp_path, "w", **zip_kwargs) as zout:
            for info in zin.infolist():
                name = info.filename

                # Thumbnail is not document content (Word preview image)
                if remove_thumbnail and name.lower() == "docprops/thumbnail.jpeg":
                    continue

                data = zin.read(name)

                # Preserve timestamps/metadata as much as ZipInfo allows
                new_info = zipfile.ZipInfo(filename=name, date_time=info.date_time)
                new_info.compress_type = compression
                new_info.external_attr = info.external_attr
                new_info.internal_attr = info.internal_attr
                new_info.flag_bits = info.flag_bits

                # Keep original "stored vs deflated" isn't needed; we always deflate for size
                zout.writestr(new_info, data)

    # Atomic-ish replace
    tmp_path.replace(docx_path)
    return docx_path


def recompress_all_docx_in_folder(folder: str | Path, remove_thumbnail: bool = True) -> tuple[int, int]:
    """
    Recompress all .docx files in a folder in place.
    Returns (processed_count, failed_count).
    """
    folder = Path(folder)
    processed = 0
    failed = 0

    for p in folder.glob("*.docx"):
        try:
            recompress_docx_inplace(p, remove_thumbnail=remove_thumbnail)
            processed += 1
        except Exception:
            failed += 1

    return processed, failed


def ensure_output_dir():
    """
    Create (once) a stable temp output directory for this Streamlit session.
    Prevents RAM blowups by keeping large zips on disk.
    """
    if "output_dir" not in st.session_state:
        st.session_state["output_dir"] = tempfile.mkdtemp(prefix="srd_streamlit_")
    return Path(st.session_state["output_dir"])

def write_bytesio_to_file(buf: BytesIO, out_path: Path):
    out_path.write_bytes(buf.getvalue())
    return str(out_path)



def force_document_font(doc, font_name="Arial", font_size=12):
    from docx.shared import Pt

    # ---- 1) Update Normal style ----
    try:
        normal = doc.styles["Normal"]
        normal.font.name = font_name
        normal.font.size = Pt(font_size)
    except:
        pass

    # ---- 2) Loop over all paragraphs & runs ----
    for paragraph in doc.paragraphs:
        # Update paragraph style if possible
        try:
            paragraph.style.font.name = font_name
            paragraph.style.font.size = Pt(font_size)
        except:
            pass

        for run in paragraph.runs:
            
            has_image = bool(run._r.xpath(".//w:drawing")) or bool(run._r.xpath(".//w:pict"))

            if has_image:
                continue  # DO NOT TOUCH IMAGE RUNS

            # Otherwise update font
            run.font.size = Pt(font_size)

    # ---- 3) Update table text safely ----
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        # Skip images inside table cells too
                        has_image = bool(run._r.xpath(".//w:drawing")) or bool(run._r.xpath(".//w:pict"))

                        if has_image:
                            continue

                        run.font.name = font_name
                        run.font.size = Pt(font_size)

def clean_whitespace(doc):
    """
    Remove paragraphs that are visually empty:
      - no text (after stripping)
      - no images
      - no tables
      - no page breaks
    This WILL remove those big empty blocks, but keeps figures.
    """
    for p in list(doc.paragraphs):  # iterate over a copy
        # 1. All visible text in this paragraph
        text = "".join(run.text for run in p.runs).strip()

        # 2. Does this paragraph contain an image?
        has_image = bool(p._element.xpath(".//w:drawing")) or bool(
            p._element.xpath(".//w:pict")
        )

        # 3. Does it contain a table?
        has_table = bool(p._element.xpath(".//w:tbl"))

        # 4. Does it contain a page break?
        has_pagebreak = any(
            run._r.xpath(".//w:br[@w:type='page']")
            for run in p.runs
        )

        # 5. Only remove if it's truly "visually empty"
        if text == "" and not (has_image or has_table or has_pagebreak):
            p._element.getparent().remove(p._element)



            
def merge_docx_files(doc_paths, output_path, font_name="Arial", font_size=12):
    if not doc_paths:
        return

    master = Document(doc_paths[0])
    composer = Composer(master)

    # Page breaks between abstracts
    for doc_path in doc_paths[1:]:
        p = master.add_paragraph()
        p.add_run().add_break(WD_BREAK.PAGE)
        composer.append(Document(doc_path))

    composer.save(output_path)

    # Reload merged doc
    merged_doc = Document(output_path)

    # ---- SAFE FORMAT APPLICATION ----

    for paragraph in merged_doc.paragraphs:
        for run in paragraph.runs:

            # Skip image runs
            has_image = bool(run._r.xpath(".//w:drawing")) or bool(run._r.xpath(".//w:pict"))
            if has_image:
                continue

            run.font.name = font_name
            run.font.size = Pt(font_size)

    # Apply safe table formatting
    for table in merged_doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:

                        has_image = bool(run._r.xpath(".//w:drawing")) or bool(run._r.xpath(".//w:pict"))
                        if has_image:
                            continue

                        run.font.name = font_name
                        run.font.size = Pt(font_size)

    merged_doc.save(output_path)

    
def create_reviewer_docx_packets(assignments_df, processed_dir, out_zip_path):
    reviewer_packets = {}
    processed_dir = Path(processed_dir)

    for reviewer, group in assignments_df.groupby("reviewer_name"):

        # Expected abstract numbers based on assignments
        abstract_nums = group["abstract_id"].dropna().astype(int).tolist()
        abstract_nums_sorted = sorted(abstract_nums)

        doc_paths = []
        filename_parts = []   # we will fill this with numbers or "missingX"

        for num in abstract_nums_sorted:
            doc_path = processed_dir / f"srd_abstract_{num}.docx"

            if doc_path.exists():
                doc_paths.append(doc_path)
                filename_parts.append(str(num))
            else:
                # Mark missing abstract in filename
                filename_parts.append(f"missing{num}")

        # If no available docs → skip
        if not doc_paths:
            continue

        # Filename now includes missing abstracts explicitly
        nums_str = "-".join(filename_parts)
        out_file = f"{reviewer}_Abstracts_{nums_str}.docx"
        out_path = processed_dir / out_file

        # Merge DOCX files (only existing ones)
        merge_docx_files(
            doc_paths,
            output_path=out_path
        )

        reviewer_packets[reviewer] = out_path

    # ZIP the packets
    with zipfile.ZipFile(
    out_zip_path,
    "w",
    compression=zipfile.ZIP_DEFLATED,
    compresslevel=9,
    ) as zf:
        for reviewer, docx_file in reviewer_packets.items():
            zf.write(docx_file, arcname=docx_file.name)

    return out_zip_path



def clean_name(name):
    if name is None:
        return None

    # 1. Remove superscripts like ^a, ^1, ^xyz
    name = re.sub(r"\^[A-Za-z0-9]+", "", name)

    # 2. Remove a caret at the end (just "^")
    name = re.sub(r"\^$", "", name)

    # 3. Remove any remaining stray "^"
    name = name.replace("^", "")

    # Remove asterisks
    name = name.replace("*", "")

    # Remove footnote digits at end of name
    name = re.sub(r"\d+$", "", name)

    # Collapse extra spaces
    name = re.sub(r"\s+", " ", name)

    return name.strip()


def extract_docx_text_with_superscripts(filepath):
    doc = Document(filepath)
    lines = []
    for p in doc.paragraphs:
        line = ""
        for r in p.runs:
            r_text = r.text
            r_elem = r._element

            # Check if run is superscript
            vert_align = r_elem.find(
                ".//w:vertAlign",
                namespaces={
                    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                },
            )
            if (
                vert_align is not None
                and vert_align.attrib.get(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
                )
                == "superscript"
            ):
                # Convert superscript run to ^text
                line += "^" + r_text
            else:
                line += r_text

        if line.strip():
            lines.append(line.strip())

    return lines



MARKER_TIERS = [
    # Tier 1: most specific (preferred)
    [
        "choose your research type",
        "kies uw onderzoekstype",
        "your research type",
        "research type:",
        "onderzoekstype:",
        "type of research:",
    ],

    # Tier 2: explicit category phrases
    [
        "research classification",
        "category of research",
        "clinical research",
        "fundamental research",
        "translational research",
        "basic research",
        "applied research",
    ],

    # Tier 3: still OK but broader
    [
        "research type",
        "type of research",
        "type onderzoek",
        "onderzoek type",
    ],
]

# Tier 4: ultra-generic fallback (only if nothing else found)
FALLBACK_WORDS = ["clinical", "fundamental"]


def _looks_like_option_line(text: str) -> bool:
    """
    Heuristic: accept 'clinical'/'fundamental' only if the paragraph looks like an option/label.
    Reject long sentences like affiliations.
    """
    t = text.strip()

    # too long -> likely sentence / affiliation
    if len(t) > 40:
        return False

    # contains "department", "university", etc. -> likely affiliation/header text
    if re.search(r"\b(department|universit|erasmus|mc|rotterdam|affiliation)\b", t, flags=re.I):
        return False

    # if it starts with a number or bullet, or is just a short label, we allow it
    return True


def find_first_marker(doc):
    # 1) Search tier 1-3 in order
    for tier_idx, tier in enumerate(MARKER_TIERS, start=1):
        tier = [m.lower() for m in tier]

        for i, p in enumerate(doc.paragraphs):
            text = p.text.strip()
            low = text.lower()

            if not low:
                continue

            for marker in tier:
                if marker in low:
                    return {
                        "paragraph_index": i,
                        "marker": marker,
                        "tier": tier_idx,
                        "text": text,
                    }

    # 2) Final fallback tier: "clinical"/"fundamental" ONLY with safeguards
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        low = text.lower()
        if not low:
            continue

        if not _looks_like_option_line(text):
            continue

        for w in FALLBACK_WORDS:
            # whole-word match
            if re.search(rf"\b{re.escape(w)}\b", low):
                return {
                    "paragraph_index": i,
                    "marker": w,
                    "tier": 4,
                    "text": text,
                }

    return None



def extract_surname(name):
    # Titles to strip before surname detection
    TITLE_PATTERNS = [
        r"\bMD\b", r"\bM\.D\.\b", r"\bPhD\b", r"\bP\.h\.D\.\b",
        r"\bMSc\b", r"\bBSc\b", r"\bMBA\b", r"\bMPH\b",
        r"\bDr\b", r"\bDr.\b", r"\bProf\b", r"\bProf.\b",
    ]

    DUTCH_PREFIXES = {
        "van", "de", "der", "den", "het", "ter", "ten",
        "van de", "van der", "van den"
    }
    if not isinstance(name, str):
        return None

    # 1. Clean the name
    clean = name.strip()

    # remove trailing footnotes like: ^, ^*, ^a, digits
    clean = re.sub(r"\^.*$", "", clean)
    clean = re.sub(r"\d+$", "", clean).strip()

    # remove academic titles
    for pattern in TITLE_PATTERNS:
        clean = re.sub(pattern, "", clean, flags=re.IGNORECASE)

    clean = re.sub(r"\s+", " ", clean).strip()

    # split
    parts = clean.split()
    if not parts:
        return None

    # If only one part remains, that's the surname
    if len(parts) == 1:
        return parts[0].lower()

    # 2. Check for Dutch multi-word prefixes
    last_two = " ".join(parts[-2:]).lower()
    if last_two in DUTCH_PREFIXES:
        return last_two  # e.g. "van der"

    # 3. Normal surname = last word
    return parts[-1].lower()

def extract_initials(name):
    if not isinstance(name, str):
        return None

    name = remove_titles(name).strip()

    # Step 1: dotted initials (e.g. F.M.)
    dotted = re.findall(r"([A-Za-z])\.", name)
    if dotted:
        return "".join(d.upper() for d in dotted)

    # Step 2: initials from words (Laura Maria → LM)
    tokens = re.split(r"[ \-]+", name)
    letters = []

    for t in tokens:
        # skip tokens that are too short or numeric
        if t and t[0].isalpha():
            letters.append(t[0].upper())

    return "".join(letters) if letters else None

def remove_titles(name):
    TITLES = {
    r"\bMD\b", r"\bM\.D\.\b", r"\bPhD\b", r"\bP\.h\.D\.\b",
    r"\bMSc\b", r"\bBSc\b",
    r"\bDr\b", r"\bDr.\b", r"\bProf\b", r"\bProf.\b",
    r"\bIng\b", r"\bir\.\b",
    r"\bMBA\b", r"\bMPH\b"}
    if not isinstance(name, str):
        return name
    clean = name
    for t in TITLES:
        clean = re.sub(t, "", clean, flags=re.IGNORECASE)
    clean = re.sub(r"\s+", " ", clean).strip()
    return clean
def match_author_name(input_name, ref_df, ref_col="name"):
    if pd.isna(input_name) or len(str(input_name).strip()) < 2:
        return None

    input_name = str(input_name).strip()
    input_lower = input_name.lower()

    # Extract features from input
    in_surname = extract_surname(input_name)
    in_initials = extract_initials(input_name)

    candidates = [str(c).strip() for c in ref_df[ref_col].dropna().unique()]

    # ---- 1) Exact full-name match
    for c in candidates:
        if c.lower() == input_lower:
            return c

    # ---- 2) Surname exact match filter
    same_surname = [c for c in candidates if extract_surname(c) == in_surname]

    if len(same_surname) == 1:
        return same_surname[0]

    # ---- 3) Initial matching (improved)
    if in_initials:
        scored = []
        for c in same_surname:
            cand_init = extract_initials(c)
            if not cand_init:
                continue

            score = 0

            # NEW: strong bonus for matching first initial
            if cand_init[0] == in_initials[0]:
                score += 300
            else:
                score -= 200

            # Bonus: full initials match
            if cand_init == in_initials:
                score += 200

            # Bonus: same number of initials
            if len(cand_init) == len(in_initials):
                score += 100

            # Fuzzy similarity (0–100)
            score += fuzz.ratio(cand_init, in_initials)

            scored.append((c, score))

        if scored:
            best = max(scored, key=lambda x: x[1])

            # threshold relaxed to allow imperfect initials but correct first initial
            if best[1] >= 150:
                return best[0]

    # ---- 4) Fuzzy surname fallback
    best_surname = None
    best_score = 0
    for c in candidates:
        score = fuzz.partial_ratio(in_surname, extract_surname(c))
        if score > best_score:
            best_score = score
            best_surname = c
    if best_score >= 90:
        return best_surname

    # ---- 5) Fuzzy full-name fallback
    best_match = None
    best_score = 0
    for c in candidates:
        score = fuzz.partial_ratio(input_lower, c.lower())
        if score > best_score:
            best_score = score
            best_match = c
    if best_score >= 85:
        return best_match

    return None

def fuzzy_merge(df1, df2, key1, key2, threshold=90, scorer=fuzz.partial_ratio):
    matches = []
    # Ensure df2 key column is all strings and drop NaNs
    df2_clean = df2.dropna(subset=[key2]).copy()
    df2_clean[key2] = df2_clean[key2].astype(str)

    for idx, row in df1.iterrows():
        val = row[key1]
        if pd.isna(val):
            continue
        name = str(val)

        match = process.extractOne(name, df2_clean[key2].tolist(), scorer=scorer)
        if match and match[1] >= threshold:
            matched_name = match[0]
            matched_row = df2_clean[df2_clean[key2] == matched_name].iloc[0]
            combined = {**row.to_dict(), **matched_row.to_dict()}
            matches.append(combined)

    return pd.DataFrame(matches)


def read_excel_with_auto_header_from_bytes(data: bytes, sheet_name=0):
    """Your original auto-header function, adapted for in-memory bytes."""
    from io import BytesIO

    temp = pd.read_excel(BytesIO(data), sheet_name=sheet_name, header=None)
    header_row = temp.notna().any(axis=1).idxmax()
    df = pd.read_excel(BytesIO(data), sheet_name=sheet_name, header=header_row)
    return df


def extract_ids(x):
    nums = re.findall(r"\d+", str(x))
    return [int(n) for n in nums] if nums else None


def department_conflict(author_deps, reviewer_dep, threshold=85):
    """
    author_deps: list of author's departments (cleaned English)
    reviewer_dep: single reviewer department (cleaned English)
    threshold: min fuzzy similarity to consider a conflict
    """
    if reviewer_dep is None:
        return False

    reviewer_dep = str(reviewer_dep).lower().strip()

    for dep in author_deps:
        if dep is None:
            continue
        dep = str(dep).lower().strip()
        score = fuzz.partial_ratio(dep, reviewer_dep)
        if score >= threshold:
            return True  # conflict
    return False


def assign_reviewers(authors, reviewers, max_reviews=8, reviewers_per_abs=3, conflict_threshold=85):
    """
    authors: DataFrame with columns ['name', 'abstract_id', 'departments']
             where 'departments' is a list of department strings
    reviewers: DataFrame with columns ['reviewer_name', 'reviewer_department', 'assigned_count']
    """
    assignments = []

    for idx, row in authors.iterrows():
        author_name = row["name"]
        abstract_id = row["abstract_id"]
        author_deps = list(row["departments"])

        # 1) Build mask: below max_reviews AND NOT in conflicting department
        mask = (reviewers["assigned_count"] < max_reviews) & (
            ~reviewers["reviewer_department"].apply(
                lambda d: department_conflict(author_deps, d, threshold=conflict_threshold)
            )
        )

        eligible = reviewers[mask]

        # 2) If we don't have enough eligible reviewers, stop with a clear error
        if len(eligible) < reviewers_per_abs:
            raise ValueError(
                f" Not enough eligible reviewers for {author_name} (abstract {abstract_id}). "
                f"Needed {reviewers_per_abs}, but only {len(eligible)} are non-conflicting and below {max_reviews}."
            )

        # 3) Load balancing: prefer reviewers with the fewest assignments
        eligible = eligible.sort_values(by="assigned_count", ascending=True)

        chosen = eligible.head(reviewers_per_abs)

        # 4) Update counts and record assignments
        for i, reviewer in chosen.iterrows():
            reviewers.loc[i, "assigned_count"] += 1

            assignments.append(
                {
                    "author_name": author_name,
                    "abstract_id": abstract_id,
                    "reviewer_name": reviewer["reviewer_name"],
                    "reviewer_department": reviewer["reviewer_department"],
                    "reviewer_load_after": reviewers.loc[i, "assigned_count"],
                    "author_departments": ", ".join(author_deps),
                }
            )

    return pd.DataFrame(assignments), reviewers


def prepare_ref_and_authors(ref_file, trans_file):
    """Prepares ref_df (for name → abstract_nr) and authors DF (for assignments)."""

    # translations
    trans_df = pd.read_excel(trans_file).drop_duplicates()
    
    required_cols = {"English", "department"}
    missing = required_cols - set(trans_df.columns)
    if missing:
        raise ValueError(f"Missing required columns in registrations file: {missing}")
        
        
    ref_df = pd.read_excel(ref_file)
    separators = r"[;,+/]+"
    ref_df.columns = ref_df.columns.str.lower().str.replace(" ", "_")
    
    required_cols = {"name", "department_", "abstract_nr._"}
    missing = required_cols - set(ref_df.columns)
    if missing:
        raise ValueError(f"Missing required columns in registrations file: {missing}")
    
    
    
    # Clean only specific columns:
    for col in ["name", "department_", "abstract_nr._"]:
        if col in ref_df.columns:
            ref_df[col] = ref_df[col].astype(str).str.strip()
    
    # Then handle NaNs in department_ properly
    ref_df["department_"] = ref_df["department_"].replace("nan", "").fillna("")

    # SPLIT into lists
    ref_df["department"] = ref_df["department_"].str.split(separators, regex=True)

    # EXPLODE correctly
    ref_expanded = ref_df.explode("department")

    # Clean
    ref_expanded["department"] = ref_expanded["department"].str.strip()
    ref_expanded = ref_expanded[ref_expanded["department"] != ""]

    # fuzzy merge with translation
    ref_df_merged = fuzzy_merge(ref_expanded, trans_df, "department", "department")

    # Make list of IDs
    ref_df_merged["abstract_id_list"] = ref_df_merged["abstract_nr._"].apply(extract_ids)

    # Explode to create separate rows for each abstract ID
    ref_df_merged = ref_df_merged.explode("abstract_id_list")

    # Rename for convenience
    ref_df_merged = ref_df_merged.rename(columns={"abstract_id_list": "abstract_id"})

    # Build authors DF: one row per (name, abstract_id), with list of English departments
    authors = (
        ref_df_merged.groupby(["name", "abstract_id"])["English"]
        .apply(lambda x: list(set(x.dropna())))
        .reset_index()
        .rename(columns={"English": "departments"})
    )

    return ref_df, authors


def process_doc(filepath, ref_df, output_folder, remaining_ids):

    # -----------------------------
    # 1) Extract raw text lines
    # -----------------------------
    txt = extract_docx_text_with_superscripts(filepath)

    # Helper to find line index
    def find_line(phrase):
        for i, line in enumerate(txt):
            if phrase.lower() in line.lower():
                return i
        return None

    author_line = find_line("author")
    aff_line = find_line("affiliations")

    # If missing metadata → just copy file unchanged
    if aff_line is None or author_line is None:
        doc = Document(filepath)
        out_path = Path(output_folder) / (Path(filepath).stem + "_unchanged.docx")
        doc.save(out_path)
        # NEW: recompress the saved DOCX (no content/layout changes)
        recompress_docx_inplace(out_path)
        return {
            "file": Path(filepath).name,
            "name": None,
            "matched_name": None,
            "abstract_nr": None,
            "matched": False,
        }

    # -----------------------------
    # 2) Extract FIRST author name
    # -----------------------------
    def extract_first_author(txt, author_line_index):
        line = txt[author_line_index]
        line_clean = (
            line.replace("Authors:", "")
                .replace("Author:", "")
                .replace("Authors", "")
                .replace("Author", "")
                .strip()
        )
        if line_clean and not line_clean.lower().startswith("affiliation"):
            first = line_clean.split(",")[0].strip()
            first = re.sub(r"\d+$", "", first).strip()
            if first:
                return first

        # Otherwise check following lines
        i = author_line_index + 1
        while i < len(txt):
            candidate = txt[i].strip()
            if not candidate:
                i += 1
                continue
            if "affiliation" in candidate.lower():
                break
            first = candidate.split(",")[0].strip()
            first = re.sub(r"\d+$", "", first).strip()
            if first:
                return first
            i += 1

        return ""

    name = extract_first_author(txt, author_line)
    name = clean_name(name)

    # -----------------------------
    # 3) Fuzzy-match author name
    # -----------------------------
    matched_name = match_author_name(name, ref_df, "name")

    # -----------------------------
    # 4) Assign abstract number
    # -----------------------------
    abstract_nr = None
    if matched_name:
        key = matched_name.strip()
        if key in remaining_ids and len(remaining_ids[key]) > 0:
            abstract_nr = remaining_ids[key].pop(0)

    # -----------------------------
    # 5) Load the actual DOCX
    # -----------------------------
    doc = Document(filepath)

    # Find the marker to cut off the top
    res = find_first_marker(doc)
    research_idx = res["paragraph_index"] if res else None

    if research_idx is None:
        out_path = Path(output_folder) / (Path(filepath).stem + "_unchanged.docx")
        doc.save(out_path)
        return {
            "file": Path(filepath).name,
            "name": name,
            "matched_name": matched_name,
            "abstract_nr": abstract_nr,
            "matched": abstract_nr is not None,
        }

    # Remove paragraphs before the marker
    for _ in range(research_idx):
        p = doc.paragraphs[1]
        p._element.getparent().remove(p._element)

    body = doc._element.body

    # -----------------------------
    # 6) Clean whitespace BEFORE inserting abstract number
    # -----------------------------
    clean_whitespace(doc)

    # -----------------------------
    if abstract_nr is not None:
        title_p = doc.paragraphs[0]
    
        p_label = doc.add_paragraph()
        p_label.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p_label.add_run(f"Abstract number: {abstract_nr}")
        run.bold = True
        run.font.color.rgb = RGBColor(255, 0, 0)
    
        # Move paragraph directly below title
        body.remove(p_label._p)
        title_p._p.addnext(p_label._p)

    # -----------------------------
    # 8) Insert 5 clean blank lines UNDER header
    # -----------------------------
    for _ in range(5):
        empty_p = OxmlElement("w:p")
        body.insert(1, empty_p)

    # -----------------------------
    # 9) Apply uniform font (Arial, size 12)
    # -----------------------------
    force_document_font(doc, font_name="Arial", font_size=12)

    # -----------------------------
    # 10) Save DOCX
    # -----------------------------
    if abstract_nr is not None:
        output_filename = f"srd_abstract_{abstract_nr}.docx"
    else:
        output_filename = Path(filepath).stem + "_no_number.docx"

    out_path = Path(output_folder) / output_filename
    doc.save(out_path)

    # -----------------------------
    # 11) Return metadata
    # -----------------------------
    return {
        "file": Path(filepath).name,
        "name": name,
        "matched_name": matched_name,
        "abstract_nr": abstract_nr,
        "matched": abstract_nr is not None,
    }

def build_remaining_ids_dict(ref_df):
    """
    Build dictionary:
        name -> sorted list of abstract numbers
    Handles single and multi-number entries (e.g. '04+76').
    """
    id_dict = {}

    for _, row in ref_df.iterrows():
        name = str(row.get("name", "")).strip()
        raw = row.get("abstract_nr._", "")
        
        ids = extract_ids(raw)  # your existing helper
        
        if not ids:
            continue
        
        if name not in id_dict:
            id_dict[name] = []

        for x in ids:
            if x not in id_dict[name]:
                id_dict[name].append(x)

    # Sort numbers for stable assignment
    for k in id_dict:
        id_dict[k] = sorted(id_dict[k])

    return id_dict

def run_pipeline(ref_file, trans_file, reviewer_file, docx_files):
    """Main wrapper: runs everything and returns DFs + bytes for downloads."""
    # Prepare ref_df + authors
    ref_df, authors = prepare_ref_and_authors(ref_file, trans_file)
    remaining_ids = build_remaining_ids_dict(ref_df)

    # Reviewers (with auto-header detection)
    reviewer_bytes = reviewer_file.getvalue()
    reviewer_df = read_excel_with_auto_header_from_bytes(
        reviewer_bytes, sheet_name="Reviewers"
    )

    # Clean reviewer_df column names
    reviewer_df.columns = reviewer_df.columns.str.lower().str.strip()

    # rename correct columns
    reviewer_df = reviewer_df.rename(
        columns={
            "reviewer signup": "reviewer_name",
            "department": "reviewer_department",
        }
    )

    # clean reviewer_department + remove NaN
    reviewer_df["reviewer_department"] = (
        reviewer_df["reviewer_department"].astype(str).str.strip()
    )
    reviewer_df = reviewer_df.dropna(subset=["reviewer_name"])

    # initialize load counter
    reviewer_df["assigned_count"] = 0

    # Assign reviewers
    assignments_df, reviewer_df_updated = assign_reviewers(authors, reviewer_df)

    # Temp folders for DOCX processing
    workdir = Path(tempfile.mkdtemp())
    input_dir = workdir / "input_docs"
    output_dir = workdir / "output_docs"
    input_dir.mkdir(exist_ok=True, parents=True)
    output_dir.mkdir(exist_ok=True, parents=True)

    results = []
    for uf in docx_files:
        if uf is None:
            continue
        file_path = input_dir / uf.name
        with open(file_path, "wb") as f:
            f.write(uf.getvalue())

        res = process_doc(str(file_path), ref_df, str(output_dir), remaining_ids)
        results.append(res)

    results_df = pd.DataFrame(results)

    # Build Excel bytes
    # ---- WRITE LARGE OUTPUTS TO DISK (NOT RAM) ----
    out_dir = ensure_output_dir()

    # 1) Assignments Excel to disk
    assignments_path = out_dir / "reviewer_assignments.xlsx"
    assignments_df.to_excel(assignments_path, index=False)

    # 2) Processed abstracts zip to disk
    processed_zip_path = out_dir / "processed_abstracts.zip"
    with zipfile.ZipFile(
    processed_zip_path,
    "w",
    compression=zipfile.ZIP_DEFLATED,
    compresslevel=9,
    ) as zf:
        for f in output_dir.iterdir():
            zf.write(f, arcname=f.name)

    # 3) Reviewer packets zip to disk
    reviewer_doc_zip_path = out_dir / "reviewer_merged_packets_doc.zip"
    create_reviewer_docx_packets(assignments_df, output_dir, reviewer_doc_zip_path)

    return (
        assignments_df,
        results_df,
        str(assignments_path),
        str(processed_zip_path),
        str(reviewer_doc_zip_path),
    )

# =========================
#  STREAMLIT APP
# =========================

st.title("SRD Abstracts – Reviewer Assignment & DOCX Processor (Python)")

st.markdown(
    """
Upload your files on the left, then click **Run processing**.
The app will:
1. Match abstracts to authors and departments  
2. Assign reviewers (avoiding department conflicts)  
3. Clean the DOCX files and add abstract numbers  
4. Let you download:
   - An Excel with reviewer assignments  
   - A ZIP with all processed abstracts  
"""
)

with st.sidebar:
    st.header("1. Upload input files")

    ref_file = st.file_uploader(
        "Registrations overview (Excel)", type=["xlsx", "xls"], key="ref"
    )
    trans_file = st.file_uploader(
        "Department translations (Excel)", type=["xlsx", "xls"], key="trans"
    )
    reviewer_file = st.file_uploader(
        "Reviewers (Excel)", type=["xlsx", "xls"], key="review"
    )
    docx_files = st.file_uploader(
        "Abstract DOCX files (multiple)", type=["docx"], accept_multiple_files=True
    )

    st.header("2. Run")
    run_btn = st.button("Run processing")


if "assignments_df" not in st.session_state:
    st.session_state["assignments_df"] = None
    st.session_state["results_df"] = None
    st.session_state["assignments_path"] = None
    st.session_state["zip_path"] = None
    st.session_state["reviewer_doc_zip_path"] = None
    st.session_state["output_dir"] = None

if run_btn:
    if not (ref_file and trans_file and reviewer_file and docx_files):
        st.error("Please upload all required files (3x Excel + DOCX abstracts).")
    else:
        with st.spinner("Running pipeline... this may take a moment."):
            try:
                (
                    assignments_df,
                    results_df,
                    assignments_path,
                    zip_path,
                    reviewer_doc_zip_path,
                ) = run_pipeline(ref_file, trans_file, reviewer_file, docx_files)
                
                st.session_state["assignments_df"] = assignments_df
                st.session_state["results_df"] = results_df
                st.session_state["assignments_path"] = assignments_path
                st.session_state["zip_path"] = zip_path
                st.session_state["reviewer_doc_zip_path"] = reviewer_doc_zip_path
                st.success("Processing complete!")
            except Exception as e:
                st.error(f"Error during processing: {e}")


# Show outputs if available
if st.session_state["assignments_df"] is not None:
    st.subheader("Reviewer assignments")
    st.dataframe(st.session_state["assignments_df"], use_container_width=True)

    st.subheader("DOCX processing summary")
    st.dataframe(st.session_state["results_df"], use_container_width=True)

    col1, col2,col3 = st.columns(3)
    with col1:
        with open(st.session_state["assignments_path"], "rb") as f:
            st.download_button(
                "Download assignments (Excel)",
                data=f,
                file_name="reviewer_assignments.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    
    with col2:
        with open(st.session_state["zip_path"], "rb") as f:
            st.download_button(
                "Download processed abstracts (ZIP)",
                data=f,
                file_name="processed_abstracts.zip",
                mime="application/zip",
            )
    
    with col3:
        with open(st.session_state["reviewer_doc_zip_path"], "rb") as f:
            st.download_button(
                "Download reviewer packets (DOC ZIP)",
                data=f,
                file_name="reviewer_merged_packets_doc.zip",
                mime="application/zip",
            )
