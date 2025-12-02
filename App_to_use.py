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

# =========================
#  YOUR HELPER FUNCTIONS
# =========================

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


def find_first_marker(doc):
    marker_order = [
        "choose your research type",
        "clinical",
        "fundamental",
    ]

    for marker in marker_order:
        for i, p in enumerate(doc.paragraphs):
            if marker in p.text.lower():
                return i

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
    for idx, row in df1.iterrows():
        name = row[key1]
        match = process.extractOne(name, df2[key2].tolist(), scorer=scorer)
        if match and match[1] >= threshold:
            matched_name = match[0]
            matched_row = df2[df2[key2] == matched_name].iloc[0]
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

    separators = r"[;,+/]+"

    # reference registrations
    ref_df = pd.read_excel(ref_file)
    ref_df.columns = ref_df.columns.str.lower().str.replace(" ", "_")
    ref_df = ref_df.apply(lambda col: col.astype(str).str.strip())

    # Expand multiple departments
    ref_df["department_"] = ref_df["department_"].fillna("")

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
    # 1) Extract raw text lines from the docx (for parsing)
    txt = extract_docx_text_with_superscripts(filepath)

    # Helper to find a line index containing a phrase
    def find_line(phrase):
        for i, line in enumerate(txt):
            if phrase.lower() in line.lower():
                return i
        return None

    author_line = find_line("author")
    aff_line = find_line("affiliations")

    # If no affiliations, skip
    if aff_line is None or author_line is None:
        return {
            "file": Path(filepath).name,
            "name": None,
            "matched_name": None,
            "abstract_nr": None,
            "matched": False,
        }

    # 2) Extract first author from the "Author(s)" line
    def extract_first_author(txt, author_line_index):
        """
        Extracts the first author name.
        If the author name is not on the same line as 'Author(s)',
        it checks the next lines until a non-empty valid name is found.
        """
    
        # 1) Try same line first
        line = txt[author_line_index]
        line_clean = (
            line.replace("Authors:", "")
                .replace("Author:", "")
                .replace("Authors", "")
                .replace("Author", "")
                .strip()
        )
    
        # If the same line contains a name like "L.A.E.M. van Houtum¹"
        if line_clean and not line_clean.lower().startswith("affiliation"):
            # First token until comma if multiple authors
            first = line_clean.split(",")[0].strip()
            first = re.sub(r"\d+$", "", first).strip()  # remove footnote digits
            if len(first) > 0:
                return first
    
        # 2) Otherwise → check following lines until we find a name
        i = author_line_index + 1
        while i < len(txt):
            candidate = txt[i].strip()
            # skip blank lines
            if not candidate:
                i += 1
                continue
            # skip lines like "Department", "Affiliations"
            if "affiliation" in candidate.lower():
                break
            # extract only the first author before commas
            first = candidate.split(",")[0].strip()
            first = re.sub(r"\d+$", "", first).strip()
            if len(first) > 0:
                return first
            i += 1
    
        return ""  # failed (rare)
    name = extract_first_author(txt, author_line)
    name = clean_name(name)

    # 3) Match this name to ref_df["name"] using your fuzzy matcher
    matched_name = match_author_name(name, ref_df, "name")

    # 4) Get abstract_nr for this matched name (if present)
# Assign one unique abstract number per matched author
    abstract_nr = None
    
    if matched_name:
        key = matched_name.strip()
    
        if key in remaining_ids and len(remaining_ids[key]) > 0:
            abstract_nr = remaining_ids[key].pop(0)   # TAKE & REMOVE ONE NUMBER

    # 5) Now open the original doc as a Document to modify it
    doc = Document(filepath)

    # Find "choose your research type" paragraph index
    research_idx = find_first_marker(doc)

    # If not found, just save a copy without changes (but still return metadata)
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

    # 6) Remove all paragraphs before that index
    for _ in range(research_idx):
        p = doc.paragraphs[0]
        p._element.getparent().remove(p._element)

    # 7) Insert 5 blank lines + abstract number at the top, keeping formatting
    body = doc._element.body

    # Insert "Abstract number: X" paragraph (if we have a number)
    if abstract_nr is not None:
        p_label = OxmlElement("w:p")
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = f"Abstract number: {abstract_nr}"
        r.append(t)
        p_label.append(r)
        body.insert(0, p_label)

    # Insert 5 completely empty paragraphs *above* that
    for _ in range(5):
        new_p = OxmlElement("w:p")
        body.insert(0, new_p)

    # 8) Save modified document
    out_path = Path(output_folder) / (Path(filepath).stem + "_after_research_type.docx")
    doc.save(out_path)

    # 9) Return summary info for this file
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
    assignments_buffer = BytesIO()
    assignments_df.to_excel(assignments_buffer, index=False)
    assignments_buffer.seek(0)

    # Build ZIP bytes with processed docs
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for f in output_dir.iterdir():
            zf.write(f, arcname=f.name)
    zip_buffer.seek(0)

    return assignments_df, results_df, assignments_buffer, zip_buffer


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
    st.session_state["assignments_bytes"] = None
    st.session_state["zip_bytes"] = None

if run_btn:
    if not (ref_file and trans_file and reviewer_file and docx_files):
        st.error("Please upload all required files (3x Excel + DOCX abstracts).")
    else:
        with st.spinner("Running pipeline... this may take a moment."):
            try:
                (
                    assignments_df,
                    results_df,
                    assignments_bytes,
                    zip_bytes,
                ) = run_pipeline(ref_file, trans_file, reviewer_file, docx_files)

                st.session_state["assignments_df"] = assignments_df
                st.session_state["results_df"] = results_df
                st.session_state["assignments_bytes"] = assignments_bytes
                st.session_state["zip_bytes"] = zip_bytes

                st.success("Processing complete!")
            except Exception as e:
                st.error(f"Error during processing: {e}")


# Show outputs if available
if st.session_state["assignments_df"] is not None:
    st.subheader("Reviewer assignments")
    st.dataframe(st.session_state["assignments_df"], use_container_width=True)

    st.subheader("DOCX processing summary")
    st.dataframe(st.session_state["results_df"], use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "Download assignments (Excel)",
            data=st.session_state["assignments_bytes"],
            file_name="reviewer_assignments.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with col2:
        st.download_button(
            "Download processed abstracts (ZIP)",
            data=st.session_state["zip_bytes"],
            file_name="processed_abstracts.zip",
            mime="application/zip",
        )
