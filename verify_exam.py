#!/usr/bin/env python3
"""
verify_exam.py

Checks that randomized exam versions are internally consistent with the original.

For each shuffled version it verifies:
  • Every section has exactly the same number of questions as the original
  • Every question (identified by its stem text) appears exactly once
  • Multiple-choice answer choices match the original set (same choices, any order)
  • "None of the above" is the last answer choice in every MC question
  • The version labels are distinct across all files
  • Question ordering actually differs from the original (catches accidental no-ops)

Usage:
    # Auto-detect _A, _B, … versions next to the original (most common)
    python3 verify_exam.py /path/to/Exam1.docx                            # Mac/Linux
    py      verify_exam.py C:\path\to\Exam1.docx                          # Windows

    # Or pass version files explicitly
    python3 verify_exam.py original.docx shuffled_A.docx shuffled_B.docx [...]

Requirements:
    Python 3.7+  (uses only the standard library plus lxml)

    Mac/Linux (recommended — use a virtual environment):
        python3 -m venv venv
        source venv/bin/activate
        pip install lxml

    Mac/Linux (quick alternative):
        pip3 install lxml --break-system-packages

    Windows:
        pip install lxml
        If that fails (e.g. missing C compiler), try:
        conda install lxml          # if you use Anaconda/Miniconda
        or download a pre-built wheel from https://pypi.org/project/lxml/
"""

import os
import re
import sys
import zipfile

from lxml import etree

# ── XML helpers ───────────────────────────────────────────────────────────────
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def WP(tag):
    return f"{{{W}}}{tag}"

def get_num_props(p):
    pPr = p.find(WP("pPr"))
    if pPr is None:
        return None, None
    numPr = pPr.find(WP("numPr"))
    if numPr is None:
        return None, None
    nid  = numPr.find(WP("numId"))
    ilvl = numPr.find(WP("ilvl"))
    return (nid.get(WP("val"))  if nid  is not None else None,
            ilvl.get(WP("val")) if ilvl is not None else None)

def get_style(p):
    pPr = p.find(WP("pPr"))
    if pPr is None:
        return None
    s = pPr.find(WP("pStyle"))
    return s.get(WP("val")) if s is not None else None

def get_text(p):
    return "".join(t.text or "" for t in p.iter(WP("t")))

def is_heading(p, level):
    return get_style(p) == f"Heading{level}"


# ── document loading ──────────────────────────────────────────────────────────

def load_body_children(docx_path):
    """Return the non-sectPr direct children of <w:body> from a .docx file."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        raw = zf.read("word/document.xml")
    body = etree.fromstring(raw).find(WP("body"))
    return [c for c in body if c.tag != WP("sectPr")]


# ── section / question extraction ─────────────────────────────────────────────

def detect_section_numIds(body_children):
    """
    Return {key: frozenset_of_numIds} for the four exam sections.
    An empty frozenset means that section is absent.

    Uses frozensets so that sections where different questions were formatted
    with different Word list styles (different numIds) are handled correctly.
    """
    heading_idx = {}
    for i, c in enumerate(body_children):
        if c.tag != WP("p"):
            continue
        if is_heading(c, 2):
            text = get_text(c).lower()
            if "true" in text:
                heading_idx["tf"]  = i
            elif "multiple" in text:
                heading_idx["mc"]  = i
            elif "fill" in text or "blank" in text:
                heading_idx["fib"] = i

    n     = len(body_children)
    mc_s  = heading_idx.get("mc",  n)
    fib_s = heading_idx.get("fib", n)
    # Default tf_s to mc_s (not 0) so paragraphs before the first detected
    # heading — title, instructions, cover content — are never included in
    # any section and are always left completely untouched.
    tf_s  = heading_idx.get("tf",  mc_s)

    # Find first FIB numId (used only for the FIB/Workout boundary)
    fib_first_id = None
    for i in range(fib_s, n):
        if body_children[i].tag != WP("p"):
            continue
        nid, ilvl = get_num_props(body_children[i])
        if nid and ilvl == "0":
            fib_first_id = nid
            break

    # Find where Workout begins
    wo_s = n
    if fib_first_id:
        for i in range(fib_s, n):
            if body_children[i].tag != WP("p"):
                continue
            nid, ilvl = get_num_props(body_children[i])
            if nid and ilvl == "0" and nid != fib_first_id:
                wo_s = i
                break

    # Collect ALL ilvl=0 numIds in each section's range
    def all_numIds(start, end):
        ids = set()
        for i in range(start, end):
            if body_children[i].tag != WP("p"):
                continue
            nid, ilvl = get_num_props(body_children[i])
            if nid and ilvl == "0":
                ids.add(nid)
        return frozenset(ids)

    return {
        "tf":  all_numIds(tf_s,  mc_s),
        "mc":  all_numIds(mc_s,  fib_s),
        "fib": all_numIds(fib_s, wo_s),
        "wo":  all_numIds(wo_s,  n),
    }


def extract_questions(body_children, q_numIds):
    """
    Return a list of question dicts for a set of numIds:
        { 'stem_text': str, 'answer_texts': [str] }
    stem_text    = plain text of the ilvl=0 paragraph
    answer_texts = plain text of each ilvl=1 paragraph (in document order)

    q_numIds is a frozenset so mixed-numId sections work correctly.
    """
    questions = []
    cur_stem, cur_answers = None, []

    for p in body_children:
        if p.tag != WP("p"):
            continue
        nid, ilvl = get_num_props(p)
        if nid in q_numIds and ilvl == "0":
            if cur_stem is not None:
                questions.append({"stem_text": cur_stem, "answer_texts": cur_answers})
            cur_stem    = get_text(p)
            cur_answers = []
        elif nid in q_numIds and ilvl == "1" and cur_stem is not None:
            cur_answers.append(get_text(p).strip())

    if cur_stem is not None:
        questions.append({"stem_text": cur_stem, "answer_texts": cur_answers})

    return questions


def get_version_label(body_children):
    """Return the first 'Version X' string found in the document."""
    pat = re.compile(r"Version\s+\S+", re.IGNORECASE)
    for c in body_children:
        if c.tag == WP("p"):
            m = pat.search(get_text(c))
            if m:
                return m.group(0)
    return "(no version label found)"


# ── verification ──────────────────────────────────────────────────────────────

SECTION_LABELS = {
    "tf":  "True/False",
    "mc":  "Multiple Choice",
    "fib": "Fill-in-the-Blank",
    "wo":  "Workout",
}


def verify_one(orig_children, orig_numIds, ver_children, filename):
    """
    Compare one shuffled version against the original.
    Prints a per-check report and returns (passes, failures, warnings).
    """
    passes = failures = warnings = 0

    def PASS(msg): nonlocal passes;   passes   += 1; print(f"    ✓  {msg}")
    def FAIL(msg): nonlocal failures; failures += 1; print(f"    ✗  {msg}")
    def WARN(msg): nonlocal warnings; warnings += 1; print(f"    ⚠  {msg}")

    ver_numIds = detect_section_numIds(ver_children)

    for key, label in SECTION_LABELS.items():
        o_ids = orig_numIds.get(key, frozenset())
        v_ids = ver_numIds.get(key, frozenset())

        if not o_ids and not v_ids:
            continue  # section absent in both — skip silently

        if not o_ids:
            WARN(f"{label}: not in original but present in version")
            continue
        if not v_ids:
            FAIL(f"{label}: present in original but MISSING from version")
            continue

        orig_qs = extract_questions(orig_children, o_ids)
        ver_qs  = extract_questions(ver_children,  v_ids)

        # Question count
        if len(orig_qs) == len(ver_qs):
            PASS(f"{label}: correct question count ({len(orig_qs)})")
        else:
            FAIL(f"{label}: count mismatch — original {len(orig_qs)}, version {len(ver_qs)}")

        # Every question present
        orig_stems = {q["stem_text"] for q in orig_qs}
        ver_stems  = {q["stem_text"] for q in ver_qs}
        missing = orig_stems - ver_stems
        extra   = ver_stems  - orig_stems
        if not missing and not extra:
            PASS(f"{label}: all question stems present")
        else:
            for s in missing:
                FAIL(f"{label}: question missing — {s[:60]!r}")
            for s in extra:
                FAIL(f"{label}: unexpected question — {s[:60]!r}")

        # Order changed
        if [q["stem_text"] for q in orig_qs] != [q["stem_text"] for q in ver_qs]:
            PASS(f"{label}: question order differs from original")
        else:
            WARN(f"{label}: question order is the SAME as original "
                 f"(expected with small sections — just verify manually)")

        # MC-specific checks
        if key == "mc":
            set_ok = nota_ok = True
            for oq in orig_qs:
                vq = next((q for q in ver_qs if q["stem_text"] == oq["stem_text"]), None)
                if vq is None:
                    continue  # already reported as missing

                if sorted(oq["answer_texts"]) != sorted(vq["answer_texts"]):
                    set_ok = False
                    FAIL(f"  MC answer-set mismatch for: {oq['stem_text'][:50]!r}")

                if any(t.strip().lower().startswith("none of") for t in oq["answer_texts"]):
                    if not vq["answer_texts"][-1].strip().lower().startswith("none of"):
                        nota_ok = False
                        FAIL(f"  'None of …' option not last for: {oq['stem_text'][:50]!r}")

            if set_ok:
                PASS("Multiple Choice: all answer-choice sets match original")
            if nota_ok:
                PASS("Multiple Choice: 'None of …' option is last in every question that has one")

    return passes, failures, warnings


# ── main ──────────────────────────────────────────────────────────────────────

def find_version_files(original_path):
    """
    Look for files named <stem>_A.docx, <stem>_B.docx, … next to original_path
    and return them in alphabetical order.  Returns an empty list if none found.
    """
    import glob
    input_dir = os.path.dirname(os.path.abspath(original_path))
    stem      = os.path.splitext(os.path.basename(original_path))[0]
    pattern   = os.path.join(input_dir, f"{stem}_*.docx")
    matches   = sorted(glob.glob(pattern))
    # Exclude the original itself in case it somehow matches
    return [p for p in matches if os.path.abspath(p) != os.path.abspath(original_path)]


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    original_path = sys.argv[1]
    version_paths = sys.argv[2:]

    # Auto-detect versions when none are given explicitly
    if not version_paths:
        version_paths = find_version_files(original_path)
        if not version_paths:
            print(f"ERROR: no version files found next to {original_path}")
            print("       Pass them explicitly or run randomize_exam.py first.")
            sys.exit(1)
        print(f"Auto-detected {len(version_paths)} version file(s):")
        for p in version_paths:
            print(f"  {os.path.basename(p)}")
        print()

    for path in [original_path] + version_paths:
        if not os.path.isfile(path):
            print(f"ERROR: file not found: {path}")
            sys.exit(1)

    print(f"Loading original:  {original_path}")
    orig_children = load_body_children(original_path)
    orig_numIds   = detect_section_numIds(orig_children)

    print(f"  Version label:    {get_version_label(orig_children)}")
    print(f"  Sections found:   " +
          ", ".join(SECTION_LABELS[k] for k in ("tf", "mc", "fib", "wo")
                    if orig_numIds.get(k)))
    print()

    total_p = total_f = total_w = 0
    version_labels = []

    for path in version_paths:
        print(f"Verifying: {os.path.basename(path)}")
        children = load_body_children(path)
        label    = get_version_label(children)
        version_labels.append(label)
        print(f"  Version label:  {label}")

        p, f, w = verify_one(orig_children, orig_numIds, children,
                              os.path.basename(path))
        total_p += p
        total_f += f
        total_w += w
        print()

    # Version label uniqueness
    print("Version label uniqueness:")
    seen = {}
    all_unique = True
    for path, lbl in zip(version_paths, version_labels):
        name = os.path.basename(path)
        if lbl in seen:
            print(f"  ✗  DUPLICATE label '{lbl}' in {name} and {seen[lbl]}")
            total_f += 1
            all_unique = False
        else:
            seen[lbl] = name
    if all_unique:
        print(f"  ✓  All {len(version_labels)} version labels are unique")
        total_p += 1

    print()
    print("─" * 50)
    print(f"  PASSED:   {total_p}")
    print(f"  WARNINGS: {total_w}")
    print(f"  FAILED:   {total_f}")
    print("─" * 50)

    if total_f > 0:
        print("One or more checks FAILED — review output above.")
        sys.exit(1)
    elif total_w > 0:
        print("All checks passed (with warnings — review above).")
    else:
        print("All checks passed.")


if __name__ == "__main__":
    main()
