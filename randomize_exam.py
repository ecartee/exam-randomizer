#!/usr/bin/env python3
"""
randomize_exam.py

Creates N randomized versions of a Word exam document (.docx).

What gets shuffled:
  • True/False questions    — question order only (True/False options stay put)
  • Multiple Choice         — question order + answer choices
                              ("None of the above" is always kept last)
  • Fill-in-the-Blank       — question order only (sub-parts a, b, c… stay with their question)
  • Workout / free-response — question order only

The title line (e.g. "Version A") is automatically updated for each output file.
All output files contain the exact same questions; only the order differs.

Usage:
    # Two versions (default) — outputs written next to the input file
    python3 randomize_exam.py /path/to/Exam1.docx                 # Mac/Linux
    py      randomize_exam.py C:\\path\\to\\Exam1.docx              # Windows

    # Four versions next to the input file
    python3 randomize_exam.py /path/to/Exam1.docx --versions 4

    # Explicit output paths (optional, for placing files elsewhere)
    python3 randomize_exam.py input.docx vA.docx vB.docx vC.docx

    # Fix the seed for reproducibility (each version gets base_seed + index)
    python3 randomize_exam.py /path/to/Exam1.docx --base-seed 99

Options:
    --base-seed N   Pin the seed for the first version (subsequent versions use
                    base_seed+1, base_seed+2, …).  By default, seeds are drawn
                    from the OS random-number generator and printed to stdout so
                    you can record them and re-run with --base-seed if needed.

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

NOTE on floating shapes (fill-in-the-blank blanks):
    If your FIB blanks are drawn as floating Word shapes (inserted via
    Insert → Shapes), they are anchored to a specific paragraph and move with
    it.  The horizontal position is typically relative to the column (fixed),
    while the vertical position is relative to the anchor paragraph (moves).
    In practice this looks fine after shuffling, but always open each output
    file in Word and scroll through the FIB section before printing to
    confirm the blanks are positioned correctly.
"""

import argparse
import copy
import os
import random
import re
import secrets
import sys
import zipfile

from lxml import etree

# ── XML namespace shortcut ────────────────────────────────────────────────────
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def WP(tag):
    return f"{{{W}}}{tag}"

# ── paragraph helpers ─────────────────────────────────────────────────────────

def get_num_props(p):
    """Return (numId_str, ilvl_str) for a <w:p>, or (None, None)."""
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
    """Return the w:pStyle value for a paragraph, or None."""
    pPr = p.find(WP("pPr"))
    if pPr is None:
        return None
    style = pPr.find(WP("pStyle"))
    return style.get(WP("val")) if style is not None else None

def get_text(p):
    """Concatenate all <w:t> text inside a paragraph."""
    return "".join(t.text or "" for t in p.iter(WP("t")))

def is_heading(p, level):
    return get_style(p) == f"Heading{level}"


# ── core shuffling logic ──────────────────────────────────────────────────────

def group_questions(paras, q_numIds):
    """
    Split a flat list of paragraphs (one section) into:
        pre    — paragraphs that appear before the first question
        blocks — list of question-blocks, each a list of paragraphs:
                 [question_stem, *answer_choices_or_sub_parts, *trailing_spacers]

    q_numIds is a frozenset of numId strings; any ilvl=0 paragraph whose numId
    is in that set is treated as a question stem.  This handles exams where
    different questions in the same section were formatted with different list
    styles (and therefore ended up with different numIds in the XML).
    """
    pre = []
    blocks = []
    current = None

    for p in paras:
        nid, ilvl = get_num_props(p)
        if nid in q_numIds and ilvl == "0":
            if current is not None:
                blocks.append(current)
            current = [p]
        elif current is not None:
            current.append(p)
        else:
            pre.append(p)

    if current is not None:
        blocks.append(current)

    return pre, blocks


def is_none_option(text):
    """
    Return True if an answer choice is a catch-all 'None of …' option that
    should always be pinned last.  Matches 'None of the above', 'None of
    these', 'None of the above answers', etc.
    """
    return text.strip().lower().startswith("none of")


def shuffle_mc_answers(block, q_numIds, rng):
    """
    Shuffle the ilvl=1 answer-choice paragraphs within one MC question block.
    Any answer whose plain text starts with "none of" (e.g. "None of the
    above", "None of these") is pinned last.
    Returns a new list (does not modify in place).

    q_numIds is a frozenset so we handle mixed-numId sections correctly.
    """
    answer_idx = [
        i for i, p in enumerate(block)
        if get_num_props(p)[0] in q_numIds and get_num_props(p)[1] == "1"
    ]
    if not answer_idx:
        return block

    answers = [block[i] for i in answer_idx]

    nota = None
    rest = []
    for p in answers:
        if is_none_option(get_text(p)):
            nota = p
        else:
            rest.append(p)

    rng.shuffle(rest)
    shuffled = rest + ([nota] if nota is not None else [])

    new_block = list(block)
    for idx, p in zip(answer_idx, shuffled):
        new_block[idx] = p
    return new_block


def ensure_pPr(p):
    """Return the <w:pPr> child of p, creating it as the first child if absent."""
    pPr = p.find(WP("pPr"))
    if pPr is None:
        pPr = etree.Element(WP("pPr"))
        p.insert(0, pPr)
    return pPr


def is_page_break_only(p):
    """
    Return True if p is a paragraph whose sole purpose is a manual page break
    (i.e. it contains a <w:br w:type="page"/> run and no visible text).
    """
    if p.tag != WP("p"):
        return False
    if any(t.text and t.text.strip() for t in p.iter(WP("t"))):
        return False
    return any(br.get(WP("type")) == "page" for br in p.iter(WP("br")))


def strip_page_breaks(block):
    """
    Remove explicit page breaks from a list of paragraphs so Word can reflow
    text naturally after shuffling.

      • Drops paragraphs that are pure page-break runs (<w:br w:type="page"/>).
      • Removes <w:pageBreakBefore/> from paragraph properties (this property
        forces a paragraph to always start on a new page).
      • Removes inline <w:br w:type="page"/> elements embedded inside runs
        (these can appear when a page break is inserted mid-paragraph).

    Soft line returns (Shift+Enter) within question content are intentional
    spacing chosen by the exam author and are left untouched.

    Section-heading paragraphs (in the 'pre' list, not inside any block) are
    left untouched, so intentional section-level breaks are preserved.
    """
    result = []
    for p in block:
        if is_page_break_only(p):
            continue
        pPr = p.find(WP("pPr"))
        if pPr is not None:
            pb = pPr.find(WP("pageBreakBefore"))
            if pb is not None:
                pPr.remove(pb)
        for run in p.findall(WP("r")):
            for br in list(run.findall(WP("br"))):
                if br.get(WP("type")) == "page":
                    run.remove(br)
        result.append(p)
    return result


def is_blank_paragraph(p):
    """Return True if p is an empty or whitespace-only paragraph (no visible text)."""
    if p.tag != WP("p"):
        return False
    return not any(t.text and t.text.strip() for t in p.iter(WP("t")))


def strip_trailing_spacers(block):
    """
    Remove blank paragraphs from the *end* of a question block only.

    Blank paragraphs in the interior of a block (e.g. between a question
    stem and its first answer choice, or between answer choices) are left
    untouched — they are intentional within-question formatting.

    Only trailing blanks — those that appear after the last list item and
    before the next question stem — are stripped.  A single uniform blank
    paragraph is then inserted between every pair of questions by the
    caller, replacing the source document's ad-hoc between-question spacing.

    List-item paragraphs (those with a numId) are never stripped even if they
    have no visible text — blank list items are structural parts of the question
    (e.g. fill-in-the-blank answer lines).
    """
    while block and is_blank_paragraph(block[-1]) and get_num_props(block[-1])[0] is None:
        block = block[:-1]
    return block


def strip_trailing_soft_returns(block):
    """
    Strip soft-return <w:br> elements that trail after the last visible text
    in the last list-item paragraph of a block.

    The last answer choice in a question often has trailing Shift+Enter breaks
    added for visual spacing (e.g. 'False\\n\\n').  These are between-question
    spacing, not within-question content, so they are removed here.  Soft
    returns that appear before or between visible text in the same paragraph
    are untouched.
    """
    last_list_para = next(
        (p for p in reversed(block) if p.tag == WP("p") and get_num_props(p)[0] is not None),
        None,
    )
    if last_list_para is None:
        return block

    # Find the last <w:t> element with visible text, then remove all soft-return
    # <w:br> elements that follow it in document order.
    all_desc   = list(last_list_para.iter())
    last_text_idx = max(
        (i for i, el in enumerate(all_desc)
         if el.tag == WP("t") and el.text and el.text.strip()),
        default=-1,
    )
    if last_text_idx == -1:
        return block

    for i, el in enumerate(all_desc):
        if i > last_text_idx and el.tag == WP("br"):
            if el.get(WP("type")) not in ("page", "column"):
                el.getparent().remove(el)

    return block


def add_keep_together(block):
    """
    Add <w:keepNext/> and <w:keepLines/> to every paragraph in the block
    except the last.

    keepNext  — keeps each paragraph on the same page as the one that follows
                it, so the whole question travels together.
    keepLines — keeps all lines of an individual paragraph together, preventing
                a long question stem from being split mid-paragraph.
    """
    last_p_idx = max((i for i, p in enumerate(block) if p.tag == WP("p")), default=None)
    for i, p in enumerate(block):
        if p.tag != WP("p") or i == last_p_idx:
            continue
        pPr = ensure_pPr(p)
        if pPr.find(WP("keepNext")) is None:
            etree.SubElement(pPr, WP("keepNext"))
        if pPr.find(WP("keepLines")) is None:
            etree.SubElement(pPr, WP("keepLines"))
    return block


def update_version_label(body, new_label):
    """
    Replace the first occurrence of 'Version <letter>' with new_label.

    Word often splits a heading into many small <w:t> runs (one per formatted
    span), so 'Version ' and the letter may live in adjacent nodes rather than
    a single one.  This function handles both cases:

      1. Simple — the whole 'Version X' token is inside one <w:t>.
      2. Split  — 'Version ' ends one <w:t> and the letter starts the next
                  (possibly with leading whitespace).
    """
    pattern    = re.compile(r"Version\s+[A-Z]", re.IGNORECASE)
    new_letter = new_label.split()[-1]          # "A" from "Version A"

    for p in body.iter(WP("p")):
        if not pattern.search(get_text(p)):
            continue

        t_nodes = [t for t in p.iter(WP("t")) if t.text is not None]

        # Case 1: entire match in one node
        for t in t_nodes:
            if pattern.search(t.text):
                t.text = pattern.sub(new_label, t.text)
                return

        # Case 2: "Version " ends one node, letter begins the next
        for i, t in enumerate(t_nodes):
            if not re.search(r"Version\s*$", t.text, re.IGNORECASE):
                continue
            for j in range(i + 1, len(t_nodes)):
                ahead = t_nodes[j].text
                m = re.match(r"(\s*)([A-Z])(.*)", ahead, re.IGNORECASE)
                if m:
                    t_nodes[j].text = m.group(1) + new_letter + m.group(3)
                    return
                if ahead.strip():   # non-whitespace, non-letter — give up
                    break
        return  # paragraph matched full text but structure was unexpected


# ── section detection ─────────────────────────────────────────────────────────

def detect_sections(body_children):
    """
    Identify the four shuffleable sections and return a list of section dicts:
        {
          'label':            str,          human-readable name
          'paras':            [p],          all paragraphs in this section
          'q_numIds':         frozenset,    numIds of question-stem paragraphs
          'shuffle_answers':  bool,         True only for Multiple Choice
        }

    q_numIds is a *set* (frozenset) rather than a single value so that the
    script handles exams where different questions in the same section were
    formatted with different Word list styles — a common result of
    copy-pasting from multiple sources.

    Sections are located as follows:
        True/False      — Heading2 whose text contains "true"
        Multiple Choice — Heading2 whose text contains "multiple"
        Fill in Blank   — Heading2 whose text contains "fill" or "blank"
        Workout         — no Heading2; detected as the first ilvl=0 numId
                          after FIB that doesn't appear in the FIB numId set
    """
    # ── Pass 1: find Heading2 positions by keyword ────────────────────────────
    heading_idx = {}
    for i, child in enumerate(body_children):
        if child.tag != WP("p"):
            continue
        if is_heading(child, 2):
            text = get_text(child).lower()
            if "true" in text:
                heading_idx["tf"]  = i
            elif "multiple" in text:
                heading_idx["mc"]  = i
            elif "fill" in text or "blank" in text:
                heading_idx["fib"] = i

    n     = len(body_children)
    mc_s  = heading_idx.get("mc",  n)
    fib_s = heading_idx.get("fib", n)
    # Default tf_s to mc_s (not 0) so that paragraphs before the first
    # detected heading — title, instructions, cover-page content — are never
    # included in any section and are always passed through untouched.
    tf_s  = heading_idx.get("tf",  mc_s)

    # ── Pass 2: find the *first* ilvl=0 numId in FIB — used only to locate
    #            the FIB/Workout boundary, not as the sole question identifier
    def first_q_numId(start, end):
        for i in range(start, end):
            if body_children[i].tag != WP("p"):
                continue
            nid, ilvl = get_num_props(body_children[i])
            if nid is not None and ilvl == "0":
                return nid
        return None

    fib_first_id = first_q_numId(fib_s, n)

    # ── Pass 3: find where Workout begins.
    #
    # Step A: find the first ilvl=0 list item after FIB whose numId differs
    #         from the FIB numId — that's the first Workout question stem.
    # Step B: scan *backward* from that list item looking for a Heading 2
    #         that immediately precedes it (with only blank paragraphs in
    #         between).  If found, start the Workout section at the heading
    #         so that it lands in the Workout section's 'pre' (preamble) and
    #         gets the page-break treatment rather than being glued to the
    #         last FIB question block.
    wo_s = n
    if fib_first_id is not None:
        first_wo_list = n
        for i in range(fib_s, n):
            if body_children[i].tag != WP("p"):
                continue
            nid, ilvl = get_num_props(body_children[i])
            if nid is not None and ilvl == "0" and nid != fib_first_id:
                first_wo_list = i
                break

        wo_s = first_wo_list   # default: start at the first list item

        # Walk backward to find an immediately preceding Heading 2
        if first_wo_list < n:
            for i in range(first_wo_list - 1, fib_s - 1, -1):
                if body_children[i].tag != WP("p"):
                    continue
                if is_heading(body_children[i], 2):
                    wo_s = i   # promote: section starts at the heading
                    break
                # Stop if we hit a list item (FIB content) or non-blank text
                nid, _ = get_num_props(body_children[i])
                if nid is not None:
                    break
                if any(t.text and t.text.strip()
                       for t in body_children[i].iter(WP("t"))):
                    break

    # ── Pass 4: collect ALL ilvl=0 numIds within each section's actual range
    def all_numIds(start, end):
        ids = set()
        for i in range(start, end):
            if body_children[i].tag != WP("p"):
                continue
            nid, ilvl = get_num_props(body_children[i])
            if nid is not None and ilvl == "0":
                ids.add(nid)
        return frozenset(ids)

    tf_ids  = all_numIds(tf_s,  mc_s)
    mc_ids  = all_numIds(mc_s,  fib_s)
    fib_ids = all_numIds(fib_s, wo_s)   # only the FIB range, not Workout
    wo_ids  = all_numIds(wo_s,  n)

    # ── Move the workout boundary back to just after the last FIB list item.
    #    Paragraphs between the last FIB question and the first workout question
    #    (heading, instructions, blank lines) belong to the workout preamble.
    if wo_ids and fib_ids:
        last_fib_item = fib_s
        for i in range(fib_s, wo_s):
            if body_children[i].tag == WP("p"):
                nid, _ = get_num_props(body_children[i])
                if nid in fib_ids:
                    last_fib_item = i
        if last_fib_item + 1 < wo_s:
            wo_s = last_fib_item + 1

    # ── Find the first "INTENTIONALLY LEFT BLANK" paragraph.
    #    This acts as a hard stop for ALL sections — those pages must never be
    #    included in any question block, regardless of which section is last.
    questions_end = n
    for i in range(tf_s, n):
        if body_children[i].tag == WP("p"):
            if "INTENTIONALLY LEFT BLANK" in get_text(body_children[i]).upper():
                questions_end = i
                break

    # ── Build section list
    sections = []
    if tf_ids:
        sections.append({"label": "True/False",      "paras": body_children[tf_s:min(mc_s,  questions_end)], "q_numIds": tf_ids,  "shuffle_answers": False})
    if mc_ids:
        sections.append({"label": "Multiple Choice", "paras": body_children[mc_s:min(fib_s, questions_end)], "q_numIds": mc_ids,  "shuffle_answers": True})
    if fib_ids:
        sections.append({"label": "Fill in Blank",   "paras": body_children[fib_s:min(wo_s, questions_end)], "q_numIds": fib_ids, "shuffle_answers": False})
    if wo_ids:
        sections.append({"label": "Workout",         "paras": body_children[wo_s:questions_end],              "q_numIds": wo_ids,  "shuffle_answers": False})
    return sections


# ── document-level shuffle ────────────────────────────────────────────────────

def build_shuffled_body_children(body_children, sections, rng):
    """
    Return a new ordered list of body children with questions shuffled within
    each section.  Paragraphs outside any shuffleable section are unchanged.

    For each section:
      • Manual page breaks are stripped from question content so Word can reflow.
      • <w:keepNext/> and <w:keepLines/> are added to all but the last paragraph
        of each block, preventing questions from being split across page breaks.
      • A <w:pageBreakBefore/> is added to the first paragraph of every section
        after the first, ensuring a clean page break between section types.
      • Workout questions each get their own page (pageBreakBefore on every
        question after the first in that section).
    """
    section_replacements = {}
    for sec_idx, sec in enumerate(sections):
        pre, blocks = group_questions(sec["paras"], sec["q_numIds"])
        blocks = [list(b) for b in blocks]   # shallow copies for safety
        pre    = strip_page_breaks(pre)       # remove source-doc breaks from preamble too
        blocks = [strip_page_breaks(b) for b in blocks]
        blocks = [strip_trailing_spacers(b) for b in blocks]
        blocks = [strip_trailing_soft_returns(b) for b in blocks]
        rng.shuffle(blocks)
        if sec["shuffle_answers"]:
            blocks = [shuffle_mc_answers(b, sec["q_numIds"], rng) for b in blocks]
        # add_keep_together runs last so keepNext/keepLines reflect the final
        # paragraph order within each block (after any answer shuffling)
        blocks = [add_keep_together(b) for b in blocks]

        # Flatten with a single blank paragraph between every pair of questions
        # (source-doc spacers have been stripped; spacing is now uniform)
        flat_blocks = []
        for i, block in enumerate(blocks):
            flat_blocks.extend(block)
            if i < len(blocks) - 1:
                flat_blocks.append(etree.Element(WP("p")))  # one blank line separator
        flat = pre + flat_blocks

        if sec["label"] == "Workout":
            # Workout page breaks are handled explicitly rather than via the
            # generic section logic, so we can target the heading precisely.
            #
            # If there is a preamble (heading / instructions before the first
            # question), the page break goes on the first paragraph of that
            # preamble so the heading and first question share a page.
            # If there is no preamble, the page break goes on the first question.
            if pre:
                first_pre_p = next((p for p in pre if p.tag == WP("p")), None)
                if first_pre_p is not None:
                    pPr = ensure_pPr(first_pre_p)
                    if pPr.find(WP("pageBreakBefore")) is None:
                        etree.SubElement(pPr, WP("pageBreakBefore"))
                # Every question after the first gets its own page
                for block in blocks[1:]:
                    first_p = next((p for p in block if p.tag == WP("p")), None)
                    if first_p is not None:
                        pPr = ensure_pPr(first_p)
                        if pPr.find(WP("pageBreakBefore")) is None:
                            etree.SubElement(pPr, WP("pageBreakBefore"))
                # Safety net: hard-remove any pageBreakBefore from the first
                # workout question so it is always on the same page as the
                # heading above it, regardless of what the source document held.
                if blocks:
                    first_q_p = next((p for p in blocks[0] if p.tag == WP("p")), None)
                    if first_q_p is not None:
                        pPr_q = first_q_p.find(WP("pPr"))
                        if pPr_q is not None:
                            pb = pPr_q.find(WP("pageBreakBefore"))
                            if pb is not None:
                                pPr_q.remove(pb)
            else:
                # No heading detected — every question gets its own page
                for block in blocks:
                    first_p = next((p for p in block if p.tag == WP("p")), None)
                    if first_p is not None:
                        pPr = ensure_pPr(first_p)
                        if pPr.find(WP("pageBreakBefore")) is None:
                            etree.SubElement(pPr, WP("pageBreakBefore"))

        # Every non-Workout section after the first starts on a fresh page
        elif sec_idx > 0 and flat:
            first_p = next((p for p in flat if p.tag == WP("p")), None)
            if first_p is not None:
                pPr = ensure_pPr(first_p)
                if pPr.find(WP("pageBreakBefore")) is None:
                    etree.SubElement(pPr, WP("pageBreakBefore"))

        if sec["paras"]:
            section_replacements[id(sec["paras"][0])] = flat

    sec_para_ids   = {id(p): sec for sec in sections for p in sec["paras"]}
    emitted_secs   = set()
    result         = []

    for child in body_children:
        cid = id(child)
        if cid not in sec_para_ids:
            result.append(child)
        else:
            sec    = sec_para_ids[cid]
            sec_id = id(sec["paras"][0]) if sec["paras"] else None
            if sec_id not in emitted_secs:
                emitted_secs.add(sec_id)
                result.extend(section_replacements.get(sec_id, sec["paras"]))
            # else: already emitted — skip

    return result


# ── docx read / write ─────────────────────────────────────────────────────────

DOCUMENT_XML = "word/document.xml"

def read_document_xml(docx_path):
    """Read and parse word/document.xml from a .docx zip."""
    with zipfile.ZipFile(docx_path, "r") as zf:
        raw = zf.read(DOCUMENT_XML)
    return etree.fromstring(raw)


def write_docx(input_docx, output_docx, new_root):
    """
    Copy every file from input_docx into output_docx, replacing
    word/document.xml with the serialised new_root element.
    """
    new_xml = etree.tostring(new_root, xml_declaration=True,
                             encoding="UTF-8", pretty_print=False)
    with zipfile.ZipFile(input_docx, "r") as zin, \
         zipfile.ZipFile(output_docx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == DOCUMENT_XML:
                zout.writestr(item, new_xml)
            else:
                zout.writestr(item, zin.read(item.filename))


# ── per-version entry point ───────────────────────────────────────────────────

def make_version(input_docx, output_docx, version_label, seed):
    """Parse, shuffle, and write one randomized version."""

    root = read_document_xml(input_docx)
    body = root.find(WP("body"))

    # Separate <w:sectPr> (page layout) — must stay as the last body element
    sectPr     = None
    non_sectPr = []
    for child in list(body):
        if child.tag == WP("sectPr"):
            sectPr = child
        else:
            non_sectPr.append(child)

    sections = detect_sections(non_sectPr)
    if not sections:
        print(f"  WARNING: no shuffleable sections detected — {input_docx}")

    # Deep-copy so the same in-memory tree can produce multiple independent versions
    copied   = [copy.deepcopy(c) for c in non_sectPr]
    id_map   = {id(orig): cp for orig, cp in zip(non_sectPr, copied)}
    secs_copy = [dict(sec, paras=[id_map[id(p)] for p in sec["paras"]])
                 for sec in sections]

    rng      = random.Random(seed)
    shuffled = build_shuffled_body_children(copied, secs_copy, rng)

    # Rebuild body
    for child in list(body):
        body.remove(child)
    for child in shuffled:
        body.append(child)
    if sectPr is not None:
        body.append(copy.deepcopy(sectPr))

    # Ensure every "INTENTIONALLY LEFT BLANK" page starts on a new page
    for child in list(body):
        if child.tag == WP("p") and "INTENTIONALLY LEFT BLANK" in get_text(child).upper():
            pPr = ensure_pPr(child)
            if pPr.find(WP("pageBreakBefore")) is None:
                etree.SubElement(pPr, WP("pageBreakBefore"))

    update_version_label(body, version_label)
    write_docx(input_docx, output_docx, root)
    print(f"  ✓ {output_docx}")


# ── main ──────────────────────────────────────────────────────────────────────

def auto_output_paths(input_path, n):
    """
    Generate n output paths in the same directory as input_path, named
    <stem>_A.docx, <stem>_B.docx, … (or <stem>_1.docx, … beyond Z).
    """
    input_dir = os.path.dirname(os.path.abspath(input_path))
    stem      = os.path.splitext(os.path.basename(input_path))[0]
    def label(i):
        return chr(ord('A') + i) if i < 26 else str(i + 1)
    return [os.path.join(input_dir, f"{stem}_{label(i)}.docx") for i in range(n)]


def main():
    parser = argparse.ArgumentParser(
        description="Create randomized versions of a Word exam (.docx).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("input", help="Source exam .docx file")
    parser.add_argument(
        "outputs", nargs="*",
        help="Output .docx path for each version (A, B, C, …). "
             "If omitted, versions are written to the same folder as the input, "
             "named <stem>_A.docx, <stem>_B.docx, etc.",
    )
    parser.add_argument(
        "--versions", type=int, default=2, metavar="N",
        help="Number of versions to generate when output paths are not given "
             "explicitly (default: 2). Ignored if output paths are provided.",
    )
    parser.add_argument(
        "--base-seed", type=int, default=None, metavar="N",
        help="Pin the seed for the first version (subsequent versions use base_seed+1, etc.). "
             "Omit to use a fresh random seed each run.",
    )
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"ERROR: file not found: {args.input}")
        sys.exit(1)

    # Resolve output paths — explicit if provided, otherwise auto-generated
    if args.outputs:
        output_paths = args.outputs
    else:
        output_paths = auto_output_paths(args.input, args.versions)

    def version_label(i):
        return f"Version {chr(ord('A') + i)}" if i < 26 else f"Version {i + 1}"

    n = len(output_paths)

    # Determine seeds
    if args.base_seed is not None:
        seeds = [args.base_seed + i for i in range(n)]
        seed_note = f"(fixed --base-seed {args.base_seed})"
    else:
        # Draw one base seed from the OS RNG, then use consecutive offsets so
        # that --base-seed <seeds[0]> reproduces the entire run exactly.
        # Capped at 2**31-1 so the value fits comfortably as a CLI argument.
        base = secrets.randbelow(2**31 - n)
        seeds = [base + i for i in range(n)]
        seed_note = "(randomly generated — record these to reproduce later)"

    # Detect sections once from the input so we can tailor the closing tip
    # and run pre-flight checks.
    input_root     = read_document_xml(args.input)
    input_body     = input_root.find(WP("body"))
    input_children = [c for c in input_body if c.tag != WP("sectPr")]
    sections       = detect_sections(input_children)
    has_fib        = any(s["label"] == "Fill in Blank" for s in sections)

    # ── Pre-flight: warn about duplicate question stems within any section ────
    for sec in sections:
        _, blocks = group_questions(sec["paras"], sec["q_numIds"])
        stems = [get_text(block[0]) for block in blocks if block]
        seen, dupes = set(), []
        for stem in stems:
            if stem in seen and stem not in dupes:
                dupes.append(stem)
            seen.add(stem)
        for stem in dupes:
            print(f"WARNING: duplicate question stem in {sec['label']} — "
                  f"{stem[:70]!r}")

    print(f"Generating {n} version{'s' if n != 1 else ''} from: {args.input}")
    print(f"Seeds {seed_note}:")
    for i, (out_path, seed) in enumerate(zip(output_paths, seeds)):
        print(f"  {version_label(i)}: seed={seed}  →  {out_path}")
    print()

    for i, (out_path, seed) in enumerate(zip(output_paths, seeds)):
        label = version_label(i)
        print(f"  Creating {label} …")
        make_version(args.input, out_path, label, seed)

    print(f"\nDone — {n} file{'s' if n != 1 else ''} written.")
    if args.base_seed is None:
        print("TIP: to regenerate these exact versions, re-run with "
              f"--base-seed {seeds[0]}")
    if has_fib:
        print("Open each in Word and scroll through the fill-in-the-blank section "
              "to confirm blank positions.")


if __name__ == "__main__":
    main()
