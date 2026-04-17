# Exam Randomizer

Two Python scripts for generating multiple randomized versions of a Word-based exam, and verifying that those versions are correct.

Designed for exams written in `.docx` format (rather than LaTeX) so that accessibility features like screen-reader compatibility are preserved.

---

## What it does

`randomize_exam.py` takes a single source exam and produces N versions with shuffled question order. For multiple-choice questions it also shuffles the answer choices. Every version contains the exact same questions at the same point values — only the order differs.

`verify_exam.py` checks one or more randomized versions against the original and reports any problems: missing questions, lost answer choices, a misplaced "None of the above", duplicate version labels, etc.

### What gets shuffled

| Section | Question order | Answer choices |
|---|---|---|
| True/False | ✓ | — (always just True / False) |
| Multiple Choice | ✓ | ✓ ("None of the above" always stays last) |
| Fill-in-the-Blank | ✓ | — (sub-parts a, b, c… stay with their question) |
| Workout / free-response | ✓ | — |

Sections are detected automatically from the document structure, so an exam with only some of these sections (e.g. multiple choice only) works without any configuration.

---

## Requirements

Python 3.7+ and `lxml`. No other dependencies.

**Mac / Linux**

Modern macOS (and some Linux distributions) block `pip` from installing into the system Python. The recommended approach is a virtual environment, which you create once and reuse:

```bash
# Run these once, from the exam_randomizer folder
python3 -m venv venv
source venv/bin/activate
pip install lxml
```

Each time you open a new terminal session, re-activate the environment before running the scripts:

```bash
source venv/bin/activate
```

Alternatively, if you just want a quick install without a virtual environment:

```bash
pip3 install lxml --break-system-packages
```

**Windows**
```bash
pip install lxml
```
If that fails (e.g. missing C compiler), try `conda install lxml` (Anaconda/Miniconda), or download a pre-built wheel from https://pypi.org/project/lxml/.

---

## Usage

### Randomize

Output paths are optional. When omitted, versions are written to the **same folder as the input file**, named `<stem>_A.docx`, `<stem>_B.docx`, etc. — so you only ever need to type the input path once:

```bash
# Two versions (default) — outputs written next to the input file
python3 ~/exam_randomizer/randomize_exam.py /path/to/exams/Exam1.docx

# Four versions
python3 ~/exam_randomizer/randomize_exam.py /path/to/exams/Exam1.docx --versions 4
```

You can still specify output paths explicitly if you need them elsewhere:

```bash
python3 randomize_exam.py Exam1.docx Exam1_A.docx Exam1_B.docx Exam1_C.docx
```

On **Windows**, replace `python3` with `py`.

Each run prints the seeds it used:

```
Generating 2 versions from: /path/to/exams/Exam1.docx
Seeds (randomly generated — record these to reproduce later):
  Version A: seed=1047382910  →  /path/to/exams/Exam1_A.docx
  Version B: seed=583920471   →  /path/to/exams/Exam1_B.docx
```

### Reproduce an exact previous run

```bash
python3 randomize_exam.py /path/to/exams/Exam1.docx --base-seed 1047382910
```

`--base-seed N` pins the first version's seed to N and uses N+1, N+2, … for subsequent versions, giving a fully reproducible output.

### Verify

When called with only the original file, `verify_exam.py` automatically finds the `_A`, `_B`, … versions next to it:

```bash
python3 ~/exam_randomizer/verify_exam.py /path/to/exams/Exam1.docx
```

Or pass version files explicitly:

```bash
python3 verify_exam.py Exam1.docx Exam1_A.docx Exam1_B.docx
```

```
Sections found:   True/False, Multiple Choice, Fill-in-the-Blank, Workout

Verifying: Exam1_A.docx
  Version label:  Version A
    ✓  True/False: correct question count (5)
    ✓  True/False: all question stems present
    ✓  True/False: question order differs from original
    ✓  Multiple Choice: correct question count (15)
    ...
    ✓  Multiple Choice: 'None of the above' is last in every question

──────────────────────────────────────────────
  PASSED:   52    WARNINGS: 0    FAILED:   0
──────────────────────────────────────────────
All checks passed.
```

`verify_exam.py` exits with code 1 if any check fails, so it can be used in a script.

---

## Recommended workflow

1. **Randomize** — run `randomize_exam.py` to generate all versions.

2. **Adjust manually** — open each version in Word and:
   - Add or remove blank lines and page breaks as needed so questions sit where you want them on the page.
   - Reposition any floating images that didn't follow their questions correctly (see [Floating images](#floating-images) below).

3. **Verify** — run `verify_exam.py` against the adjusted files. Because the verifier works from list-numbering structure and question text rather than whitespace or image positions, it will still pass correctly after manual edits.

---

## Templates

Two ready-to-edit starting points are included. Download the one that matches your exam format, replace the placeholder content with your own questions, and run the scripts straight away.

**`Exam_Template1.docx`** — midterm format, with all four section types: True/False, Multiple Choice, Fill-in-the-Blank, and Workout.

**`Exam_Template2.docx`** — final exam format, Multiple Choice only (35 questions).

`Exam_Template1.docx` is designed to demonstrate edge cases the randomizer handles:

| Section | What it demonstrates |
|---|---|
| True/False | Mixed list styles (two different numIds in the same section) — common when questions are copied from multiple sources |
| Multiple Choice | `None of the above` / `None of these` pinning; math equations; a question formatted with sub-statements (I, II, III) |
| Fill-in-the-Blank | Soft line returns (Shift+Enter) used to create blank space for student answers — these are preserved after shuffling |
| Workout | Sub-parts (a, b, c…) that stay with their question; per-question full-page breaks |

**What to change:** the cover-page fields (course name, exam number), instructions, point values, and of course all the questions themselves.

**What to keep:** the Heading 2 section headings (the keywords "True/False", "Multiple Choice", "Fill in the Blank" are what the scripts use to find each section), and the numbered-list formatting on every question stem and answer choice.

---

## Document format requirements

The scripts detect sections by looking for **Heading 2** paragraphs whose text contains recognisable keywords:

| Keyword in heading | Section detected |
|---|---|
| "true" | True/False |
| "multiple" | Multiple Choice |
| "fill" or "blank" | Fill-in-the-Blank |
| *(Heading 2 immediately before the first workout list item)* | Workout |

> ⚠️ **Each section heading must be its own standalone paragraph** formatted as Heading 2 — not typed inline within a question or answer choice. A common mistake is inserting a page break inside a list item and typing the next section heading on the same line: the heading text ends up inside the list item and is never recognised as a section boundary. If a section heading appears to be missing or lands in the wrong place after shuffling, check that it is a separate paragraph in the original document.

Questions must be formatted as **Word numbered lists** (the automatic kind, not manually typed numbers). Answer choices for True/False and Multiple Choice should be **bullet-point sub-items** under each question.

**Each section's questions must use their own Word list** — do not continue the same numbered list across a section heading into the next section. The script finds the boundary between adjacent sections by looking for a change in list identity (`numId` in the underlying XML). If two sections share one continuous list, the boundary between them will not be detected and questions from both sections will be shuffled together.

A quick way to check whether a new exam will work: run `verify_exam.py` against the original file and two versions. If all checks pass, the format was understood correctly.

---

## Accessibility

Because the script deep-copies entire paragraph elements (including all embedded XML), accessibility attributes travel with their content. Alt text on images and floating shapes is preserved verbatim, section headings are never shuffled, and automatic list numbering renumbers correctly after shuffling. Document-level settings (language, styles, page layout) are copied byte-for-byte and are unaffected.

If the original document is WCAG 2.1 AA compliant, the randomized versions should be too. A manual spot-check is still recommended, particularly for:

- **Math equations** — equation accessibility is copied unchanged from the original, so if alt text wasn't set there, it won't be in the versions either.
- **Fill-in-the-blank blanks** — see the floating shapes note below.
- **Page-anchored images** — images positioned relative to the page (rather than to a paragraph) won't move with their question if that question is shuffled.

---

## Security note

The shuffling uses seeds drawn from the operating system's random-number generator (`secrets` module), which is cryptographically unpredictable. A student who reads this source code gains no information about the ordering used for any particular exam run. The only way to know the ordering in advance is to have the seed, which is generated fresh each time and never stored by the scripts.

---

## Page flow

The script handles page layout automatically:

1. **Strips manual page breaks** from within question content: standalone page-break paragraphs and `pageBreakBefore` paragraph properties are removed so breaks placed around specific questions in the original don't land in arbitrary positions after shuffling. Soft line returns (Shift+Enter) within question content are left untouched — they are intentional spacing chosen by the exam author.

2. **Normalises spacing between questions.** Blank paragraphs that trail each question block in the source document are stripped, and a single uniform blank paragraph is inserted between every pair of questions. Spacing is therefore consistent across all versions regardless of how the source document was formatted.

3. **Keeps questions together** by adding `keepNext` (keep this paragraph on the same page as the next) and `keepLines` (keep all lines of this paragraph together) to every paragraph in a question block except the last. Word uses these flags — rather than manual spacing — to decide where page breaks fall after shuffling.

4. **Adds a clean page break before each section** after the first (e.g. Multiple Choice, Fill-in-the-Blank, Workout each start on a new page).

5. **Gives each Workout question its own page.** Every question after the first in the Workout section starts on a new page.

6. **Ensures "This page intentionally left blank" pages always start on a new page.** Include however many you need in the source document and the script will handle the breaks automatically.

> ⚠️ **"This page intentionally left blank" text must appear in its own paragraph, after the last question in the document** — not inside a question's list item. If it is typed inside a list item (e.g. after a page break within an answer choice), it will be treated as part of that question and shuffled along with it.

---

## Floating images

> ⚠️ **Figures may end up far from their questions if not set up correctly.** Before running the randomizer, check every figure in the source document using the steps below.

If a question has an associated figure, how well it follows its question after shuffling depends on how the image is anchored in Word.

**Inline images** (inserted directly into the text flow) always move with their paragraph and are unaffected by shuffling.

**Floating images** (inserted via Insert → Pictures and then set to wrap text) have a vertical position that is relative to either the paragraph or the page:
- *Relative to paragraph* — the image moves with its anchor paragraph and will follow its question correctly.
- *Relative to page* — the image stays at a fixed position on the page regardless of where its anchor paragraph ends up. After shuffling, this will place the image far from its question.

**To fix a page-anchored floating image:** right-click it in Word → **Format Picture → Layout & Properties → Position**, and change the vertical "Relative to" setting from *Page* to *Paragraph*. Do this for every figure in the source document before generating versions.

---

## A note on fill-in-the-blank blanks

If your FIB blanks are drawn as **floating Word shapes** (Insert → Shapes), they are anchored to their paragraph and move with it when questions are shuffled. The vertical position tracks the anchor paragraph; the horizontal position is typically fixed relative to the column. This looks correct in practice, but it is worth opening each output file in Word and scrolling through the FIB section before printing, just to confirm the blanks are positioned as expected.
