import io
import json
import re
from xml.sax.saxutils import escape
from datetime import datetime
from typing import List, Optional

import streamlit as st
from docx import Document
from groq import Groq
from pydantic import BaseModel, Field, ValidationError, field_validator
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
)

# -------------------------------------------------------------------
# Page setup and visual theme
# -------------------------------------------------------------------
st.set_page_config(
    page_title="GradeWise AI | Assignment Feedback",
    page_icon="✦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap');
    :root { --ink:#172033; --muted:#738096; --line:#e7ebf2; --blue:#536dfe; --teal:#13b8a6; }
    html, body, [class*="css"] { font-family:'DM Sans',sans-serif; }
    .stApp { background: radial-gradient(circle at 90% 0%, #eef2ff 0, transparent 28%), #f6f8fc; color:var(--ink); }
    [data-testid="stHeader"] { background:transparent; }
    [data-testid="stSidebar"] { background:#101a30; border-right:1px solid #25324a; }
    [data-testid="stSidebar"] * { color:#edf2ff !important; }
    [data-testid="stSidebar"] .stCaption { color:#a9b7d0 !important; }
    .block-container { padding-top:2rem; padding-bottom:3rem; max-width:1440px; }
    .hero { padding:26px 30px; border-radius:22px; color:white; background:linear-gradient(115deg,#17233e 0%,#293e73 58%,#526dff 100%); box-shadow:0 14px 38px rgba(33,52,99,.16); margin-bottom:22px; }
    .hero-kicker { color:#b9c8ff; text-transform:uppercase; letter-spacing:2px; font-size:11px; font-weight:700; }
    .hero h1 { font-family:'Manrope',sans-serif; font-size:34px; line-height:1.15; margin:8px 0; color:#fff; font-weight:800; }
    .hero p { color:#e0e7ff; margin:0; font-size:14px; }
    .panel { background:white; border:1px solid var(--line); border-radius:18px; padding:20px 22px; box-shadow:0 5px 20px rgba(32,48,82,.035); }
    .section-label { font-size:11px; text-transform:uppercase; letter-spacing:1.4px; font-weight:800; color:#7b88a0; margin-bottom:7px; }
    .metric { background:white; border:1px solid var(--line); padding:16px 18px; border-radius:16px; }
    .metric .value { font-family:'Manrope'; font-weight:800; font-size:25px; color:#202d49; }
    .metric .caption { color:#7b88a0; font-size:12px; }
    .criterion-title { font-family:'Manrope'; font-weight:800; font-size:18px; color:#1c2943; }
    .stButton > button[kind="primary"] { background:linear-gradient(100deg,#536dfe,#725cff); border:0; border-radius:11px; font-weight:700; min-height:44px; }
    .stButton > button { border-radius:10px; }
    div[data-testid="stExpander"] { background:white; border:1px solid var(--line); border-radius:14px; }
    .note-box { background:#fff9e9; border:1px solid #f4e5b8; border-radius:12px; padding:12px 14px; color:#6b5621; font-size:13px; }
    .success-strip { padding:12px 15px; border-radius:12px; background:#eafaf5; color:#176b58; border:1px solid #ccefe3; font-weight:600; }
    .small-muted { color:#7b88a0; font-size:12px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# -------------------------------------------------------------------
# Structured response models
# -------------------------------------------------------------------
class CriterionFeedback(BaseModel):
    criterion: str = Field(min_length=1, description="Criterion name from the supplied rubric")
    score: float = Field(ge=0, description="Suggested score for this criterion")
    max_score: float = Field(gt=0, description="Maximum points for this criterion")
    evidence_quote: str = Field(description="Exact, verbatim quote from the student submission")
    strengths: List[str] = Field(min_length=1)
    improvement_steps: List[str] = Field(min_length=1)

    @field_validator("criterion")
    @classmethod
    def clean_criterion(cls, value):
        return value.strip()

    @field_validator("evidence_quote")
    @classmethod
    def quote_must_not_be_empty(cls, value):
        if not value.strip():
            raise ValueError("An evidence quote is required; use an explicit note if no evidence exists.")
        return value.strip()

    @field_validator("score")
    @classmethod
    def score_is_finite(cls, value):
        if value != value or value in (float("inf"), float("-inf")):
            raise ValueError("Score must be a finite number.")
        return value


class AssignmentEvaluation(BaseModel):
    overall_summary: str = Field(min_length=1)
    criteria: List[CriterionFeedback] = Field(min_length=1)
    teacher_review_notes: List[str] = Field(default_factory=list)

    @field_validator("criteria")
    @classmethod
    def scores_within_maximum(cls, criteria):
        for item in criteria:
            if item.score > item.max_score:
                raise ValueError(
                    f"Score for '{item.criterion}' ({item.score}) exceeds maximum ({item.max_score})."
                )
        return criteria


# -------------------------------------------------------------------
# File extraction and helpers
# -------------------------------------------------------------------
def extract_text(uploaded_file) -> str:
    """Extract readable text from supported teacher-provided files."""
    if uploaded_file is None:
        return ""
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    try:
        if name.endswith(".txt"):
            for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
                try:
                    return data.decode(encoding)
                except UnicodeDecodeError:
                    continue
            return data.decode("utf-8", errors="replace")
        if name.endswith(".pdf"):
            reader = PdfReader(io.BytesIO(data))
            pages = []
            for page in reader.pages:
                pages.append(page.extract_text() or "")
            return "\n\n".join(pages)
        if name.endswith(".docx"):
            document = Document(io.BytesIO(data))
            paragraphs = [p.text for p in document.paragraphs if p.text.strip()]
            for table in document.tables:
                for row in table.rows:
                    paragraphs.append(" | ".join(cell.text.strip() for cell in row.cells))
            return "\n".join(paragraphs)
    except Exception as exc:
        raise ValueError(f"Could not read {uploaded_file.name}: {exc}") from exc
    raise ValueError("Unsupported file type. Please use PDF, DOCX, or TXT.")


def combine_input(uploaded_file, pasted_text: str, label: str) -> str:
    file_text = extract_text(uploaded_file)
    pasted_text = (pasted_text or "").strip()
    if file_text.strip() and pasted_text:
        return f"{file_text.strip()}\n\n--- Additional pasted {label} text ---\n{pasted_text}"
    return file_text.strip() or pasted_text


def safe_filename(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9_-]+", "-", value.strip()).strip("-")
    return (value[:48] or "assignment-feedback").lower()


def validate_evidence_quotes(evaluation: AssignmentEvaluation, submission: str):
    """Warn on quotes not found verbatim; never silently rewrite model evidence."""
    source = submission.casefold()
    warnings = []
    for item in evaluation.criteria:
        quote = item.evidence_quote.strip().strip('"“”')
        if quote and quote.casefold() not in source:
            warnings.append(
                f"Evidence quote for **{item.criterion}** was not found as an exact text match. "
                "Please verify it against the original submission (PDF extraction can alter spacing)."
            )
    return warnings


def call_groq(rubric: str, submission: str, level: str, tone: str) -> AssignmentEvaluation:
    api_key = st.secrets.get("GROQ_API_KEY", "")
    if not api_key:
        raise RuntimeError(
            "GROQ_API_KEY is missing. Add it in Streamlit Cloud → App settings → Secrets."
        )

    client = Groq(api_key=api_key)
    system_prompt = """
You are GradeWise, a careful academic assessment assistant supporting a teacher.
Evaluate the student's submission ONLY against the criteria, descriptors, point values,
and requirements explicitly present in the teacher-provided grading rubric.

NON-NEGOTIABLE RULES:
1. Do not invent, add, infer, or apply criteria that are not in the rubric.
2. Do not penalize or reward grammar, spelling, style, citations, formatting, originality,
   completeness, or any other feature unless the rubric explicitly assesses it.
3. Treat the rubric as the sole authority for scoring. Do not use external standards,
   personal preferences, assumptions about the course, or the academic level to change
   the rubric. Academic level may affect wording only, not the criteria or scoring standard.
4. For every rubric criterion, return one criterion result with its exact criterion name,
   suggested score, maximum points, a short direct verbatim evidence quote from the
   student's work, specific strengths, and actionable improvement steps.
5. Evidence quotes must be copied exactly from the submission. Never fabricate quotes.
   If no supporting evidence exists, set evidence_quote to:
   "No direct supporting evidence found in the submitted work."
   That exact sentence is permitted as a no-evidence marker and is not a student quote.
6. Do not rewrite, complete, paraphrase into a replacement, or generate a model answer
   for the student's assignment. Improvement steps must coach the student, not do the work.
7. Use the requested tone while remaining constructive, respectful, specific, and fair.
8. Return every assessable criterion in the rubric, in the same order. Do not merge criteria.
9. If the rubric is ambiguous, internally flag the ambiguity in teacher_review_notes and
   make the most conservative rubric-grounded judgment. Do not invent missing point values.
10. Return valid JSON only, matching the requested structure. No markdown fences.

JSON structure:
{
  "overall_summary": "brief rubric-based overview",
  "criteria": [
    {
      "criterion": "rubric criterion name",
      "score": 0,
      "max_score": 0,
      "evidence_quote": "verbatim quote or exact no-evidence marker",
      "strengths": ["specific strength"],
      "improvement_steps": ["actionable next step"]
    }
  ],
  "teacher_review_notes": ["private note for teacher; never student-facing"]
}
"""
    user_prompt = f"""
ACADEMIC LEVEL (wording only): {level}
DESIRED FEEDBACK TONE: {tone}

TEACHER'S GRADING RUBRIC:
<grading_rubric>
{rubric}
</grading_rubric>

STUDENT'S ASSIGNMENT:
<student_submission>
{submission}
</student_submission>

Assess only the supplied rubric. Preserve rubric criterion order. Return JSON only.
"""
    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        temperature=0.1,
        max_tokens=6000,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    )
    raw = response.choices[0].message.content or ""
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise ValueError(f"The AI returned invalid JSON: {exc}") from exc
    try:
        result = AssignmentEvaluation.model_validate(parsed)
    except ValidationError as exc:
        raise ValueError(f"The AI response did not pass validation:\n{exc}") from exc

    # Additional deterministic guard: every criterion returned must be named in the rubric.
    rubric_lower = rubric.casefold()
    ungrounded = [
        item.criterion for item in result.criteria
        if item.criterion.casefold() not in rubric_lower
    ]
    if ungrounded:
        raise ValueError(
            "Rubric alignment check failed. These returned criterion names were not found "
            f"in the rubric text: {', '.join(ungrounded)}. Please clarify the rubric and try again."
        )
    return result


# -------------------------------------------------------------------
# PDF report generation (private teacher notes are intentionally excluded)
# -------------------------------------------------------------------
def build_pdf_report(
    student_name: str,
    assignment_title: str,
    level: str,
    overall_summary: str,
    criteria: list,
) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, rightMargin=18*mm, leftMargin=18*mm,
        topMargin=17*mm, bottomMargin=17*mm,
        title=f"Feedback report - {assignment_title}",
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="ReportTitle", parent=styles["Title"], fontName="Helvetica-Bold",
        fontSize=21, leading=25, textColor=colors.HexColor("#1D2B4B"),
        alignment=TA_CENTER, spaceAfter=5*mm,
    ))
    styles.add(ParagraphStyle(
        name="Subtle", parent=styles["Normal"], fontSize=9,
        textColor=colors.HexColor("#6B7890"), leading=13,
    ))
    styles.add(ParagraphStyle(
        name="Criterion", parent=styles["Heading2"], fontSize=13,
        textColor=colors.HexColor("#344B91"), leading=17, spaceBefore=5*mm,
        spaceAfter=2*mm,
    ))
    styles.add(ParagraphStyle(
        name="BodySmall", parent=styles["BodyText"], fontSize=9.2,
        leading=13, spaceAfter=2*mm,
    ))
    story = [
        Paragraph("ASSIGNMENT FEEDBACK REPORT", styles["ReportTitle"]),
        Paragraph(
            f"<b>Student:</b> {escape(student_name or 'Not specified')} &nbsp;&nbsp; "
            f"<b>Assignment:</b> {escape(assignment_title or 'Not specified')}",
            styles["Subtle"],
        ),
        Paragraph(f"<b>Academic level:</b> {escape(level)}", styles["Subtle"]),
        Paragraph(f"<b>Generated:</b> {datetime.now().strftime('%d %b %Y, %H:%M')}", styles["Subtle"]),
        Spacer(1, 5*mm),
        Paragraph("Overall rubric-based feedback", styles["Heading2"]),
        Paragraph(escape(overall_summary).replace("\n", "<br/>"), styles["BodySmall"]),
        Spacer(1, 2*mm),
    ]
    total_score = sum(float(c["score"]) for c in criteria)
    total_max = sum(float(c["max_score"]) for c in criteria)
    story.append(Paragraph(
        f"<b>Suggested total:</b> {total_score:g} / {total_max:g}",
        styles["BodySmall"],
    ))

    for index, item in enumerate(criteria, 1):
        group = [
            Paragraph(f"{index}. {escape(item['criterion'])}", styles["Criterion"]),
            Paragraph(
                f"<b>Suggested score:</b> {float(item['score']):g} / {float(item['max_score']):g}",
                styles["BodySmall"],
            ),
            Paragraph("<b>Evidence from submission</b>", styles["BodySmall"]),
            Paragraph(escape(item["evidence_quote"]).replace("\n", "<br/>"), styles["BodySmall"]),
            Paragraph("<b>Strengths</b>", styles["BodySmall"]),
        ]
        for strength in item["strengths"]:
            group.append(Paragraph("• " + escape(strength), styles["BodySmall"]))
        group.append(Paragraph("<b>Next steps for improvement</b>", styles["BodySmall"]))
        for step in item["improvement_steps"]:
            group.append(Paragraph("• " + escape(step), styles["BodySmall"]))
        story.append(KeepTogether(group))

    story.append(Spacer(1, 6*mm))
    story.append(Paragraph(
        "This report contains suggested AI-assisted feedback for teacher review. "
        "The teacher remains responsible for checking evidence, scores, and final decisions.",
        styles["Subtle"],
    ))
    doc.build(story)
    return buffer.getvalue()


# -------------------------------------------------------------------
# Sidebar
# -------------------------------------------------------------------
with st.sidebar:
    st.markdown("## ✦ GradeWise AI")
    st.caption("Rubric-first assessment workspace")
    st.markdown("---")
    st.markdown("**Assessment settings**")
    academic_level = st.selectbox(
        "Student academic level",
        ["Middle school", "High school", "Undergraduate", "Postgraduate", "Professional / vocational"],
        index=1,
    )
    feedback_tone = st.selectbox(
        "Feedback tone",
        ["Supportive and encouraging", "Formal and academic", "Direct and concise", "Detailed and coaching-focused"],
        index=0,
    )
    st.markdown("---")
    st.markdown("**Privacy & review**")
    st.caption(
        "Your rubric and submission are sent to Groq to generate feedback. "
        "Private teacher notes are kept separate from the student-facing report."
    )
    st.caption("Model: `llama-3.3-70b-versatile`")
    st.markdown("---")
    st.caption("GradeWise provides suggestions, not final grades. Always verify before sharing.")

# -------------------------------------------------------------------
# Header
# -------------------------------------------------------------------
st.markdown(
    """
    <div class="hero">
      <div class="hero-kicker">Teacher assessment studio · AI-assisted review</div>
      <h1>Feedback that follows<br>your rubric, not guesswork.</h1>
      <p>Upload a rubric and student work. Review every suggested score, evidence quote and next step before sharing.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if "evaluation" not in st.session_state:
    st.session_state.evaluation = None
if "report_meta" not in st.session_state:
    st.session_state.report_meta = {}

# -------------------------------------------------------------------
# Input workspace
# -------------------------------------------------------------------
st.markdown('<div class="section-label">01 / Assessment materials</div>', unsafe_allow_html=True)
left, right = st.columns(2, gap="large")

with left:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown("### Grading rubric")
    st.caption("Upload a document or paste the full rubric, including point values and descriptors.")
    rubric_file = st.file_uploader(
        "Upload rubric", type=["pdf", "docx", "txt"], key="rubric_file",
        help="Supported formats: PDF, DOCX and TXT.",
    )
    rubric_paste = st.text_area(
        "Or paste rubric text", height=190, key="rubric_paste",
        placeholder="Example:\nResearch quality (10 points): ...\nEvidence and analysis (15 points): ...",
    )
    st.markdown('</div>', unsafe_allow_html=True)

with right:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown("### Student assignment")
    st.caption("Upload the student's work or paste the submission text.")
    assignment_file = st.file_uploader(
        "Upload assignment", type=["pdf", "docx", "txt"], key="assignment_file",
        help="Supported formats: PDF, DOCX and TXT.",
    )
    assignment_paste = st.text_area(
        "Or paste assignment text", height=190, key="assignment_paste",
        placeholder="Paste the complete student submission here...",
    )
    st.markdown('</div>', unsafe_allow_html=True)

meta1, meta2 = st.columns(2)
with meta1:
    student_name = st.text_input("Student name (optional)", placeholder="e.g. Alex Morgan")
with meta2:
    assignment_title = st.text_input("Assignment title (optional)", placeholder="e.g. Environmental Science Report")

st.markdown("")
generate_col, reset_col, _ = st.columns([1.5, 1, 3])
with generate_col:
    generate_clicked = st.button("✦ Generate rubric feedback", type="primary", use_container_width=True)
with reset_col:
    if st.button("Clear review", use_container_width=True):
        st.session_state.evaluation = None
        st.session_state.report_meta = {}
        st.rerun()

if generate_clicked:
    try:
        rubric_text = combine_input(rubric_file, rubric_paste, "rubric")
        submission_text = combine_input(assignment_file, assignment_paste, "assignment")
        if len(rubric_text) < 20:
            st.error("Please provide a complete grading rubric (at least 20 characters).")
        elif len(submission_text) < 30:
            st.error("Please provide the student's assignment (at least 30 characters).")
        else:
            with st.spinner("Reading the rubric and assessing the submission…"):
                evaluation = call_groq(
                    rubric_text, submission_text, academic_level, feedback_tone
                )
            st.session_state.evaluation = evaluation.model_dump()
            st.session_state.report_meta = {
                "student_name": student_name,
                "assignment_title": assignment_title,
                "academic_level": academic_level,
                "submission_text": submission_text,
                "rubric_text": rubric_text,
            }
            st.success("Feedback generated and passed Pydantic validation. Review each criterion below.")
    except Exception as exc:
        st.error(str(exc))

# -------------------------------------------------------------------
# Editable teacher review workspace
# -------------------------------------------------------------------
if st.session_state.evaluation:
    st.markdown("---")
    st.markdown('<div class="section-label">02 / Teacher review workspace</div>', unsafe_allow_html=True)
    st.markdown("## Review and finalize")
    st.caption("Every field below is editable. Suggested scores are not final until you approve them.")

    evaluation_data = st.session_state.evaluation
    meta = st.session_state.report_meta
    quote_warnings = validate_evidence_quotes(
        AssignmentEvaluation.model_validate(evaluation_data),
        meta.get("submission_text", ""),
    )
    if quote_warnings:
        for warning in quote_warnings:
            st.warning(warning)

    st.markdown('<div class="note-box">Teacher-only notes are private workspace notes. They are excluded from the downloadable student feedback report.</div>', unsafe_allow_html=True)
    evaluation_data["overall_summary"] = st.text_area(
        "Overall feedback summary",
        value=evaluation_data.get("overall_summary", ""),
        height=105,
        key="overall_summary_edit",
    )

    criteria = evaluation_data.get("criteria", [])
    total_score = sum(float(c.get("score", 0)) for c in criteria)
    total_max = sum(float(c.get("max_score", 0)) for c in criteria)
    percentage = (total_score / total_max * 100) if total_max else 0
    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(f'<div class="metric"><div class="value">{total_score:g} / {total_max:g}</div><div class="caption">Suggested total score</div></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(f'<div class="metric"><div class="value">{percentage:.1f}%</div><div class="caption">Suggested percentage</div></div>', unsafe_allow_html=True)
    with m3:
        st.markdown(f'<div class="metric"><div class="value">{len(criteria)}</div><div class="caption">Rubric criteria assessed</div></div>', unsafe_allow_html=True)

    st.markdown("### Criterion-by-criterion feedback")
    for idx, criterion in enumerate(criteria):
        title = criterion.get("criterion", f"Criterion {idx + 1}")
        with st.expander(f"{idx + 1:02d}  ·  {title}", expanded=(idx == 0)):
            st.markdown(f'<div class="criterion-title">{title}</div>', unsafe_allow_html=True)
            score_col, max_col = st.columns(2)
            with score_col:
                criterion["score"] = st.number_input(
                    "Suggested score",
                    min_value=0.0,
                    max_value=float(criterion.get("max_score", 100)),
                    value=min(float(criterion.get("score", 0)), float(criterion.get("max_score", 100))),
                    step=0.5,
                    key=f"score_{idx}",
                )
            with max_col:
                st.number_input(
                    "Maximum points (rubric)",
                    min_value=0.5,
                    max_value=10000.0,
                    value=float(criterion.get("max_score", 1)),
                    step=0.5,
                    disabled=True,
                    key=f"max_{idx}",
                )
            criterion["evidence_quote"] = st.text_area(
                "Evidence quote from student work",
                value=criterion.get("evidence_quote", ""),
                height=85,
                key=f"quote_{idx}",
                help="Verify that this is a direct quote. The teacher may edit it.",
            )
            strengths_text = "\n".join(criterion.get("strengths", []))
            edited_strengths = st.text_area(
                "Strengths (one per line)",
                value=strengths_text,
                height=100,
                key=f"strengths_{idx}",
            )
            criterion["strengths"] = [line.strip() for line in edited_strengths.splitlines() if line.strip()]
            steps_text = "\n".join(criterion.get("improvement_steps", []))
            edited_steps = st.text_area(
                "Actionable improvement steps (one per line)",
                value=steps_text,
                height=110,
                key=f"steps_{idx}",
            )
            criterion["improvement_steps"] = [line.strip() for line in edited_steps.splitlines() if line.strip()]

    st.markdown("### Private AI review notes")
    private_notes = st.session_state.evaluation.get("teacher_review_notes", [])
    private_notes_text = st.text_area(
        "Visible only in this teacher workspace; not included in the student report",
        value="\n".join(private_notes),
        height=110,
        key="private_notes_edit",
    )
    st.session_state.evaluation["teacher_review_notes"] = [
        line.strip() for line in private_notes_text.splitlines() if line.strip()
    ]

    # Keep all edits in session state before downloads.
    st.session_state.evaluation["criteria"] = criteria
    try:
        AssignmentEvaluation.model_validate(st.session_state.evaluation)
        st.markdown('<div class="success-strip">✓ Review data is structurally valid. You can download the finalized student report.</div>', unsafe_allow_html=True)
        pdf_bytes = build_pdf_report(
            meta.get("student_name", ""),
            meta.get("assignment_title", ""),
            meta.get("academic_level", academic_level),
            st.session_state.evaluation["overall_summary"],
            st.session_state.evaluation["criteria"],
        )
        download_name = safe_filename(meta.get("assignment_title", "assignment-feedback"))
        st.download_button(
            "⬇ Download finalized feedback report (PDF)",
            data=pdf_bytes,
            file_name=f"{download_name}-feedback.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True,
        )
        # JSON is useful for a teacher's editable/archive workflow. It includes private notes.
        archive_json = json.dumps(st.session_state.evaluation, ensure_ascii=False, indent=2)
        with st.expander("Download teacher review archive (includes private notes)"):
            st.download_button(
                "Download teacher archive (JSON)",
                data=archive_json.encode("utf-8"),
                file_name=f"{download_name}-teacher-review.json",
                mime="application/json",
            )
    except ValidationError as exc:
        st.error(f"Please correct the review fields before downloading:\n{exc}")

else:
    st.markdown("---")
    st.markdown(
        """
        <div class="panel">
          <div class="section-label">How it works</div>
          <div style="display:flex;gap:18px;flex-wrap:wrap;">
            <div style="flex:1;min-width:180px;"><b>01 · Provide</b><p class="small-muted">Add the grading rubric and student submission.</p></div>
            <div style="flex:1;min-width:180px;"><b>02 · Assess</b><p class="small-muted">Groq evaluates only the criteria in your rubric.</p></div>
            <div style="flex:1;min-width:180px;"><b>03 · Review</b><p class="small-muted">Edit scores, evidence and feedback before sharing.</p></div>
            <div style="flex:1;min-width:180px;"><b>04 · Export</b><p class="small-muted">Download a clean PDF report for the student.</p></div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown(
    "<div style='text-align:center;color:#98a2b5;font-size:11px;padding:24px 0 5px;'>GRADEWISE AI · TEACHER-IN-THE-LOOP ASSESSMENT · VERIFY BEFORE PUBLISHING</div>",
    unsafe_allow_html=True,
)
