"""
Tailor a resume to a job description using GitHub Models API,
preserving the exact DOCX structure, formatting, and styles.
"""

import os
import sys
import json
import copy
import argparse
import re
from pathlib import Path

import requests
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


GITHUB_MODELS_URL = "https://models.inference.ai.azure.com/chat/completions"
MODEL_NAME = "gpt-4o"

SECTIONS_TO_TAILOR = [
    "experience",
    "skills & certifications",
]

SECTIONS_TO_SKIP = [
    "education",
]


def read_jd(jd_path: str) -> str:
    ext = Path(jd_path).suffix.lower()
    if ext == ".docx":
        doc = Document(jd_path)
        return "\n".join(p.text for p in doc.paragraphs)
    elif ext == ".pdf":
        try:
            import fitz
            with fitz.open(jd_path) as pdf:
                return "\n".join(page.get_text() for page in pdf)
        except ImportError:
            sys.exit("PyMuPDF (fitz) required for PDF JDs. Install with: pip install PyMuPDF")
    else:
        with open(jd_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()


def extract_resume_text(doc: Document) -> str:
    lines = []
    for para in doc.paragraphs:
        if para.text.strip():
            lines.append(para.text.strip())
    return "\n".join(lines)


def parse_sections(doc: Document) -> list[dict]:
    """
    Parse the resume into logical sections based on Heading 1 markers
    and the known layout pattern.
    """
    sections = []
    current_section = None
    current_job = None

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        style = para.style.name if para.style else ""

        if i == 0 and style == "Heading 1":
            current_section = {"type": "name", "para_indices": [i], "text": text}
            sections.append(current_section)
            current_section = None
            continue

        if i == 1:
            sections.append({"type": "contact", "para_indices": [i], "text": text})
            continue

        if style == "Heading 1" and text.lower() in ("experience",):
            current_section = {"type": "experience", "heading_idx": i, "jobs": [], "para_indices": [i]}
            sections.append(current_section)
            continue

        if style == "Heading 1" and "education" in text.lower():
            current_section = {"type": "education", "para_indices": [i], "jobs": []}
            sections.append(current_section)
            continue

        if style == "Heading 1" and "certification" in text.lower():
            current_section = {"type": "skills", "para_indices": [i]}
            sections.append(current_section)
            continue

        if current_section:
            current_section["para_indices"].append(i)

    return sections


def build_experience_blocks(doc: Document) -> list[dict]:
    """
    Parse the resume into experience blocks: each block has a company header
    and a list of bullet-point paragraph indices.
    """
    blocks = []
    current_block = None
    in_experience = False

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        style = para.style.name if para.style else ""

        if style == "Heading 1" and text.lower() == "experience":
            in_experience = True
            continue

        if style == "Heading 1" and ("education" in text.lower() or "certification" in text.lower()):
            in_experience = False
            if current_block:
                blocks.append(current_block)
                current_block = None
            continue

        if not in_experience:
            continue

        if style == "Normal" and "\t" in text and not text.startswith("-"):
            if current_block:
                blocks.append(current_block)
            current_block = {
                "date_company_idx": i,
                "title_idx": None,
                "bullet_indices": [],
                "company_text": text,
            }
            continue

        if style == "Heading 1" and current_block and current_block["title_idx"] is None:
            current_block["title_idx"] = i
            current_block["title_text"] = text
            continue

        if style == "List Paragraph" and current_block:
            current_block["bullet_indices"].append(i)
            continue

    if current_block:
        blocks.append(current_block)

    return blocks


def build_skills_block(doc: Document) -> dict:
    """Extract the skills & certifications section paragraph indices."""
    skill_indices = []
    in_skills = False
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        style = para.style.name if para.style else ""

        if style == "Heading 1" and "certification" in text.lower():
            in_skills = True
            skill_indices.append(i)
            continue

        if in_skills:
            if style == "Heading 1" and "certification" not in text.lower() and "skill" not in text.lower():
                break
            skill_indices.append(i)

    return {"para_indices": skill_indices}


def call_github_model(token: str, system_prompt: str, user_prompt: str) -> str:
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": MODEL_NAME,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.4,
        "max_tokens": 4000,
    }
    resp = requests.post(GITHUB_MODELS_URL, headers=headers, json=payload, timeout=120)
    if resp.status_code != 200:
        print(f"API Error {resp.status_code}: {resp.text[:500]}")
        resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]


def tailor_bullets(token: str, jd_text: str, company: str, title: str,
                   original_bullets: list[str]) -> list[str]:
    """
    Use the LLM to rewrite experience bullets tailored to the JD,
    keeping the same count and similar length.
    """
    system_prompt = (
        "You are an expert resume writer. You tailor resume bullet points to match "
        "a target job description while keeping them truthful and grounded in the "
        "original experience. Preserve technical depth and quantified metrics. "
        "Do NOT invent new experiences or metrics that aren't in the original. "
        "Do NOT add any prefix numbering. Return ONLY the bullet texts, one per line, "
        "with no extra commentary. Keep the same number of bullets."
    )
    user_prompt = (
        f"TARGET JOB DESCRIPTION:\n{jd_text}\n\n"
        f"CURRENT ROLE: {title} at {company}\n\n"
        f"ORIGINAL BULLETS (rewrite these):\n"
        + "\n".join(f"- {b}" for b in original_bullets)
        + "\n\nRewrite each bullet to better align with the target JD. "
        "Keep the same number of bullets. Return one bullet per line, no dashes or numbers."
    )

    result = call_github_model(token, system_prompt, user_prompt)
    lines = [ln.strip().lstrip("-•0123456789. ") for ln in result.strip().split("\n") if ln.strip()]

    if len(lines) != len(original_bullets):
        if len(lines) > len(original_bullets):
            lines = lines[:len(original_bullets)]
        else:
            lines.extend(original_bullets[len(lines):])

    return lines


def tailor_skills(token: str, jd_text: str, original_skills_text: str) -> str:
    system_prompt = (
        "You are an expert resume writer. You reorder and lightly adjust the skills "
        "section of a resume to better match a target job description. "
        "Do NOT invent skills the candidate doesn't have. You may reorder, "
        "emphasize, or slightly rephrase existing skills. Keep the exact same "
        "category structure and formatting pattern. Return the full skills text."
    )
    user_prompt = (
        f"TARGET JOB DESCRIPTION:\n{jd_text}\n\n"
        f"ORIGINAL SKILLS SECTION:\n{original_skills_text}\n\n"
        "Reorder and adjust to better match the JD. Keep the same categories and format."
    )
    return call_github_model(token, system_prompt, user_prompt)


def replace_paragraph_text_preserve_format(para, new_text: str):
    """
    Replace the text content of a paragraph while preserving
    all run-level formatting (bold, italic, font, size, color).
    """
    if not para.runs:
        if para.text != new_text:
            para.text = new_text
        return

    if len(para.runs) == 1:
        para.runs[0].text = new_text
        return

    # For multi-run paragraphs: put all text in the first run,
    # clear the rest — preserving the first run's formatting
    first_run = para.runs[0]
    first_run.text = new_text
    for run in para.runs[1:]:
        run.text = ""


def update_resume(doc: Document, token: str, jd_text: str) -> Document:
    exp_blocks = build_experience_blocks(doc)
    print(f"Found {len(exp_blocks)} experience blocks")

    for block in exp_blocks:
        company = block["company_text"]
        title = block.get("title_text", "")
        bullet_indices = block["bullet_indices"]

        if not bullet_indices:
            continue

        original_bullets = [doc.paragraphs[idx].text.strip() for idx in bullet_indices]
        print(f"\nTailoring: {title} at {company}")
        print(f"  {len(original_bullets)} bullets to rewrite...")

        new_bullets = tailor_bullets(token, jd_text, company, title, original_bullets)

        for idx, new_text in zip(bullet_indices, new_bullets):
            replace_paragraph_text_preserve_format(doc.paragraphs[idx], new_text)
            print(f"  [OK] Updated bullet {idx}")

    # Tailor skills section
    skills_block = build_skills_block(doc)
    if skills_block["para_indices"]:
        skill_paras = skills_block["para_indices"]
        original_skills_lines = []
        for idx in skill_paras:
            t = doc.paragraphs[idx].text.strip()
            if t:
                original_skills_lines.append(t)

        skills_text = "\n".join(original_skills_lines)
        print(f"\nTailoring skills section ({len(skill_paras)} paragraphs)...")
        new_skills = tailor_skills(token, jd_text, skills_text)

        new_skill_lines = [ln.strip() for ln in new_skills.strip().split("\n") if ln.strip()]

        body_para_indices = [
            idx for idx in skill_paras
            if doc.paragraphs[idx].style.name in ("Body Text", "Normal")
            and doc.paragraphs[idx].text.strip()
        ]

        for idx, new_line in zip(body_para_indices, new_skill_lines):
            replace_paragraph_text_preserve_format(doc.paragraphs[idx], new_line)
            print(f"  [OK] Updated skill para {idx}")

    return doc


def convert_to_pdf(docx_path: str, pdf_path: str):
    """Convert DOCX to PDF using LibreOffice (available in GitHub Actions)."""
    import subprocess
    output_dir = str(Path(pdf_path).parent)
    result = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf",
         "--outdir", output_dir, docx_path],
        capture_output=True, text=True, timeout=120
    )
    if result.returncode != 0:
        print(f"LibreOffice PDF conversion failed: {result.stderr}")
        soffice_result = subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf",
             "--outdir", output_dir, docx_path],
            capture_output=True, text=True, timeout=120
        )
        if soffice_result.returncode != 0:
            print(f"soffice also failed: {soffice_result.stderr}")
            raise RuntimeError("PDF conversion failed")

    expected_pdf = Path(output_dir) / (Path(docx_path).stem + ".pdf")
    if expected_pdf.exists() and str(expected_pdf) != pdf_path:
        expected_pdf.rename(pdf_path)


def main():
    parser = argparse.ArgumentParser(description="Tailor resume to a job description")
    parser.add_argument("--jd", required=True, help="Path to the job description file")
    parser.add_argument("--template", default="template/resume_template.docx",
                        help="Path to the resume template DOCX")
    parser.add_argument("--output-dir", default="output", help="Output directory")
    parser.add_argument("--token", default=None,
                        help="GitHub token (or set GITHUB_TOKEN env var)")
    args = parser.parse_args()

    token = args.token or os.environ.get("GH_MODELS_TOKEN") or os.environ.get("GITHUB_TOKEN")
    if not token:
        sys.exit("ERROR: Provide --token or set GH_MODELS_TOKEN environment variable")

    jd_path = Path(args.jd)
    if not jd_path.exists():
        sys.exit(f"ERROR: JD file not found: {jd_path}")

    template_path = Path(args.template)
    if not template_path.exists():
        sys.exit(f"ERROR: Template not found: {template_path}")

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    jd_name = jd_path.stem.replace(" ", "_")
    output_docx = output_dir / f"Rehan_Malik_Resume_{jd_name}.docx"
    output_pdf = output_dir / f"Rehan_Malik_Resume_{jd_name}.pdf"

    print(f"Reading JD: {jd_path}")
    jd_text = read_jd(str(jd_path))
    print(f"JD length: {len(jd_text)} characters")

    print(f"Loading template: {template_path}")
    doc = Document(str(template_path))

    print("Tailoring resume...")
    updated_doc = update_resume(doc, token, jd_text)

    print(f"Saving DOCX: {output_docx}")
    updated_doc.save(str(output_docx))

    print(f"Converting to PDF: {output_pdf}")
    try:
        convert_to_pdf(str(output_docx), str(output_pdf))
        print("PDF created successfully!")
    except Exception as e:
        print(f"PDF conversion skipped ({e}). DOCX is still available.")

    print("\nDone!")
    print(f"  DOCX: {output_docx}")
    print(f"  PDF:  {output_pdf}")


if __name__ == "__main__":
    main()
