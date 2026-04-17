"""Visual side-by-side comparison to spot structural issues."""
from docx import Document
from docx.oxml.ns import qn


def dump_visual(path, label):
    doc = Document(path)
    print(f"\n{'='*70}")
    print(f"  {label}")
    print(f"  {path}")
    print(f"{'='*70}\n")
    
    for i, para in enumerate(doc.paragraphs):
        style = para.style.name if para.style else "?"
        text = para.text
        
        # Check for borders (section underlines)
        pPr = para._element.find(qn('w:pPr'))
        has_bottom_border = False
        if pPr is not None:
            pBdr = pPr.find(qn('w:pBdr'))
            if pBdr is not None:
                bottom = pBdr.find(qn('w:bottom'))
                if bottom is not None:
                    has_bottom_border = True
        
        border_marker = " ___BORDER___" if has_bottom_border else ""
        
        # Show run breakdown for non-trivial paragraphs
        run_info = ""
        if len(para.runs) > 1:
            parts = []
            for r in para.runs:
                markers = ""
                if r.bold: markers += "B"
                if r.italic: markers += "I"
                if r.underline: markers += "U"
                if r.font.color and r.font.color.rgb:
                    markers += f"c:{r.font.color.rgb}"
                if markers:
                    parts.append(f"[{markers}]{r.text}")
                else:
                    parts.append(r.text)
            run_info = f"\n        RUNS: {'|'.join(parts)}"
        elif len(para.runs) == 1:
            r = para.runs[0]
            markers = ""
            if r.bold: markers += "B"
            if r.italic: markers += "I"  
            if r.underline: markers += "U"
            if markers:
                run_info = f" [{markers}]"
        
        if text.strip() or has_bottom_border:
            print(f"  [{i:2d}] ({style:16s}){border_marker}")
            print(f"        {text[:120]}")
            if run_info:
                print(f"       {run_info}")
            print()


dump_visual(r"template\resume_template.docx", "ORIGINAL TEMPLATE")
print("\n\n" + "#"*70 + "\n\n")
dump_visual(r"output\Rehan_Malik_Resume_data_engineer_amazon.docx", "GENERATED (Amazon Data Engineer)")
