import os
import sys
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_run_language(run, lang_code):
    """Sets the language of a specific run using XML attributes."""
    rPr = run._element.get_or_add_rPr()
    for attr in ['w:val', 'w:eastAsia', 'w:bidi']:
        lang = OxmlElement('w:lang')
        lang.set(qn(attr), lang_code)
        rPr.append(lang)

def deep_clean_docx():
    # 1. Handle File Paths
    if len(sys.argv) > 2:
        input_file, output_file = sys.argv[1], sys.argv[2]
    else:
        input_file = input("Enter input .docx path: ").strip('"')
        output_file = input("Enter output .docx path: ").strip('"')

    if not os.path.exists(input_file):
        print(f"Error: {input_file} not found.")
        return

    doc = Document(input_file)

    # 2. Setup Language and Style Mapping
    lang_code = input("\nEnter language code (e.g., en-US, en-GB) or Enter to skip: ").strip()

    used_styles = {p.style.name for p in doc.paragraphs}
    style_map = {}
    print(f"\nFound {len(used_styles)} styles in use.")

    for name in sorted(used_styles):
        print(f"Style: '{name}'")
        choice = input("  1: Normal, 2: Heading 1, 3: Heading 2, [Enter]: Skip: ")
        if choice == '1': style_map[name] = 'Normal'
        elif choice == '2': style_map[name] = 'Heading 1'
        elif choice == '3': style_map[name] = 'Heading 2'

    # 3. Processing Loop
    for para in doc.paragraphs:
        # Reassign Paragraph Style
        if para.style.name in style_map:
            para.style = doc.styles[style_map[para.style.name]]

        # Reset Paragraph Overrides
        pf = para.paragraph_format
        pf.line_spacing = pf.space_before = pf.space_after = pf.alignment = None

        for run in para.runs:
            # A. Remove Character Style overrides (e.g., "Body Text Char")
            run.style = None

            # B. Save specific formatting we want to keep
            b, i, u, s = run.bold, run.italic, run.underline, run.font.strike

            # C. Sledgehammer: Remove all font face and size tags from XML
            rPr = run._element.get_or_add_rPr()
            # Targets: Fonts (Latin/CS/EA) and Font Sizes
            tags_to_kill = [
                qn('w:rFonts'), qn('w:sz'), qn('w:szCs'),
                qn('w:rFonts'), qn('w:ascii'), qn('w:hAnsi'), qn('w:cs')
            ]
            for tag in tags_to_kill:
                element = rPr.find(tag)
                if element is not None:
                    rPr.remove(element)

            # D. Re-apply preserved traits
            run.bold, run.italic, run.underline, run.font.strike = b, i, u, s

            # E. Force Language
            if lang_code:
                set_run_language(run, lang_code)

    # 4. Save
    try:
        doc.save(output_file)
        print(f"\nCleaned file successfully saved to: {output_file}")
    except Exception as e:
        print(f"Failed to save: {e}")

if __name__ == "__main__":
    deep_clean_docx()
