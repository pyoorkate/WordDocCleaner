import os
import sys
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

spinner = ["/", "-", "\\", "|"]

print("\n==================================")
print(".docx file formatting cleaner v0.9")
print("==================================")
print("\nStrips formatting, resets styles, and preserves core formatting.\nIncludes: Metadata stripping, isolated character style review, and empty line removal.")

def set_run_language(run, lang_code):
    rPr = run._element.get_or_add_rPr()
    for attr in ['w:val', 'w:eastAsia', 'w:bidi']:
        lang = OxmlElement('w:lang')
        lang.set(qn(attr), lang_code)
        rPr.append(lang)

def review_isolated_formatting(doc):
    print("\n--- Starting Isolated Formatting Review ---")
    for para in doc.paragraphs:
        current_pos = 0
        full_text = para.text
        for run in para.runs:
            sys.stdout.write(f"\r {spinner[current_pos % len(spinner)]} Reviewing...")
            sys.stdout.flush()
            clean_text = run.text.strip()
            run_len = len(run.text)

            if len(clean_text) == 1:
                active_formats = []
                if run.bold: active_formats.append("Bold")
                if run.italic: active_formats.append("Italic")
                if run.underline: active_formats.append("Underline")
                if run.font.strike: active_formats.append("Strikethrough")
                if run.font.name: active_formats.append(f"Font ({run.font.name})")
                if run.font.superscript: active_formats.append("Superscript")
                if run.font.subscript: active_formats.append("Subscript")
                if run.font.small_caps: active_formats.append("Small Caps")

                if active_formats:
                    start = max(0, current_pos - 30)
                    end = min(len(full_text), current_pos + 30)
                    before = full_text[start:current_pos]
                    after = full_text[current_pos + 1:end]
                    window = f"{before}[[{run.text}]]{after}"
                    print(f"\nContext: ...{window}...")
                    print(f"Target: '{run.text}' | Formatting: [{', '.join(active_formats)}]")
                    choice = input("Keep formatting for this section [y]es / [n]o (strip formatting): ").lower()
                    if choice == 'n':
                        run.bold = run.italic = run.underline = run.font.strike = run.font.superscript = run.font.subscript = run.font.small_caps = False
                        run.font.name = None
            current_pos += run_len

def ultimate_clean_docx():
    if len(sys.argv) > 2:
        input_file, output_file = sys.argv[1], sys.argv[2]
    else:
        input_file = input("\nEnter input .docx path: ").strip('"')
        output_file = input("Enter output .docx path: ").strip('"')

    if not os.path.exists(input_file):
        print("Error: File not found.")
        return

    doc = Document(input_file)
    lang_code = input("\nEnter language code (e.g., en-US) or Enter to skip: ").strip()

    # 1. Map styles
    used_styles = {p.style.name for p in doc.paragraphs}
    style_map = {}
    for name in sorted(used_styles):
        print(f"Style: '{name}'")
        choice = input("  1: Heading 1, 2: Heading 2, 3: Normal, OR [Enter]: Skip: ")
        if choice == '1': style_map[name] = 'Heading 1'
        elif choice == '2': style_map[name] = 'Heading 2'
        elif choice == '3': style_map[name] = 'Normal'

    # 2. Process Paragraphs
    print("\nProcessing paragraph styles and formatting...")
    strip_font = False
    strip_allcaps = False
    strip_smallcaps = False
    strip_super_subscript = False
    strip_underline = False


    choice = input("  Keep all caps sections [y]es / [n]o (strip all caps) ")
    if choice == 'n':
        strip_allcaps = True
    choice = input("  Keep small caps sections [y]es / [n]o (strip small caps) ")
    if choice == 'n':
        strip_smallcaps = True
    choice = input("  Keep font selections [y]es / [n]o (strip fonts) ")
    if choice == 'n':
        strip_font = True
    choice = input("  Keep super/subscript sections [y]es / [n]o (strip super/subscript) ")
    if choice == 'n':
        strip_super_subscript = True
    choice = input("  Keep underlined sections [y]es / [n]o (strip underline) ")
    if choice == 'n':
        strip_underline = True

    for i, para in enumerate(doc.paragraphs):
        sys.stdout.write(f"\r {spinner[i % len(spinner)]} Processing...")
        sys.stdout.flush()

        # APPLY STYLE WITH ERROR HANDLING
        if para.style.name in style_map:
            target_style = style_map[para.style.name]
            try:
                para.style = doc.styles[target_style]
            except KeyError:
                # If 'Heading 2' fails, try 'Heading2' or just skip
                try:
                    para.style = doc.styles[target_style.replace(" ", "")]
                except KeyError:
                    pass

        # Reset Geometry
        pf = para.paragraph_format
        pf.line_spacing = pf.space_before = pf.space_after = pf.alignment = None
        pf.left_indent = pf.right_indent = pf.first_line_indent = None

        for run in para.runs:
            if run.font.hidden:
                run.text = ""
                continue

            run.style = None
            b, i, u, s = run.bold, run.italic, run.underline, run.font.strike
            rPr = run._element.get_or_add_rPr()
            tags_to_kill = [
                qn('w:sz'), qn('w:szCs'), qn('w:color'),
                qn('w:highlight'), qn('w:shd'), qn('w:cs')
            ]
            if strip_font == True:
                tags_to_kill.extend([
                    qn('w:rFonts')
                ])
            if strip_allcaps == True:
                tags_to_kill.extend([
                    qn('w:caps')
                ])
            if strip_smallcaps == True:
                tags_to_kill.extend([
                    qn('w:smallCaps')
                ])
            if strip_super_subscript == True:
                tags_to_kill.extend([
                    qn('w:vertAlign')
                ])
            if strip_underline == True:
                tags_to_kill.extend([
                    qn('w:u')
                ])
                u = False

            """Previously this also stripped out high ascii / ansi values
            which I'm not sure was a great plan, so I'm removing it for the
            moment. The XML for those are: qn('w:ascii'), qn('w:hAnsi'), """
            for tag in tags_to_kill:

                element = rPr.find(tag)
                if element is not None:
                    rPr.remove(element)

            run.bold, run.italic, run.underline, run.font.strike = b, i, u, s
            if lang_code:
                set_run_language(run, lang_code)

    # 3. Remove Empty Paragraphs
    print("\n\nRemove empty paragraphs (extra carriage returns)?")
    rem_choice = input("  1: YES, [Enter]: Skip: ")
    if rem_choice == '1':
        print("Cleaning up whitespace...")
        for i, para in enumerate(list(doc.paragraphs)):
            sys.stdout.write(f"\r {spinner[i % len(spinner)]} Analyzing...")
            sys.stdout.flush()
            if not para.text.strip():
                p = para._element
                p.getparent().remove(p)

    # 4. Review isolated characters
    print("\n\nWould you like to review isolated formatted characters?")
    review_choice = input("  1: YES, [Enter]: Skip: ")
    if review_choice == '1':
        review_isolated_formatting(doc)

    # 5. Strip Metadata
    print("\nStrip Metadata?")
    choice = input("  1: YES, [Enter]: Skip: ")
    if choice == '1':
        core_props = doc.core_properties
        core_props.author = core_props.comments = core_props.keywords = ""
        core_props.last_modified_by = core_props.title = ""

    doc.save(output_file)
    print(f"\nDocument scrubbed and saved to: {output_file}")

if __name__ == "__main__":
    ultimate_clean_docx()
