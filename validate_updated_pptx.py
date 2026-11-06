"""
validate_updated_pptx.py
Validates ARMEK_Final_Project_Review_Updated.pptx against all project review requirements.
"""

import pptx
from pptx.enum.shapes import MSO_SHAPE_TYPE

def validate():
    filename = "ARMEK_Final_Project_Review_Updated.pptx"
    prs = pptx.Presentation(filename)
    
    total_slides = len(prs.slides)
    print(f"=== Presentation Validation Report: {filename} ===")
    print(f"Total Slides: {total_slides}")
    assert total_slides == 11, f"Expected exactly 11 slides, found {total_slides}"

    all_text = []
    prototype_matches = []

    for idx, slide in enumerate(prs.slides):
        slide_num = idx + 1
        slide_text = []
        has_table = False
        table_dims = None
        
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text
                if txt.strip():
                    slide_text.append(txt.strip())
            elif shape.has_table:
                has_table = True
                t = shape.table
                table_dims = (len(t.rows), len(t.columns))
                for row in t.rows:
                    for cell in row.cells:
                        txt = cell.text_frame.text
                        if txt.strip():
                            slide_text.append(txt.strip())

        full_slide_str = " | ".join(slide_text)
        all_text.append(full_slide_str)
        
        # Check for forbidden word 'prototype'
        for st in slide_text:
            if "prototype" in st.lower():
                prototype_matches.append((slide_num, st))

        print(f"\n--- Slide {slide_num} ---")
        title_snippet = slide_text[0][:60] if slide_text else "EMPTY"
        print(f"Summary / First item: {title_snippet}")
        if has_table:
            print(f"Contains Table: {table_dims[0]} rows x {table_dims[1]} cols")

        # Specific slide checks
        if slide_num == 1:
            assert "Kuppam Engineering College" in full_slide_str or "KUPPAM ENGINEERING COLLEGE" in full_slide_str, "Slide 1 missing College name"
            assert "Mrs.K.Shandhini" in full_slide_str, "Slide 1 missing Guide name"
            assert "PAWANKALYAN" in full_slide_str and "RAMYA" in full_slide_str and "NANDINI" in full_slide_str and "DILLI RAJESH" in full_slide_str, "Slide 1 missing team members"
            assert "register" not in full_slide_str.lower() and "reg no" not in full_slide_str.lower(), "Slide 1 should NOT contain register numbers"
            assert "ARMEK – Agentic AI-Powered Loan Underwriting and Risk Assessment System" in full_slide_str, "Slide 1 missing title"
            print("  -> Slide 1 College Layout & Metadata verified OK!")

        elif slide_num == 4:
            assert has_table, "Slide 4 must contain a native PowerPoint table"
            assert table_dims == (4, 5), f"Slide 4 table should be 4x5, got {table_dims}"
            print("  -> Slide 4 Literature Survey Table (3 papers + header) verified OK!")

        elif slide_num == 6:
            assert "Advantages" in full_slide_str and "Disadvantages" in full_slide_str, "Slide 6 missing Advantages/Disadvantages"
            print("  -> Slide 6 Advantages & Disadvantages verified OK!")

        elif slide_num == 9:
            assert "Architecture" in full_slide_str or "ARCHITECTURE" in full_slide_str, "Slide 9 missing Architecture"
            assert "simulated" in full_slide_str.lower(), "Slide 9 must clarify external bureau simulation"
            print("  -> Slide 9 Architecture & Simulation clarification verified OK!")

        elif slide_num == 10:
            assert has_table, "Slide 10 must contain a native PowerPoint table"
            assert table_dims == (8, 5), f"Slide 10 table should be 8x5, got {table_dims}"
            print("  -> Slide 10 References Table (7 references + header) verified OK!")

        elif slide_num == 11:
            assert "Conclusion" in full_slide_str or "CONCLUSION" in full_slide_str, "Slide 11 missing Conclusion"
            print("  -> Slide 11 Conclusion verified OK!")

    print("\n=== Word Constraint Check ===")
    if prototype_matches:
        print(f"FAILED: Found {len(prototype_matches)} occurrences of 'prototype':")
        for s_idx, match in prototype_matches:
            print(f"  Slide {s_idx}: {match}")
    else:
        print("PASSED: Zero occurrences of 'prototype' found throughout all slides!")

    assert len(prototype_matches) == 0, "Presentation contains forbidden word 'prototype'"
    print("\n ALL 11 SLIDES VALIDATED PERFECTLY! Ready for college presentation review.")

if __name__ == "__main__":
    validate()

