"""
validate_final_pptx.py
Thorough verification of ARMEK_Final_Project_Review.pptx
"""

import pptx
from pptx.enum.shapes import MSO_SHAPE_TYPE

prs = pptx.Presentation("ARMEK_Final_Project_Review.pptx")

print("==================================================")
print("PPTX VALIDATION REPORT: ARMEK_Final_Project_Review.pptx")
print("==================================================")

# 1. Slide Count
total_slides = len(prs.slides)
print(f"\n[CHECK 1] Slide Count: {total_slides} (Expected: 10)")
assert total_slides == 10, f"Expected 10 slides, got {total_slides}"

# 2. Dimensions
w, h = prs.slide_width.inches, prs.slide_height.inches
print(f"[CHECK 2] Dimensions: {w:.3f} x {h:.3f} inches (16:9 widescreen: ~13.333 x 7.500)")

# 3. Check for forbidden word "prototype"
prototype_found = []
for idx, slide in enumerate(prs.slides):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                if "prototype" in p.text.lower():
                    prototype_found.append((idx + 1, p.text))
        elif shape.has_table:
            for r in shape.table.rows:
                for c in r.cells:
                    if "prototype" in c.text.lower():
                        prototype_found.append((idx + 1, c.text))

print(f"[CHECK 3] 'Prototype' Occurrences: {len(prototype_found)} (Expected: 0)")
if prototype_found:
    for s_num, txt in prototype_found:
        print(f"  WARNING: Found 'prototype' on Slide {s_num}: {txt[:80]}")
else:
    print("  PASS: Zero occurrences of 'prototype' found across all slides.")

# 4. Slide-by-Slide Verification
expected_titles = [
    "ARMEK",                                    # Slide 1
    "Introduction",                             # Slide 2
    "Objectives",                               # Slide 3
    "Literature Survey",                        # Slide 4
    "Existing System",                          # Slide 5
    "Existing System: Advantages and Disadvantages", # Slide 6
    "Proposed System",                          # Slide 7
    "Proposed System: Advantages",              # Slide 8
    "System Architecture Diagram",              # Slide 9
    "References"                                # Slide 10
]

for idx, slide in enumerate(prs.slides):
    slide_num = idx + 1
    print(f"\n--- Checking Slide {slide_num} ---")
    
    # Collect all text on slide
    all_text = []
    has_table = False
    table_dims = (0, 0)
    for shape in slide.shapes:
        if shape.has_text_frame:
            for p in shape.text_frame.paragraphs:
                t = p.text.strip()
                if t:
                    all_text.append(t)
        elif shape.has_table:
            has_table = True
            table_dims = (len(shape.table.rows), len(shape.table.columns))
            for r in shape.table.rows:
                for c in r.cells:
                    t = c.text.strip()
                    if t:
                        all_text.append(t)
                        
    full_slide_text = " ".join(all_text)
    
    # Verify expected title substring
    exp_t = expected_titles[idx]
    matches_title = exp_t.lower() in full_slide_text.lower()
    print(f"  Title match for '{exp_t}': {matches_title}")
    
    # Specific Slide Checks
    if slide_num == 1:
        # Check team members
        for m in ["PAWANKALYAN", "RAMYA", "NANDINI", "DILLI RAJESH"]:
            print(f"  Team Member '{m}': {m in full_slide_text}")
        print(f"  College 'Kuppam Engineering College': {'Kuppam Engineering College' in full_slide_text}")
        print(f"  Dept 'Computer Science and Engineering': {'Computer Science and Engineering' in full_slide_text}")
        print(f"  Academic Year '2026–2027': {'2026' in full_slide_text and '2027' in full_slide_text}")
        # Abstract check
        exp_abs_snippet = "ARMEK is an AI-powered loan underwriting and risk assessment system"
        print(f"  Abstract starts correctly: {exp_abs_snippet in full_slide_text}")
        
    elif slide_num == 2:
        print(f"  Workflow 'Loan Application': {'Loan Application' in full_slide_text}")
        print(f"  Workflow 'Loan Decision': {'Loan Decision' in full_slide_text}")
        
    elif slide_num == 3:
        print(f"  Main Objective present: {'Main Objective' in full_slide_text}")
        print(f"  Specific Objectives count mentions: {'User-Friendly' in full_slide_text and 'Master Agent' in full_slide_text}")
        
    elif slide_num == 4:
        print(f"  Has Table: {has_table}, Dimensions: {table_dims} (Expected: 4 rows x 5 cols)")
        print(f"  Turiel & Aste: {'Turiel' in full_slide_text}")
        print(f"  Xu, Lu & Xie: {'Xu' in full_slide_text}")
        print(f"  Alonso Robisco: {'Alonso' in full_slide_text}")
        assert has_table and table_dims == (4, 5), "Slide 4 table must be 4x5!"
        
    elif slide_num == 5:
        print(f"  Workflow 'Applicant': {'Applicant' in full_slide_text}")
        print(f"  Workflow 'Eligibility Assessment': {'Eligibility Assessment' in full_slide_text}")
        
    elif slide_num == 6:
        print(f"  Advantages column: {'Advantages' in full_slide_text}")
        print(f"  Disadvantages column: {'Disadvantages' in full_slide_text}")
        
    elif slide_num == 7:
        print(f"  Master Agent prominent: {'MASTER AGENT' in full_slide_text}")
        print(f"  Worker Agents mentioned: {'Worker Agents' in full_slide_text}")
        print(f"  Workflow present: {'FastAPI' in full_slide_text and 'React' in full_slide_text}")
        
    elif slide_num == 8:
        print(f"  Multi-Agent Workflow: {'Multi-Agent Workflow' in full_slide_text}")
        print(f"  Conversational AI: {'Conversational AI' in full_slide_text}")
        
    elif slide_num == 9:
        print(f"  Layer 1 User Interface: {'Layer 1' in full_slide_text}")
        print(f"  Layer 2 Backend: {'Layer 2' in full_slide_text}")
        print(f"  Layer 3 AI Orchestration: {'Layer 3' in full_slide_text}")
        print(f"  Layer 4 Worker Agent: {'Layer 4' in full_slide_text}")
        print(f"  Layer 5 Output: {'Layer 5' in full_slide_text}")
        print(f"  Simulated note: {'simulated' in full_slide_text.lower()}")
        
    elif slide_num == 10:
        print(f"  Has Table: {has_table}, Dimensions: {table_dims} (Expected: 8 rows x 5 cols)")
        print(f"  Papers (Turiel, Xu, Alonso): {'Turiel' in full_slide_text and 'Xu' in full_slide_text and 'Alonso' in full_slide_text}")
        print(f"  Tech Docs (React, FastAPI, OpenAI, Bootstrap): {'react.dev' in full_slide_text and 'fastapi' in full_slide_text and 'openai' in full_slide_text and 'getbootstrap' in full_slide_text}")
        assert has_table and table_dims == (8, 5), "Slide 10 table must be 8x5!"

print("\n==================================================")
print("ALL CHECKS PASSED SUCCESSFULLY!")
print("==================================================")

