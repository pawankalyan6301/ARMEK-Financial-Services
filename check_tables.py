import pptx

prs = pptx.Presentation("ARMEK_Final_Project_Review.pptx")

print("=== SLIDE 4 TABLE ===")
s4 = prs.slides[3]
for sh in s4.shapes:
    if sh.has_table:
        t = sh.table
        for r_idx, r in enumerate(t.rows):
            print(f"Row {r_idx}: {[c.text.replace(chr(10), ' | ') for c in r.cells]}")

print("\n=== SLIDE 10 TABLE ===")
s10 = prs.slides[9]
for sh in s10.shapes:
    if sh.has_table:
        t = sh.table
        for r_idx, r in enumerate(t.rows):
            print(f"Row {r_idx}: {[c.text.replace(chr(10), ' | ') for c in r.cells]}")

