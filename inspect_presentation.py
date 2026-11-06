import pptx
from pptx.enum.shapes import MSO_SHAPE_TYPE

prs = pptx.Presentation("ARMEK-Agentic-AI-Powered-Loan-Underwriting-and-Risk-Assessment-System (1).pptx")

with open("pptx_inspection_utf8.txt", "w", encoding="utf-8") as f:
    f.write(f"Total slides: {len(prs.slides)}\n")
    f.write(f"Dimensions: {prs.slide_width.inches:.3f} x {prs.slide_height.inches:.3f} inches\n\n")

    for i, slide in enumerate(prs.slides):
        f.write(f"\n==================== SLIDE {i+1} ====================\n")
        for s_idx, shape in enumerate(slide.shapes):
            f.write(f'  [Shape {s_idx+1}] Name: "{shape.name}", Type: {shape.shape_type}, Pos: ({shape.left.inches:.2f}, {shape.top.inches:.2f}), Size: ({shape.width.inches:.2f} x {shape.height.inches:.2f})\n')
            if shape.has_table:
                table = shape.table
                f.write(f"    Table: {len(table.rows)} rows x {len(table.columns)} cols\n")
                for r_idx, row in enumerate(table.rows):
                    row_texts = [cell.text.strip().replace("\n", " | ") for cell in row.cells]
                    f.write(f"      Row {r_idx+1}: {row_texts}\n")
            elif shape.has_text_frame:
                for p_idx, p in enumerate(shape.text_frame.paragraphs):
                    p_text = p.text.strip()
                    if p_text:
                        f.write(f"    P{p_idx+1}: {p_text}\n")
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                f.write(f"    Picture: format={shape.image.content_type}, size={shape.image.size}\n")
print("Done writing inspection log.")
