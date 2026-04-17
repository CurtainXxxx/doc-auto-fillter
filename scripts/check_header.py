"""检查表格1行17表头"""
from docx import Document
doc = Document("assets/template.docx")
t1 = doc.tables[1]
row = t1.rows[17]
seen = set()
for c_idx, cell in enumerate(row.cells):
    cid = id(cell._element)
    if cid in seen:
        continue
    seen.add(cid)
    print(f"  列{c_idx}: '{cell.text.strip()[:60].replace(chr(10), '|')}'")
