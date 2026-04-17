"""调试行组填充逻辑"""
from docx import Document
import re

doc = Document("assets/template.docx")
t = doc.tables[1]

# 检查行18列1的具体内容
cell = t.rows[18].cells[1]
text = cell.text.strip().replace('\r', '')
print(f"原始文本: repr={repr(text)}")
print(f"按换行分割: {text.split(chr(10))}")

for line in text.split('\n'):
    line = line.strip()
    print(f"  line='{line}'")
    match = re.match(r'^(.+?)\s*[：:]\s*(.*)', line)
    if match:
        print(f"    match! label='{match.group(1)}' value='{match.group(2)}'")
    else:
        print(f"    no match")

# 检查独立单元格
seen = set()
for ci, c in enumerate(t.rows[18].cells):
    cid = id(c._element)
    if cid not in seen:
        seen.add(cid)
        print(f"独立格{ci}: '{c.text.strip()[:50]}'")

# 模拟行组填充: row_data = ["课程目标1", "闭卷笔试", "百分制", "70", "62.3", "0.89"]
# 独立单元格: [空, "实现途径:\n评价方法:", "期末考试", 空, 空, 空]
# 映射:
# 空 → row_data[0]="课程目标1" → set_cell
# "实现途径:\n评价方法:" → 两个标签:
#   "实现途径:" 后为空 → row_data[1]="闭卷笔试" → append
#   "评价方法:" 后为空 → row_data[2]="百分制" → append
# "期末考试" → 无冒号，跳过
# 空 → row_data[3]="70" → set_cell
# 空 → row_data[4]="62.3" → set_cell
# 空 → row_data[5]="0.89" → set_cell
# 所以 data_idx 最终应该是 6 (0-5)
