import sys
lines = open("app.py", "r", encoding="utf-8").readlines()
with open("find_lines_out2.txt", "w", encoding="utf-8") as f:
    for i, line in enumerate(lines):
        if "thẩm tra" in line.lower() or "quyet dinh" in line.lower() or "quyết định" in line.lower() or "phiếu" in line.lower():
            f.write(f"{i+1}: {line.strip()}\n")
