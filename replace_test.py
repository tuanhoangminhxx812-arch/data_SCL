import docx

doc = docx.Document("Mẫu TMQT.docx")
for p in doc.paragraphs:
    if "- Tên danh mục:" in p.text:
        p.text = p.text.replace("- Tên danh mục:", f"- Tên danh mục: Công trình Test")
    if "- Giá trị vốn kế hoạch:" in p.text:
        p.text = p.text.replace("...........", "1,500")

doc.save("Test_TMQT.docx")
