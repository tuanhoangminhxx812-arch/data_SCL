with open('app.py', 'r', encoding='utf-8') as f:
    text = f.read()

text = text.replace('ws[\'D8\'] = f"Công trình: {ct}"', 'ws[\'D8\'] = ct')
text = text.replace('ws[\'D8\'] = f"Công trình: {selected_ct}" ', 'ws[\'D8\'] = selected_ct')

with open('app.py', 'w', encoding='utf-8') as f:
    f.write(text)
