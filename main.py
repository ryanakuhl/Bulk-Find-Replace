from openpyxl import load_workbook
import pyperclip, re, easygui

class FirstTab:
    def __init__(self, row):
        self.current = row[0].value.rstrip().lstrip()
        self.replacement = row[1].value.rstrip().lstrip()

wb = load_workbook(easygui.fileopenbox(msg='Select file to edit', title='Select'))
sheet_ranges = wb['Sheet1']
replacements = []
for row in sheet_ranges.rows:
    x = FirstTab(row)
    replacements.append(x)

working_html = pyperclip.paste()
for r in replacements:
    working_html = re.sub(r.current, r.replacement, working_html, flags=re.IGNORECASE)
pyperclip.copy(working_html)
