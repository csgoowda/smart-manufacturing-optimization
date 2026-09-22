import os
from pathlib import Path
import win32com.client
import json

def get_exact_pages():
    docx_path = Path("Smart_Manufacturing_Optimization_Final_Report.docx").resolve()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    
    doc = word.Documents.Open(str(docx_path))
    doc.Repaginate()
    
    pages = {}
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        txt = p.Range.Text.strip()
        if not txt:
            continue
        pg = p.Range.Information(3) # wdActiveEndPageNumber
        
        # Check headings
        if txt.startswith("CHAPTER") or txt.startswith("1.") or txt.startswith("2.") or txt.startswith("3.") or txt.startswith("4.") or txt in ["REFERENCES", "APPENDICES", "CONTENTS", "ABSTRACT", "LIST OF FIGURES", "LIST OF TABLES", "ACCEPTANCE LETTER", "CERTIFICATE", "ACKNOWLEDGEMENT"]:
            key = txt[:50]
            if key not in pages:
                pages[key] = pg
        elif txt.startswith("Figure 3."):
            num = txt.split(":")[0].strip()
            pages[num] = pg
        elif txt.startswith("Table 3.") or txt.startswith("Table C."):
            num = txt.split(":")[0].strip()
            pages[num] = pg
            
    doc.Close(SaveChanges=False)
    word.Quit()
    
    with open("exact_pages.json", "w", encoding="utf-8") as f:
        json.dump(pages, f, indent=2)
    print(f"Extracted {len(pages)} items to exact_pages.json")

if __name__ == "__main__":
    get_exact_pages()
