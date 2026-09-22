import os
from pathlib import Path
import win32com.client

def repaginate_and_inspect():
    docx_path = Path("Smart_Manufacturing_Optimization_Final_Report.docx").resolve()
    print(f"Inspecting via Word: {docx_path}")
    
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = False
        
        doc = word.Documents.Open(str(docx_path))
        
        # Update all fields in document
        print("Updating fields...")
        doc.Fields.Update()
        for s in doc.Sections:
            s.Headers(1).Range.Fields.Update()
            s.Footers(1).Range.Fields.Update()
            
        doc.Repaginate()
        page_count = doc.ComputeStatistics(2) # 2 = wdStatisticPages
        print(f"Total Pages in Document: {page_count}")
        
        # Find page numbers of each major heading
        print("\n--- HEADING PAGE NUMBERS ---")
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            txt = p.Range.Text.strip()
            if txt.startswith("CHAPTER") or txt in ["CONTENTS", "ABSTRACT", "LIST OF FIGURES", "LIST OF TABLES", "REFERENCES", "APPENDICES"]:
                pg = p.Range.Information(3) # 3 = wdActiveEndPageNumber
                print(f"'{txt[:40]}' -> Page {pg}")
            elif txt.startswith("Figure 3.") or txt.startswith("Table 3."):
                pg = p.Range.Information(3)
                print(f"'{txt[:40]}' -> Page {pg}")
                
        doc.Save()
        doc.Close(SaveChanges=True)
        print("Document saved successfully via Word.")
        
    except Exception as e:
        print("Word automation error:", e)
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass

if __name__ == "__main__":
    repaginate_and_inspect()
