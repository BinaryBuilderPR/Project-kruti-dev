"""
Deep Corpus Analysis & Department Profiler for '2024 data'.
Categorizes all files across all 23 folders/desks, extracts drafting templates,
vocabulary, court styles, administrative patterns, and error analysis.
"""

import os
import sys
import re
import json
from collections import Counter, defaultdict

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
sys.stdout.reconfigure(encoding='utf-8')

import docx
import openpyxl
from pypdf import PdfReader

from krutidev_engine.converter import kruti_to_unicode, unicode_to_kruti


DATA_DIR = "2024 data"

# Clean Hindi word validator
DEV_WORD_PATTERN = re.compile(r'^[\u0900-\u097F]+$')
MATRA_START = re.compile(r'^[\u093E-\u094C\u0901-\u0903\u094D]')


def is_clean_hindi(word: str) -> bool:
    if len(word) <= 1:
        return False
    if not DEV_WORD_PATTERN.match(word):
        return False
    if MATRA_START.match(word):
        return False
    return True


def extract_docx(fpath: str) -> str:
    text_chunks = []
    try:
        doc = docx.Document(fpath)
        for p in doc.paragraphs:
            t = p.text.strip()
            if not t:
                continue
            # Check if font is Kruti
            fonts = set(run.font.name for run in p.runs if run.font.name)
            is_kruti = any('kruti' in str(f).lower() for f in fonts)
            if is_kruti:
                text_chunks.append(kruti_to_unicode(t))
            elif any('\u0900' <= c <= '\u097F' for c in t):
                text_chunks.append(t)
            else:
                text_chunks.append(kruti_to_unicode(t))

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    ct = cell.text.strip()
                    if ct:
                        if any('\u0900' <= c <= '\u097F' for c in ct):
                            text_chunks.append(ct)
                        else:
                            text_chunks.append(kruti_to_unicode(ct))
    except Exception:
        pass
    return "\n".join(text_chunks)


def extract_pdf(fpath: str, max_pages: int = 8) -> str:
    text_chunks = []
    try:
        reader = PdfReader(fpath)
        for page in reader.pages[:max_pages]:
            pt = page.extract_text()
            if pt:
                if any('\u0900' <= c <= '\u097F' for c in pt):
                    text_chunks.append(pt)
                else:
                    text_chunks.append(kruti_to_unicode(pt))
    except Exception:
        pass
    return "\n".join(text_chunks)


def extract_excel(fpath: str, max_rows: int = 50) -> str:
    text_chunks = []
    try:
        wb = openpyxl.load_workbook(fpath, data_only=True, read_only=True)
        for sheet in wb.sheetnames[:2]:
            ws = wb[sheet]
            r = 0
            for row in ws.iter_rows(values_only=True):
                r += 1
                if r > max_rows:
                    break
                for val in row:
                    if isinstance(val, str) and val.strip():
                        if any('\u0900' <= c <= '\u097F' for c in val):
                            text_chunks.append(val.strip())
                        else:
                            text_chunks.append(kruti_to_unicode(val.strip()))
        wb.close()
    except Exception:
        pass
    return "\n".join(text_chunks)


def analyze_corpus():
    print(f"=== Starting Deep Corpus Profiling across '{DATA_DIR}' ===")

    folder_stats = defaultdict(lambda: {"docx": 0, "pdf": 0, "xlsx": 0, "txt": 0, "total": 0, "sample_titles": []})
    
    category_counts = Counter()
    
    # Document patterns
    court_cases = []
    establishment_records = []
    financial_records = []
    routine_letters = []
    
    frequent_subjects = Counter()
    frequent_openings = Counter()
    frequent_closings = Counter()
    frequent_designations = Counter()
    frequent_places = Counter()
    
    all_words = Counter()
    
    total_files = 0
    
    for root, dirs, files in os.walk(DATA_DIR):
        folder_name = os.path.relpath(root, DATA_DIR).split(os.sep)[0]
        if folder_name == ".":
            folder_name = "ROOT"
            
        for file in files:
            if file.startswith("~$"):
                continue
            ext = os.path.splitext(file)[1].lower()
            fpath = os.path.join(root, file)
            
            if ext not in ('.docx', '.pdf', '.xlsx', '.xls', '.txt'):
                continue
                
            total_files += 1
            st = folder_stats[folder_name]
            st["total"] += 1
            if len(st["sample_titles"]) < 5:
                st["sample_titles"].append(file)
                
            text = ""
            if ext == '.docx':
                st["docx"] += 1
                text = extract_docx(fpath)
            elif ext == '.pdf':
                st["pdf"] += 1
                text = extract_pdf(fpath)
            elif ext in ('.xlsx', '.xls'):
                st["xlsx"] += 1
                text = extract_excel(fpath)
            elif ext == '.txt':
                st["txt"] += 1
                try:
                    with open(fpath, 'r', encoding='utf-8', errors='ignore') as f:
                        text = kruti_to_unicode(f.read())
                except:
                    pass
                    
            if not text:
                continue
                
            # Classify document
            text_lower = text.lower()
            file_lower = file.lower()
            
            is_court = any(k in text or k in file_lower for k in ["wps", "न्यायालय", "याचिका", "वादी", "प्रतिवादी", "शपथपत्र", "अंतरिम", "अवहेलना", "contempt", "wp(s)", "court", "उच्च न्यायालय"])
            is_est = any(k in text or k in file_lower for k in ["स्थापना", "प्रोन्नति", "नियुक्ति", "वेतनमान", "ग्रेड", "वरीयता", "सत्यापन", "सहायक आचार्य", "शिक्षिका", "प्रधानाध्यापक", "लिपिक"])
            is_fin = any(k in text or k in file_lower for k in ["आवंटन", "वेतन", "रोकड़", "लेखा", "कोषागार", "पेंशन", "ग्रेच्युटी", "बजट", "mdm", "माह", "व्यय", "निकासी एवं व्ययन"])
            is_rti = any(k in text or k in file_lower for k in ["सूचना का अधिकार", "विधानसभा", "तारांकित", "अतारांकित", "rti", "bidhansabha"])
            
            if is_court:
                category_counts["Court / Legal Cases"] += 1
                if len(court_cases) < 15 and len(text) > 100:
                    court_cases.append({"file": file, "folder": folder_name, "excerpt": text[:300].strip()})
            elif is_est:
                category_counts["Establishment & Service"] += 1
                if len(establishment_records) < 15 and len(text) > 100:
                    establishment_records.append({"file": file, "folder": folder_name, "excerpt": text[:300].strip()})
            elif is_fin:
                category_counts["Finance, Accounts & Allotments"] += 1
                if len(financial_records) < 15 and len(text) > 100:
                    financial_records.append({"file": file, "folder": folder_name, "excerpt": text[:300].strip()})
            elif is_rti:
                category_counts["RTI & Legislative (Vidhansabha)"] += 1
            else:
                category_counts["General Official Correspondence"] += 1
                if len(routine_letters) < 15 and len(text) > 100:
                    routine_letters.append({"file": file, "folder": folder_name, "excerpt": text[:300].strip()})
                    
            # Parse lines for administrative phrasing
            lines = [l.strip() for l in text.split('\n') if l.strip()]
            for line in lines:
                # Subjects
                if line.startswith("विषय") or "विषय -" in line or "विषय %" in line or "विषयः" in line:
                    subj = re.sub(r'^विषय[\s\-\:\%]+', '', line).strip()
                    if len(subj) > 5:
                        frequent_subjects[subj[:100]] += 1
                        
                # Openings
                if any(k in line for k in ["महाशय", "महोदय", "उपर्युक्त विषयक", "सादर सूचित करना", "निर्देशानुसार सूचित", "आलोक में", "अनुरोध है कि"]):
                    frequent_openings[line[:90]] += 1
                    
                # Closings
                if any(k in line for k in ["विश्वासभाजन", "भवदीय", "आदेशानुसार", "हस्ताक्षर", "प्रतिलिपि", "ज्ञापांक"]):
                    frequent_closings[line[:90]] += 1
                    
                # Designations
                for d in ["जिला शिक्षा अधीक्षक", "क्षेत्रीय शिक्षा उप निदेशक", "उपायुक्त", "निदेशक", "सचिव", "प्रखंड शिक्षा प्रसार पदाधिकारी", "प्रधानाध्यापक", "कनीय लिपिक", "प्रधान लिपिक", "लेखापाल"]:
                    if d in line:
                        frequent_designations[d] += 1
                        
                # Places & Blocks
                for pl in ["सरायकेला", "खरसावाँ", "चांडिल", "इचागढ़", "कुकडू", "राजनगर", "गम्हरिया", "खूँटी", "राँची", "चाईबासा", "जमशेदपुर", "झारखण्ड"]:
                    if pl in line:
                        frequent_places[pl] += 1
                        
                # Words
                words = re.findall(r'[\u0900-\u097F]+', line)
                for w in words:
                    if is_clean_hindi(w):
                        all_words[w] += 1

    summary_report = {
        "total_files_scanned": total_files,
        "folder_breakdown": {k: v for k, v in sorted(folder_stats.items(), key=lambda x: x[1]["total"], reverse=True)},
        "categories_distribution": dict(category_counts.most_common()),
        "top_places_and_blocks": dict(frequent_places.most_common(20)),
        "top_designations": dict(frequent_designations.most_common(15)),
        "top_frequent_subjects": [s for s, _ in frequent_subjects.most_common(25)],
        "top_frequent_openings": [o for o, _ in frequent_openings.most_common(20)],
        "top_frequent_closings": [c for c, _ in frequent_closings.most_common(15)],
        "sample_court_cases": court_cases[:5],
        "sample_establishment": establishment_records[:5],
        "sample_financial": financial_records[:5],
        "sample_routine": routine_letters[:5],
        "total_unique_vocabulary": len(all_words),
        "top_50_vocabulary": [w for w, _ in all_words.most_common(50)]
    }
    
    out_file = os.path.join(os.path.dirname(__file__), "..", "corpus_deep_summary.json")
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(summary_report, f, ensure_ascii=False, indent=2)
        
    print(f"\nAnalysis complete! Processed {total_files} files across all desks.")
    print(f"Summary saved to '{out_file}'.")
    return summary_report


if __name__ == "__main__":
    analyze_corpus()

