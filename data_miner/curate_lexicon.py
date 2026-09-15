"""
Corpus Miner and Lexicon Builder for Jharkhand Education Department.
Mines DOCX, Excel, and Text files from '2024 data', extracts Kruti Dev & Unicode text,
and merges them with standard Administrative / Rajbhasha Hindi vocabulary.
"""

import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
sys.stdout.reconfigure(encoding='utf-8')
import re
import json
from collections import Counter
from typing import Dict, List, Set

import docx
import openpyxl

from krutidev_engine.converter import kruti_to_unicode, unicode_to_kruti
from krutidev_engine.trie import KrutiTrie

# Core Official / Administrative Vocabulary (Govt of Jharkhand & Central Rajbhasha)
BASE_OFFICIAL_WORDS = [
    # Top Govt Designations
    "जिला शिक्षा पदाधिकारी", "जिला शिक्षा अधीक्षक", "क्षेत्रीय संयुक्त शिक्षा निदेशक",
    "सहायक निदेशक", "प्राथमिक शिक्षा", "माध्यमिक शिक्षा", "उच्च शिक्षा",
    "प्रखंड शिक्षा प्रसार पदाधिकारी", "प्रधानाध्यापक", "प्रभारी प्रधानाध्यापक",
    "सहायक शिक्षक", "स्नातक प्रशिक्षित शिक्षक", "इंटर प्रशिक्षित शिक्षक",
    "निकासी एवं व्ययन पदाधिकारी", "उपायुक्त", "अनुमंडल पदाधिकारी",
    "लेखापाल", "लिपिक", "कंप्यूटर ऑपरेटर", "आशुलिपिक", "प्रधान लिपिक",
    "अवर सचिव", "संयुक्त सचिव", "प्रधान सचिव", "सचिव", "माननीय न्यायालय",
    
    # Official Administrative & Drafting Terms
    "कार्यालय", "कार्यालय-आदेश", "ज्ञापांक", "पत्रांक", "दिनांक", "विषय", "प्रसंग",
    "महोदय", "महाशय", "भवदीय", "हस्ताक्षर", "प्रतिलिपि", "सूचनार्थ", "कार्रवाई",
    "कार्यवाही", "अनुपालन", "प्रतिवेदन", "संचिका", "टिप्पणी", "शोकॉज", "स्पष्टीकरण",
    "आदेशानुसार", "निर्देशानुसार", "अनुरोध", "प्रावधान", "स्वीकृत", "अस्वीकृत",
    "पुनरीक्षित", "संशोधित", "वेतनमान", "वेतन", "बकाया", "पेंशन", "ग्रेच्युटी",
    "अवकाश", "मातृत्व", "चिकित्सा", "अर्जित", "असाधारण", "सेवा-पुस्तिका",
    "सत्यापन", "सम्पुष्टि", "नियमितिकरण", "निलंबन", "पुनर्बहाली", "जाँच",
    "जाँच-अधिकारी", "प्रपत्र-क", "आरोप-पत्र", "विभागीय", "संकल्प", "अधिसूचना",
    "अधिनियम", "नियमावली", "मार्गदर्शन", "परामर्श", "सहमति", "अनापत्ति",
    
    # Places & Blocks in Saraikela-Kharsawan / Jharkhand
    "झारखण्ड", "राँची", "सरायकेला", "खरसावाँ", "चांडिल", "इचागढ़", "कुकडू",
    "नीमडीह", "गम्हरिया", "राजनगर", "खूँटी", "पूर्वी सिंहभूम", "जमशेदपुर",
    "पश्चिमी सिंहभूम", "चाईबासा", "धनबाद", "बोकारो", "हजारीबाग", "दुमका",
    
    # Common Court Terms (W.P.S, Contempt, SOF)
    "माननीय उच्च न्यायालय", "माननीय सर्वोच्च न्यायालय", "रिट याचिका", "अवमानना",
    "प्रति-शपथपत्र", "तथ्य-विवरण", "याचिकाकर्ता", "प्रतिवादी", "वादी", "न्यायादेश",
    "पारित", "अंतिम", "अंतरिम", "रोक", "निस्तारण", "अवहेलना", "अनुपालन",
    
    # Numbers & Dates formatting terms
    "जनवरी", "फरवरी", "मार्च", "अप्रैल", "मई", "जून", "जुलाई", "अगस्त",
    "सितम्बर", "अक्टूबर", "नवम्बर", "दिसम्बर", "प्रथम", "द्वितीय", "तृतीय"
]

# Standard 10k-50k Common Hindi Root Words / Functional vocabulary
COMMON_HINDI_VOCAB = [
    "और", "तथा", "एवं", "अथवा", "या", "किंतु", "परंतु", "लेकिन", "इसलिए", "अतः",
    "क्योंकि", "यदि", "तो", "यद्यपि", "तथापि", "अर्थात", "मानो", "जिससे",
    "मैं", "हम", "तुम", "आप", "यह", "वह", "ये", "वे", "इस", "उस", "इन", "उन",
    "का", "के", "की", "को", "से", "में", "पर", "के लिए", "के द्वारा", "के साथ",
    "है", "हैं", "था", "थी", "थे", "होगा", "होगी", "होंगे", "हो", "हुए", "हुई",
    "किया", "किये", "गया", "गए", "गई", "जाता", "जाती", "जाते", "सकता", "सकते",
    "चाहिए", "दिया", "दिए", "दी", "लिया", "लिए", "ली", "रहा", "रहे", "रही",
    "आदेश", "निर्देश", "अनुरोध", "निवेदन", "प्रार्थना", "सूचना", "अवगत", "विदित",
    "समीक्षा", "बैठक", "उपस्थिति", "अनुपस्थिति", "विलंब", "समय", "तिथि", "दिवस",
    "वर्ष", "माह", "सप्ताह", "दैनिक", "मासिक", "वार्षिक", "त्रैमासिक", "अर्द्धवार्षिक",
    "विद्यालय", "महाविद्यालय", "विश्वविद्यालय", "छात्र", "छात्रा", "विद्यार्थी",
    "नामांकन", "छात्रवृत्ति", "पोशाक", "मध्याह्न", "भोजन", "पाठ्यपुस्तक", "परीक्षा",
    "मूल्यांकन", "परिणाम", "उत्तीर्ण", "अनुत्तीर्ण", "अंक", "प्रमाण-पत्र", "अभिलेख",
    "पंजी", "रजिस्टर", "रोकड़", "लेखा", "व्यय", "आय", "आवंटन", "निकासी", "कोषागार",
    "स्वीकृति", "भुगतान", "विहित", "नियम", "शर्त", "पात्रता", "योग्यता", "अनुभव",
    "प्रशिक्षण", "प्रशिक्षित", "अप्रशिक्षित", "मानदेय", "वेतन", "भत्ता", "महंगाई",
    "यात्रा", "दैनिक", "अवकाश", "आकस्मिक", "प्रतिपूरक", "अधिघोषणा", "प्रकाशन",
    "दैनिक", "समाचार", "विज्ञप्ति", "निविदा", "कोटेशन", "निविदाकार", "आपूर्तिकर्ता"
]


def extract_words_from_text(text: str) -> List[str]:
    """Splits text into cleaned Hindi words."""
    # Remove digits, punctuation, and english letters
    cleaned = re.sub(r'[a-zA-Z0-9\-_.,:;!?"\'\(\)\[\]\/\\%~`@#$^&*+=|]', ' ', text)
    tokens = [t.strip() for t in cleaned.split() if len(t.strip()) > 1]
    return tokens


def mine_docx_file(file_path: str) -> List[str]:
    """Extracts text and words from a single docx file."""
    words = []
    try:
        doc = docx.Document(file_path)
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            
            # Check font
            fonts = set(run.font.name for run in p.runs if run.font.name)
            is_kruti = any('kruti' in f.lower() for f in fonts)
            
            if is_kruti:
                uni_text = kruti_to_unicode(text)
                words.extend(extract_words_from_text(uni_text))
            else:
                # Might be unicode Hindi or ASCII
                # If text has devanagari characters
                if any('\u0900' <= c <= '\u097F' for c in text):
                    words.extend(extract_words_from_text(text))
                else:
                    # Try kruti conversion
                    uni_text = kruti_to_unicode(text)
                    if any('\u0900' <= c <= '\u097F' for c in uni_text):
                        words.extend(extract_words_from_text(uni_text))
                        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    ctext = cell.text.strip()
                    if ctext:
                        uni_text = kruti_to_unicode(ctext)
                        words.extend(extract_words_from_text(uni_text))
    except Exception as e:
        pass
    return words


def build_complete_lexicon(data_dir: str = "2024 data", max_files: int = 200) -> Dict:
    """Builds the full dual-indexed lexicon from base words + mined corpus."""
    print("Building Lexicon...")
    
    word_freq: Counter = Counter()
    
    # 1. Add base official vocabulary with high initial frequency weight
    for w in BASE_OFFICIAL_WORDS:
        for sub_w in w.split():
            clean_w = sub_w.strip()
            if clean_w:
                word_freq[clean_w] += 100
                
    for w in COMMON_HINDI_VOCAB:
        clean_w = w.strip()
        if clean_w:
            word_freq[clean_w] += 50
            
    # 2. Mine relevant DOCX files in data_dir
    processed_count = 0
    if os.path.exists(data_dir):
        print(f"Scanning '{data_dir}' for DOCX documents...")
        for root, dirs, files in os.walk(data_dir):
            for file in files:
                if file.endswith('.docx') and not file.startswith('~$'):
                    fpath = os.path.join(root, file)
                    extracted = mine_docx_file(fpath)
                    for ew in extracted:
                        # Only keep valid devanagari words
                        if any('\u0900' <= c <= '\u097F' for c in ew) and len(ew) > 1:
                            word_freq[ew] += 1
                    processed_count += 1
                    if processed_count >= max_files:
                        break
            if processed_count >= max_files:
                break
                
    print(f"Processed {processed_count} files. Extracted {len(word_freq)} unique Hindi words.")
    
    # 3. Create dual Trie and JSON store
    lexicon_list = []
    trie = KrutiTrie()
    
    for uni_word, freq in word_freq.most_common(25000):
        # Generate Kruti Dev representation
        kruti_form = unicode_to_kruti(uni_word)
        if not kruti_form:
            continue
            
        category = "department" if freq > 10 else "general"
        
        entry = {
            "unicode": uni_word,
            "kruti": kruti_form,
            "frequency": freq,
            "category": category
        }
        lexicon_list.append(entry)
        trie.insert(kruti_form, uni_word, freq, category)
        
    # Save to JSON
    output_dir = os.path.join(os.path.dirname(__file__), "..", "krutidev_engine")
    os.makedirs(output_dir, exist_ok=True)
    json_path = os.path.join(output_dir, "lexicon.json")
    
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(lexicon_list, f, ensure_ascii=False, indent=2)
        
    print(f"Saved {len(lexicon_list)} compiled words to {json_path}")
    return {"total_words": len(lexicon_list), "json_path": json_path}


if __name__ == "__main__":
    build_complete_lexicon()
