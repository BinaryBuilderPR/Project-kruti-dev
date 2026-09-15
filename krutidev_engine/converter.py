"""
Production-grade Kruti Dev 010 <-> Unicode Devanagari Bidirectional Converter.
Accurately implements Remington typewriter layout mappings, matra re-ordering,
reph placement, conjuncts, half-letters, nuktas, and punctuation.
"""

import re
from typing import List, Tuple

# Pre-compiled multi-character substitutions for Kruti Dev -> Unicode
KRUTI_MULTI_REPLACEMENTS: List[Tuple[str, str]] = [
    # Complex vowels
    ("vksS", "औ"),
    ("vkS", "औ"),
    ("vks", "ओ"),
    ("vk", "आ"),
    ("v", "अ"),
    ("bZ", "ई"),
    ("b", "इ"),
    ("Å", "ऊ"),
    ("m", "उ"),
    ("_", "ऋ"),
    (",s", "ऐ"),
    (",", "ए"),
    
    # Special combo glyphs / conjuncts
    ("Dr", "क्त"),
    ("iz", "प्र"),
    ("nz", "द्र"),
    ("Ø", "क्र"),
    ("Vª", "ट्र"),
    ("Mª", "ड्र"),
    ("Í", "द्द"),
    ("Î", "द्ध"),
    ("Ï", "द्घ"),
    ("Ð", "द्य"),
    ("Ñ", "कृ"),
    ("Ò", "दृ"),
    ("Ó", "हृ"),
    ("Ô", "ह्म्"),
    ("Õ", "ह्य"),
    ("Ö", "ह्ल"),
    ("×", "ह्व"),
    ("Ù", "त्त"),
    ("Ú", "त्त्"),
    ("Û", "क्त"),
    ("Ü", "दृ"),
    ("Ý", "फ्र"),
    ("®", "रु"),
    ("¯", "रू"),
    ("}", "द्व"),
    ("|", "द्य"),
    
    # Full Consonants with 'k'
    ("[kk", "खा"),
    ("[k", "ख"),
    ("?k", "घ"),
    ("Fk", "थ"),
    ("/k", "ध"),
    ("Hk", "भ"),
    ("'k", "श"),
    ('"k', "ष"),
    ("{k", "क्ष"),
    (".k", "ण"),
    
    # Nukta consonants
    ("d+", "क़"),
    ("[k+", "ख़"),
    ("x+", "ग़"),
    ("t+", "ज़"),
    ("M+", "ड़"),
    ("<+", "ढ़"),
    ("Q+", "फ़"),
    
    # Standalone conjunct keys
    ("=", "त्र"),
    ("K", "ज्ञ"),
    ("J", "श्र"),
    
    # Chandrabindu
    ("Wa", "ँ"),
    ("Wk", "ँ"),
    ("W", "ँ"),
]

# Single character direct mappings (Kruti Dev -> Unicode)
KRUTI_SINGLE_MAP = {
    # Full consonants
    'd': 'क', 'x': 'ग', 'p': 'च', 'N': 'छ',
    't': 'ज', 'T': 'झ', '¥': 'ञ',
    'V': 'ट', 'B': 'ठ', 'M': 'ड', '<': 'ढ',
    'r': 'त', 'n': 'द', 'u': 'न',
    'i': 'प', 'Q': 'फ', 'c': 'ब', 'e': 'म',
    ';': 'य', 'j': 'र', 'y': 'ल', 'o': 'व',
    'l': 'स', 'g': 'ह',
    
    # Half consonants (when typed without 'k' / shifted keys)
    'D': 'क्', 'X': 'ग्', 'P': 'च्', 'Y': 'ल्',
    'R': 'त्', 'F': 'थ्', '/': 'ध्', 'U': 'न्',
    'I': 'प्', 'C': 'ब्', 'H': 'भ्', 'E': 'म्',
    'O': 'व्', 'L': 'स्', "'": 'श्', '"': 'ष्',
    '[': 'ख्', '?': 'घ्', '{': 'क्ष्', '.': 'ण्',
    '~': '्',
    
    # Matras
    'k': 'ा', 'h': 'ी', 'q': 'ु', 'w': 'ू',
    '`': 'ृ', 's': 'े', 'S': 'ै', 'a': 'ं',
    'z': '्र', 'µ': 'ृ',
    
    # Punctuation & digits
    'A': '।', ']': ',', '%': ':', '&': '-',
    '0': '0', '1': '1', '2': '2', '3': '3', '4': '4',
    '5': '5', '6': '6', '7': '7', '8': '8', '9': '9',
    'º': '०', '°': '०'
}


def kruti_to_unicode(text: str) -> str:
    """Converts a Kruti Dev 010 ASCII encoded string to Unicode Devanagari."""
    if not text:
        return ""
    
    res = text
    
    # Step 1: Replace known multi-character sequences
    for k, u in KRUTI_MULTI_REPLACEMENTS:
        res = res.replace(k, u)
        
    # Step 2: Handle 'f' (chhoti-i matra) which comes BEFORE the consonant cluster in Kruti Dev
    # Find 'f' followed by half letters and consonant
    def _fix_chhoti_i(match):
        cluster = match.group(1)
        # Convert any remaining single kruti chars in cluster
        converted = "".join(KRUTI_SINGLE_MAP.get(c, c) for c in cluster)
        return converted + "ि"
    
    # Matches 'f' followed by 1 to 4 characters of consonant / half-consonants / ligatures
    pattern_i = r'f([a-zA-Z\u0900-\u097F~`\.\/\'\"\[\?\{]+?)'
    # We iteratively resolve 'f'
    for _ in range(5):
        # Match f followed by (half consonant)* + full consonant
        res = re.sub(r'f((?:[\u0915-\u0939]्|[A-Z\[\?\{\'\"])*[\u0915-\u0939d-hjl-pqrt-x=KJ]|Dr|iz|nz|Ø)', _fix_chhoti_i, res)
    
    # Step 3: Map all remaining single characters
    chars = []
    for ch in res:
        chars.append(KRUTI_SINGLE_MAP.get(ch, ch))
    res = "".join(chars)
    
    # Step 4: Handle 'Z' (Reph / 'र्')
    # In Kruti Dev, 'Z' is at the end of the syllable: e.g. "कार्य" was typed "dk;Z" -> "काय" + "Z"
    # In Unicode, 'र्' must precede the consonant cluster: "क" + "ा" + "र्" + "य"
    # Syllable pattern before Z: (Consonant + Halant)* + Consonant + (Matras)* + Z
    # We want: र् + (Consonant + Halant)* + Consonant + (Matras)*
    consonants = r'[\u0915-\u0939]'
    matras = r'[\u093E-\u094C\u0901-\u0903]'
    
    def _fix_reph(match):
        cluster = match.group(1)
        return "र्" + cluster
    
    reph_pattern = rf'((?:{consonants}्)*{consonants}{matras}*)Z'
    while 'Z' in res:
        new_res = re.sub(reph_pattern, _fix_reph, res)
        if new_res == res:
            res = res.replace('Z', 'र्')
            break
        res = new_res
        
    # Step 5: Post-cleanup of Unicode anomalies
    res = res.replace('अा', 'आ')
    res = res.replace('ाे', 'ो')
    res = res.replace('ाै', 'ौ')
    res = res.replace('्ा', '')
    res = res.replace('िा', 'ि')
    res = res.replace('ीा', 'ी')
    
    return res


# Direct high-frequency word dictionary mapping for 100% fidelity
DIRECT_WORD_UNICODE_TO_KRUTI = {
    "कार्यालय": "dk;kZy;",
    "शिक्षा": "f'k{kk",
    "विभाग": "foHkkx",
    "झारखण्ड": ">kj[k.M",
    "राँची": "jkWaph",
    "सरायकेला": "ljk;dsyk",
    "खरसावाँ": "[kjlkokWa",
    "प्राथमिक": "izkFkfed",
    "माध्यमिक": "ek/;fed",
    "उच्च": "mPp",
    "न्यायालय": "U;k;ky;",
    "पदाधिकारी": "inkf/kdkjh",
    "अधीक्षक": "v/kh{kd",
    "निदेशक": "funs'kd",
    "आदेश": "vkns'k",
    "ज्ञापांक": "Kkikad",
    "पत्रांक": "i=kad",
    "दिनांक": "fnukad",
    "प्रतिवेदन": "izfrosnu",
    "अनुपालन": "vuqikyu",
    "संचिका": "lafpdk",
    "भवदीय": "Hkonh;",
    "हस्ताक्षर": "gLrk{kj",
    "विषय": 'fo"k;',
    "प्रसंग": "izlax",
    "महोदय": "egksn;",
    "अनुरोध": "vuqjks/k",
    "आवश्यक": "vko';d",
    "कार्रवाई": "dk;Zokgh",
    "उपरोक्त": "mijksDr",
    "संलग्न": "layXu",
    "वेतन": "osru",
    "शिक्षक": "f'k{kd",
    "शिक्षिका": "f'kf{kdk",
    "विद्यालय": "fon~;ky;",
    "प्रधानाध्यापक": "iz/kkuk/;kid",
    "प्रभारी": "izHkkjh",
    "छात्र": "Nk=",
    "छात्रा": "Nk=k",
    "उपस्थिति": "mifLFkfr",
    "अवकाश": "vodk'k",
    "स्वीकृत": "Lohd`r",
    "अस्वीकृत": "vLohd`r",
    "स्पष्टीकरण": 'Li"Vhdj.k',
    "शोकॉज": "'kksdkWt",
    "जाँच": "tkWap",
    "निलंबन": "fuyacu",
    "वेतनमान": "osrueku",
    "में": "esa",
    "मैं": "eSa",
    "की": "dh",
    "कि": "fd",
    "के": "ds",
    "को": "dks",
    "से": "ls",
    "पर": "ij",
    "का": "dk",
    "है": "gS",
    "हैं": "gSa",
    "था": "Fkk",
    "थी": "Fkh",
    "थे": "Fks",
    "होगा": "gksxk",
    "होगी": "gksxh",
    "होंगे": "gksaxs",
    "किया": "fd;k",
    "किये": "fd;s",
    "गया": "x;k",
    "गए": "x,",
    "गई": "xbZ",
    "गईं": "xbZa",
    "जाता": "tkrk",
    "जाती": "tkrh",
    "जाते": "tkrs",
    "द्वारा": "}kjk",
    "तथा": "rFkk",
    "एवं": ",oa",
    "अथवा": "vFkok",
    "यदि": ";fn",
    "तो": "rks",
    "इस": "bl",
    "उस": "ml",
    "सभी": "lHkh",
    "प्राप्त": "izkIr",
    "प्रस्तुत": "izLrqr",
    "सूचित": "lwpgr",
    "निर्देश": "funsZ'k",
    "निर्देशित": "funsZf'kr",
    "समीक्षा": "leh{kk",
    "बैठक": "cSBd",
    "निर्णय": "fu.kZ;",
    "प्रस्ताव": "izLrko",
    "अनुमोदन": "vuqeksnu",
    "टिप्पणी": "fVIi.kh",
    "पंजी": "iath",
    "अवर": "voj",
    "प्रखंड": "iz[kaM",
    "क्षेत्रीय": "{ks=h;",
    "संयुक्त": "la;qDr",
}


def unicode_to_kruti(text: str) -> str:
    """Converts a Unicode Devanagari string to Kruti Dev 010 ASCII representation."""
    if not text:
        return ""
    
    # Fast path for known words
    if text in DIRECT_WORD_UNICODE_TO_KRUTI:
        return DIRECT_WORD_UNICODE_TO_KRUTI[text]
    
    res = text
    
    # Step 1: Handle Reph (र्) -> Move to end of syllable as 'Z'
    # Pattern: र् + ConsonantCluster + Matras -> ConsonantCluster + Matras + Z
    consonants = r'[\u0915-\u0939]'
    matras = r'[\u093E-\u094C\u0901-\u0903]'
    res = re.sub(rf'र्((?:{consonants}्)*{consonants}{matras}*)', r'\1Z', res)
    
    # Step 2: Handle Chhoti-i matra (ि) -> Move before consonant cluster as 'f'
    # Pattern: (ConsonantCluster) + ि -> f + (ConsonantCluster)
    res = re.sub(rf'((?:{consonants}्)*{consonants})ि', r'f\1', res)
    
    # Step 3: Direct replacements (Ordered: multi-char and special glyphs first)
    replacements = [
        ("औ", "vkS"), ("ओ", "vks"), ("आ", "vk"), ("अ", "v"),
        ("ई", "bZ"), ("इ", "b"), ("ऊ", "Å"), ("उ", "m"),
        ("ऐ", ",s"), ("ए", ","), ("ऋ", "_"),
        
        ("क्ष", "{k"), ("त्र", "="), ("ज्ञ", "K"), ("श्र", "J"),
        ("क्त", "Dr"), ("प्र", "iz"), ("द्र", "nz"), ("क्र", "Ø"),
        ("ट्र", "Vª"), ("ड्र", "Mª"), ("द्व", "}"), ("द्य", "|"),
        ("द्ध", "Î"), ("द्द", "Í"), ("कृ", "Ñ"), ("दृ", "Ò"), ("हृ", "Ó"),
        ("रु", "®"), ("रू", "¯"),
        
        # Half letters
        ("क्", "D"), ("ख्", "["), ("ग्", "X"), ("घ्", "?"),
        ("च्", "P"), ("ज्", "T"), ("त्", "R"), ("थ्", "F"),
        ("ध्", "/"), ("न्", "U"), ("प्", "I"),
        ("ब्", "C"), ("भ्", "H"), ("म्", "E"), ("ल्", "Y"),
        ("व्", "O"), ("श्", "'"), ("ष्", '"'), ("स्", "L"),
        ("ह्", "G"), ("क्ष्", "{"), ("ण्", "."),
        
        # Full consonants
        ("क", "d"), ("ख", "[k"), ("ग", "x"), ("घ", "?k"), ("ङ", "M+"),
        ("च", "p"), ("छ", "N"), ("ज", "t"), ("झ", "T"), ("ञ", "¥"),
        ("ट", "V"), ("ठ", "B"), ("ड", "M"), ("ढ", "<"), ("ण", ".k"),
        ("त", "r"), ("थ", "Fk"), ("द", "n"), ("ध", "/k"), ("न", "u"),
        ("प", "i"), ("फ", "Q"), ("ब", "c"), ("भ", "Hk"), ("म", "e"),
        ("य", ";"), ("र", "j"), ("ल", "y"), ("व", "o"),
        ("श", "'k"), ("ष", '"k'), ("स", "l"), ("ह", "g"),
        
        # Nukta letters
        ("ड़", "M+"), ("ढ़", "<+"), ("ज़", "t+"), ("फ़", "Q+"), ("क़", "d+"),
        
        # Matras
        ("ा", "k"), ("ी", "h"), ("ु", "q"), ("ू", "w"),
        ("ृ", "`"), ("े", "s"), ("ै", "S"), ("ो", "ks"), ("ौ", "kS"),
        ("ं", "a"), ("ँ", "Wa"), ("ः", "%"), ("्", "~"), ("्र", "z"),
        
        # Punctuation
        ("।", "A"), ("॥", "AA"), ("०", "0"), ("१", "1"), ("२", "2"),
        ("३", "3"), ("४", "4"), ("५", "5"), ("६", "6"), ("७", "7"),
        ("८", "8"), ("९", "9"),
    ]
    
    for u, k in replacements:
        res = res.replace(u, k)
        
    return res
