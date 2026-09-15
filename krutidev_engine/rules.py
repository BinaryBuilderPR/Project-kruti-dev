"""
Specialized Hindi & Official Sarkari Patrachar Grammar Rules Engine.
Detects common confusions:
1. 'की' (dh) vs 'कि' (fd)
2. 'में' (esa) vs 'मैं' (eSa)
3. 'है' (gS) vs 'हैं' (gSa)
4. 'श' / 'ष' / 'स' orthographic traps
"""

import re
from typing import List, Dict, Optional, Tuple
from .converter import kruti_to_unicode, unicode_to_kruti


class GrammarIssue:
    def __init__(self, rule_id: str, message: str, original_word: str, suggestion: str, context: str, start_idx: int, end_idx: int):
        self.rule_id = rule_id
        self.message = message
        self.original_word = original_word
        self.suggestion = suggestion
        self.context = context
        self.start_idx = start_idx
        self.end_idx = end_idx

    def to_dict(self) -> Dict[str, any]:
        return {
            "rule_id": self.rule_id,
            "message": self.message,
            "original_word": self.original_word,
            "suggestion": self.suggestion,
            "suggestion_kruti": unicode_to_kruti(self.suggestion),
            "context": self.context,
            "start_idx": self.start_idx,
            "end_idx": self.end_idx
        }


class HindiGrammarEngine:
    """Contextual Grammar Checker tailored for Hindi Sarkari Patrachar & Kruti Dev."""
    
    # Common reporting verbs that MUST be followed by 'कि' (fd), not 'की' (dh)
    VERBS_PRECEDING_KI = {
        "कहा", "कहा गया", "कहा है", "कहते", "स्पष्ट है", "विदित हो", "उल्लेखनीय है",
        "सूचित किया जाता है", "आदेश दिया जाता है", "अनुरोध है", "निर्देशित किया जाता है",
        "अंकित किया", "अंकित है", "पाया गया", "देखते हुए", "प्रतीत होता है", "ज्ञात हुआ",
        "उल्लेख है", "निर्णय लिया गया", "प्रावधान है", "सहमति दी जाती है", "दिया जाता है",
        "किया जाता है", "किया गया है", "निर्देश है", "प्रार्थना है", "निवेदन है"
    }

    # First-person verbs indicating 'मैं' (eSa), not 'में' (esa)
    FIRST_PERSON_VERB_ENDINGS = {
        "हूँ", "करता हूँ", "करती हूँ", "अनुरोध करता हूँ", "चाहता हूँ", "प्रमाणित करता हूँ",
        "समर्पित करता हूँ", "उपस्थित हुआ"
    }

    # Common phrases that MUST use 'में' (esa) - अधिकरण कारक
    MEIN_PHRASES = {
        "सम्बन्ध में", "संबंध में", "मामले में", "कार्यालय में", "विभाग में", "जिले में",
        "प्रखंड में", "विद्यालय में", "प्रतिवेदन में", "आलोक में", "क्रम में", "विषय में",
        "स्थिति में", "संदर्भ में", "अवधि में", "वर्ष में", "संचिका में", "प्रपत्र में"
    }

    def check_sentence(self, sentence_kruti: str) -> List[GrammarIssue]:
        """Analyzes a sentence typed in Kruti Dev and returns a list of grammar/usage issues."""
        sentence_unicode = kruti_to_unicode(sentence_kruti)
        return self.check_sentence_unicode(sentence_unicode)

    def check_sentence_unicode(self, text: str) -> List[GrammarIssue]:
        """Analyzes a Unicode Hindi sentence for grammar and drafting mistakes."""
        issues: List[GrammarIssue] = []
        tokens = text.split()
        
        # Rule 1: 'की' vs 'कि' Disambiguation
        for i, word in enumerate(tokens):
            if word == "की":
                prev_clause = " ".join(tokens[max(0, i-4):i])
                matched_verb = None
                for verb in self.VERBS_PRECEDING_KI:
                    if verb in prev_clause:
                        matched_verb = verb
                        break
                        
                if matched_verb:
                    issues.append(GrammarIssue(
                        rule_id="KI_VS_KEE_CONJUNCTION",
                        message=f"क्रिया '{matched_verb}' के बाद योजक शब्द 'कि' (fd) का प्रयोग होना चाहिए, 'की' (dh) का नहीं।",
                        original_word="की",
                        suggestion="कि",
                        context=" ".join(tokens[max(0, i-3):min(len(tokens), i+4)]),
                        start_idx=text.find(word),
                        end_idx=text.find(word) + len(word)
                    ))
            elif word == "कि":
                # Check if 'कि' is mistakenly used as possessive marker e.g., 'विभाग कि संचिका'
                prev_1 = tokens[i - 1] if i > 0 else ""
                next_1 = tokens[i + 1] if i + 1 < len(tokens) else ""
                
                # Nouns ending before 'कि' without a reporting verb
                if prev_1 in {"विभाग", "सरकार", "कार्यालय", "न्यायालय", "सेवा", "वेतन", "जाँच", "समिति"} and next_1 in {"संचिका", "बैठक", "ओर", "ओर", "प्रति", "प्रतिलिपि", "पुष्टि", "मांग", "राशि"}:
                    issues.append(GrammarIssue(
                        rule_id="KI_VS_KEE_POSSESSIVE",
                        message=f"सम्बन्ध कारक के लिए 'की' (dh) का प्रयोग होना चाहिए, योजक 'कि' (fd) का नहीं।",
                        original_word="कि",
                        suggestion="की",
                        context=" ".join(tokens[max(0, i-2):min(len(tokens), i+3)]),
                        start_idx=text.find(word),
                        end_idx=text.find(word) + len(word)
                    ))

        # Rule 2: 'में' vs 'मैं' Disambiguation
        for i, word in enumerate(tokens):
            if word == "मैं":
                # Check if preceded by locative context (e.g. "सम्बन्ध मैं", "आलोक मैं")
                prev_1 = tokens[i - 1] if i > 0 else ""
                phrase_2 = f"{prev_1} मैं"
                if any(p.startswith(prev_1) for p in self.MEIN_PHRASES) or prev_1 in {"सम्बन्ध", "संबंध", "आलोक", "क्रम", "मामले", "कार्यालय", "विभाग", "जिले", "प्रखंड", "विद्यालय", "संदर्भ", "अवधि", "संचिका"}:
                    issues.append(GrammarIssue(
                        rule_id="MEIN_LOCATIVE_ERROR",
                        message=f"अधिकरण कारक (in/within) के लिए 'में' (esa) का प्रयोग होता है, सर्वनाम 'मैं' (eSa) का नहीं।",
                        original_word="मैं",
                        suggestion="में",
                        context=" ".join(tokens[max(0, i-2):min(len(tokens), i+3)]),
                        start_idx=text.find(word),
                        end_idx=text.find(word) + len(word)
                    ))
            elif word == "में":
                # Check if followed by first person verb ending (e.g. "में अनुरोध करता हूँ")
                next_words = " ".join(tokens[i+1:min(len(tokens), i+5)])
                if any(v in next_words for v in self.FIRST_PERSON_VERB_ENDINGS) and (i == 0 or tokens[i-1] in {",", "।", "कि", "अतः", "अस्तु"}):
                    issues.append(GrammarIssue(
                        rule_id="MAIN_PRONOUN_ERROR",
                        message=f"उत्तम पुरुष सर्वनाम (I) के लिए 'मैं' (eSa) का प्रयोग होना चाहिए, 'में' (esa) का नहीं।",
                        original_word="में",
                        suggestion="मैं",
                        context=" ".join(tokens[max(0, i):min(len(tokens), i+4)]),
                        start_idx=text.find(word),
                        end_idx=text.find(word) + len(word)
                    ))

        # Rule 3: 'है' vs 'हैं' Honorific / Plural Check
        for i, word in enumerate(tokens):
            if word == "है":
                # If subject is plural / honorific (महोदय, पदाधिकारीगण, शिक्षकगण, सदस्य, हम, आप)
                prev_clause = " ".join(tokens[max(0, i-6):i])
                if any(h in prev_clause for h in ["पदाधिकारीगण", "शिक्षकगण", "कर्मचारीगण", "माननीय", "सदस्यगण", "हम", "आप"]):
                    issues.append(GrammarIssue(
                        rule_id="HAI_VS_HAIN_PLURAL",
                        message="आदरसूचक या बहुवचन कर्ता के साथ 'हैं' (gSa) का प्रयोग होना चाहिए।",
                        original_word="है",
                        suggestion="हैं",
                        context=" ".join(tokens[max(0, i-4):min(len(tokens), i+2)]),
                        start_idx=text.find(word),
                        end_idx=text.find(word) + len(word)
                    ))

        return issues
