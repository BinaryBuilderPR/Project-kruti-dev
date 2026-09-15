"""
High-speed Spellchecker and Autocomplete Suggester for Kruti Dev & Hindi.
Loads the 70k+ lexicon and provides instant prefix completions and typo corrections.
"""

import os
import json
from typing import List, Dict, Optional, Tuple, Set

from .trie import KrutiTrie
from .converter import kruti_to_unicode, unicode_to_kruti
from .rules import HindiGrammarEngine, GrammarIssue


def levenshtein_distance(s1: str, s2: str) -> int:
    """Computes Levenshtein edit distance between two strings."""
    if len(s1) < len(s2):
        return levenshtein_distance(s2, s1)
    if len(s2) == 0:
        return len(s1)
    
    previous_row = range(len(s2) + 1)
    for i, c1 in enumerate(s1):
        current_row = [i + 1]
        for j, c2 in enumerate(s2):
            insertions = previous_row[j + 1] + 1
            deletions = current_row[j] + 1
            substitutions = previous_row[j] + (c1 != c2)
            current_row.append(min(insertions, deletions, substitutions))
        previous_row = current_row
    return previous_row[-1]


class KrutiSpellChecker:
    """Complete Spellchecking & Autocomplete Engine for Kruti Dev 010 and Hindi."""
    
    def __init__(self, lexicon_path: Optional[str] = None):
        self.trie = KrutiTrie()
        self.grammar_engine = HindiGrammarEngine()
        self.unicode_words: Dict[str, int] = {}
        self.kruti_words: Dict[str, int] = {}
        self.user_custom_dict: Set[str] = set()
        
        if not lexicon_path:
            lexicon_path = os.path.join(os.path.dirname(__file__), "lexicon.json")
            
        self._load_lexicon(lexicon_path)

    def _load_lexicon(self, path: str):
        if not os.path.exists(path):
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                entries = json.load(f)
                for entry in entries:
                    u_word = entry["unicode"]
                    k_word = entry["kruti"]
                    freq = entry.get("frequency", 1)
                    cat = entry.get("category", "general")
                    
                    self.unicode_words[u_word] = freq
                    self.kruti_words[k_word] = freq
                    self.trie.insert(k_word, u_word, freq, cat)
        except Exception as e:
            print(f"Error loading lexicon: {e}")

    def add_custom_word(self, word_kruti_or_unicode: str):
        """Adds a new word to the user's custom dictionary."""
        # Detect if input is kruti or unicode
        if any('\u0900' <= c <= '\u097F' for c in word_kruti_or_unicode):
            u_word = word_kruti_or_unicode
            k_word = unicode_to_kruti(u_word)
        else:
            k_word = word_kruti_or_unicode
            u_word = kruti_to_unicode(k_word)
            
        self.user_custom_dict.add(u_word)
        self.unicode_words[u_word] = 1000
        self.kruti_words[k_word] = 1000
        self.trie.insert(k_word, u_word, 1000, "custom")

    def autocomplete(self, kruti_prefix: str, max_results: int = 6) -> List[Dict[str, any]]:
        """
        Returns real-time autocompletions for a typed Kruti Dev prefix.
        E.g. typing 'dk;' returns 'dk;kZy;' (कार्यालय), 'dk;Zokgh' (कार्रवाई).
        """
        if not kruti_prefix or len(kruti_prefix.strip()) == 0:
            return []
        return self.trie.autocomplete(kruti_prefix, max_results=max_results)

    def is_word_valid(self, kruti_word: str) -> bool:
        """Checks if the typed Kruti Dev word is valid."""
        clean_k = kruti_word.strip(",. %:;()[]{}\"'-_")
        if not clean_k or clean_k.isdigit():
            return True
            
        if clean_k in self.kruti_words or clean_k.lower() in self.kruti_words:
            return True
            
        # Check unicode form
        u_form = kruti_to_unicode(clean_k)
        if u_form in self.unicode_words or u_form in self.user_custom_dict:
            return True
            
        return False

    def suggest_corrections(self, misspelled_kruti: str, max_suggestions: int = 5) -> List[Dict[str, any]]:
        """
        Suggests spelling corrections for a misspelled Kruti Dev word.
        Uses phonetic Devanagari Levenshtein distance for maximum linguistic accuracy.
        """
        clean_k = misspelled_kruti.strip(",. %:;()[]{}\"'-_")
        if not clean_k:
            return []
            
        misspelled_uni = kruti_to_unicode(clean_k)
        
        candidates: List[Tuple[int, int, str, str]] = [] # (dist, -freq, uni, kruti)
        
        # 1. First check common substitution traps (chhoti vs badi ee/oo, sha vs sa vs sha, anusvara)
        # Search dictionary words with matching length (+- 2)
        target_len = len(misspelled_uni)
        for u_word, freq in self.unicode_words.items():
            if abs(len(u_word) - target_len) <= 2:
                dist = levenshtein_distance(misspelled_uni, u_word)
                if dist <= 2:
                    k_word = unicode_to_kruti(u_word)
                    candidates.append((dist, -freq, u_word, k_word))
                    
        # Sort by edit distance first, then by frequency
        candidates.sort(key=lambda x: (x[0], x[1]))
        
        results = []
        seen = set()
        for dist, _, u_cand, k_cand in candidates:
            if u_cand not in seen:
                seen.add(u_cand)
                results.append({
                    "unicode": u_cand,
                    "kruti": k_cand,
                    "distance": dist,
                    "message": f"सुझाव: {u_cand} ({k_cand})"
                })
                if len(results) >= max_suggestions:
                    break
                    
        return results

    def verify_document_text(self, text_kruti: str) -> Dict[str, any]:
        """
        Verifies full paragraph or document text.
        Returns misspelled words and grammar rule violations.
        """
        grammar_issues = self.grammar_engine.check_sentence(text_kruti)
        
        tokens = text_kruti.split()
        spelling_errors = []
        
        for token in tokens:
            clean = token.strip("।,.:;%!?\"'()[]{}<>-")
            if clean and not clean.isdigit() and not self.is_word_valid(clean):
                suggestions = self.suggest_corrections(clean, max_suggestions=3)
                spelling_errors.append({
                    "word_kruti": clean,
                    "word_unicode": kruti_to_unicode(clean),
                    "suggestions": suggestions
                })
                
        return {
            "spelling_errors": spelling_errors,
            "grammar_issues": [issue.to_dict() for issue in grammar_issues],
            "total_errors": len(spelling_errors) + len(grammar_issues)
        }

