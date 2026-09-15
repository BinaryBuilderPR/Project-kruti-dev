"""
High-performance Prefix Trie for Hindi / Kruti Dev word autocompletion and spellchecking.
Supports dual prefix search (Kruti Dev prefix or Unicode Devanagari prefix).
"""

from typing import List, Dict, Optional, Tuple


class TrieNode:
    __slots__ = ('children', 'is_end_of_word', 'frequency', 'unicode_word', 'kruti_word', 'category')
    
    def __init__(self):
        self.children: Dict[str, TrieNode] = {}
        self.is_end_of_word: bool = False
        self.frequency: int = 0
        self.unicode_word: str = ""
        self.kruti_word: str = ""
        self.category: str = "general"


class KrutiTrie:
    """Trie index for fast prefix autocomplete and spellchecking."""
    
    def __init__(self):
        self.root = TrieNode()
        self.total_words = 0
        self._exact_kruti_set = set()
        self._exact_unicode_set = set()

    def insert(self, kruti_word: str, unicode_word: str, frequency: int = 1, category: str = "general") -> None:
        """Inserts a word pair into the Trie."""
        if not kruti_word or not unicode_word:
            return
        
        self._exact_kruti_set.add(kruti_word.lower())
        self._exact_unicode_set.add(unicode_word)
        
        node = self.root
        for char in kruti_word:
            if char not in node.children:
                node.children[char] = TrieNode()
            node = node.children[char]
            
        node.is_end_of_word = True
        # Keep highest frequency if already present
        if frequency > node.frequency:
            node.frequency = frequency
        node.unicode_word = unicode_word
        node.kruti_word = kruti_word
        node.category = category
        self.total_words += 1

    def contains_kruti(self, kruti_word: str) -> bool:
        """Returns True if the exact Kruti Dev word exists in the dictionary."""
        return kruti_word.lower() in self._exact_kruti_set or kruti_word in self._exact_kruti_set

    def contains_unicode(self, unicode_word: str) -> bool:
        """Returns True if the exact Unicode Devanagari word exists in the dictionary."""
        return unicode_word in self._exact_unicode_set

    def autocomplete(self, prefix: str, max_results: int = 7) -> List[Dict[str, any]]:
        """
        Returns top `max_results` autocompletion candidates matching `prefix`.
        Sorted by frequency (highest first).
        """
        if not prefix:
            return []
            
        node = self.root
        for char in prefix:
            if char not in node.children:
                return []
            node = node.children[char]
            
        # Collect all words in subtree
        results: List[Tuple[int, str, str, str]] = []
        self._collect(node, results)
        
        # Sort by frequency descending
        results.sort(key=lambda x: x[0], reverse=True)
        
        formatted = []
        for freq, kruti, unicode_val, cat in results[:max_results]:
            formatted.append({
                "kruti": kruti,
                "unicode": unicode_val,
                "frequency": freq,
                "category": cat
            })
        return formatted

    def _collect(self, node: TrieNode, results: List[Tuple[int, str, str, str]]) -> None:
        if node.is_end_of_word:
            results.append((node.frequency, node.kruti_word, node.unicode_word, node.category))
            
        for child in node.children.values():
            self._collect(child, results)

