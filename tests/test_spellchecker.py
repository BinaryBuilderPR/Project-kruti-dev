"""
Unit and integration tests for Spellchecker and Grammar Rules engine.
"""

import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
sys.stdout.reconfigure(encoding='utf-8')

from krutidev_engine.spellchecker import KrutiSpellChecker
from krutidev_engine.rules import HindiGrammarEngine

def test_autocomplete_and_spellcheck():
    checker = KrutiSpellChecker()
    print(f"Loaded dictionary with {len(checker.unicode_words)} words.")
    
    # 1. Test Autocomplete for "dk;" (कार्या...)
    print("\n--- 1. Testing Autocomplete for 'dk;' (कार्या...) ---")
    suggestions = checker.autocomplete("dk;", max_results=5)
    for i, s in enumerate(suggestions, 1):
        print(f"  {i}. {s['unicode']} (Kruti: {s['kruti']}) [Freq: {s['frequency']}]")
    assert len(suggestions) > 0, "Autocomplete returned empty!"
    
    # 2. Test Autocomplete for "f'k" (शि...)
    print("\n--- 2. Testing Autocomplete for 'f\\'k' (शि...) ---")
    suggestions_sh = checker.autocomplete("f'k", max_results=5)
    for i, s in enumerate(suggestions_sh, 1):
        print(f"  {i}. {s['unicode']} (Kruti: {s['kruti']}) [Freq: {s['frequency']}]")
    assert len(suggestions_sh) > 0, "Autocomplete for f'k returned empty!"
    
    # 3. Test Spellcheck Fuzzy Suggester for a typo
    # Suppose user mistakenly writes "dk;Zy;" (missing aa matra in karyalay)
    print("\n--- 3. Testing Fuzzy Correction for typo 'dk;Zy;' (कार्यलय -> कार्यालय) ---")
    corrections = checker.suggest_corrections("dk;Zy;", max_suggestions=3)
    for c in corrections:
        print(f"  -> {c['unicode']} (Kruti: {c['kruti']}, Dist: {c['distance']})")
    assert any(c['unicode'] == 'कार्यालय' for c in corrections), "Failed to suggest कार्यालय for typo!"
    
    # 4. Test Grammar Disambiguation (की vs कि and में vs मैं)
    print("\n--- 4. Testing Grammar Disambiguation ---")
    grammar = HindiGrammarEngine()
    
    # Mistake 1: "आदेश दिया जाता है की आप उपस्थित हों" (used 'की' instead of 'कि')
    sent1 = "vkns'k fn;k tkrk gS dh vki mifLFkfr gksa"
    issues1 = checker.verify_document_text(sent1)
    print(f"Sentence 1 Grammar Issues: {len(issues1['grammar_issues'])}")
    for g in issues1['grammar_issues']:
        print(f"  Rule: {g['rule_id']} | Msg: {g['message']} | Suggestion: {g['suggestion']}")
    assert any(g['rule_id'] == 'KI_VS_KEE_CONJUNCTION' for g in issues1['grammar_issues'])
    
    # Mistake 2: "इस सम्बन्ध मैं सूचित किया जाता है" (used 'मैं' instead of 'में')
    sent2 = "bl lEcU/k eSa lwpgr fd;k tkrk gS"
    issues2 = checker.verify_document_text(sent2)
    print(f"Sentence 2 Grammar Issues: {len(issues2['grammar_issues'])}")
    for g in issues2['grammar_issues']:
        print(f"  Rule: {g['rule_id']} | Msg: {g['message']} | Suggestion: {g['suggestion']}")
    assert any(g['rule_id'] == 'MEIN_LOCATIVE_ERROR' for g in issues2['grammar_issues'])

    print("\n All Spellchecker and Grammar Engine tests PASSED successfully!")


if __name__ == "__main__":
    test_autocomplete_and_spellcheck()

