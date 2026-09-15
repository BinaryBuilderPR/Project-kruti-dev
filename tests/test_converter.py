"""
Unit tests for Kruti Dev 010 <-> Unicode converter.
"""

import os
import sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
sys.stdout.reconfigure(encoding='utf-8')
from krutidev_engine.converter import kruti_to_unicode, unicode_to_kruti

TEST_CASES = [
    ('Hkkjr', 'भारत'),
    ('dk;kZy;', 'कार्यालय'),
    ("f'k{kk", 'शिक्षा'),
    ('foHkkx', 'विभाग'),
    ("vkns'k", 'आदेश'),
    ('Kkikad', 'ज्ञापांक'),
    ('i=kad', 'पत्रांक'),
    ('ljk;dsyk', 'सरायकेला'),
    ('[kjlkokWa', 'खरसावाँ'),
    ('mijksDr', 'उपरोक्त'),
    ("f'k{kd", 'शिक्षक'),
    ('osru', 'वेतन'),
    ('fon~;ky;', 'विद्यालय'),
    ('Lohd`r', 'स्वीकृत'),
    ('dk;Zokgh', 'कार्यवाही'),
    ('esa', 'में'),
    ('eSa', 'मैं'),
    ('dh', 'की'),
    ('fd', 'कि'),
    ('gS', 'है'),
    ('gSa', 'हैं'),
    ('vkns\'kkuqlkj', 'आदेशानुसार'),
    ('izfrosnu', 'प्रतिवेदन'),
    ('vuqikyu', 'अनुपालन'),
]

def test_kruti_to_unicode():
    for k, expected_u in TEST_CASES:
        actual_u = kruti_to_unicode(k)
        assert actual_u == expected_u, f"Failed for {k}: got {actual_u}, expected {expected_u}"

def test_unicode_to_kruti():
    for expected_k, u in TEST_CASES:
        actual_k = unicode_to_kruti(u)
        assert actual_k == expected_k, f"Failed for {u}: got {actual_k}, expected {expected_k}"

if __name__ == '__main__':
    passed = 0
    total = len(TEST_CASES) * 2
    for k, expected_u in TEST_CASES:
        actual_u = kruti_to_unicode(k)
        if actual_u == expected_u:
            passed += 1
            print(f"[PASS K->U] {k:15} -> {actual_u}")
        else:
            print(f"[FAIL K->U] {k:15} -> Got: '{actual_u}', Expected: '{expected_u}'")
            
    for expected_k, u in TEST_CASES:
        actual_k = unicode_to_kruti(u)
        if actual_k == expected_k:
            passed += 1
            print(f"[PASS U->K] {u:15} -> {actual_k}")
        else:
            print(f"[FAIL U->K] {u:15} -> Got: '{actual_k}', Expected: '{expected_k}'")
            
    print(f"\nResult: {passed}/{total} tests passed ({passed/total*100:.1f}%)")
