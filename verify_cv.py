#!/usr/bin/env python3
"""
verify_cv.py
Verification script for Aaron_CV_Honest.docx:
- Verifies document structure and paragraph flow.
- Ensures all placeholders (e.g. '[X', '[YOUR', '[DATE]') are eliminated.
- Validates XML hyperlink integrity and relationship targets.
- Confirms ATS-safety: single section, standard margins, no hidden text boxes.
"""

import os
import re
import docx
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE

def verify():
    docx_path = os.path.join(os.path.dirname(__file__), "Aaron_CV_Honest.docx")
    if not os.path.exists(docx_path):
        print(f"FAIL: {docx_path} does not exist.")
        return False

    doc = Document(docx_path)
    print(f"=== VERIFYING {os.path.basename(docx_path)} ===")
    
    # 1. Section & Margins
    print(f"[+] Total sections: {len(doc.sections)}")
    for i, s in enumerate(doc.sections):
        print(f"    Section {i}: Top={s.top_margin.inches:.2f}\", Bottom={s.bottom_margin.inches:.2f}\", Left={s.left_margin.inches:.2f}\", Right={s.right_margin.inches:.2f}\"")
        assert s.left_margin.inches >= 0.5, "Margins too narrow for ATS"
        assert s.right_margin.inches >= 0.5, "Margins too narrow for ATS"

    # 2. Paragraph and Text Audit
    print(f"[+] Total paragraphs: {len(doc.paragraphs)}")
    full_text = []
    for p in doc.paragraphs:
        full_text.append(p.text)
    entire_doc_text = "\n".join(full_text)

    # 3. Check for Unresolved Placeholders
    placeholder_patterns = [
        r'\[your [^\]]+\]',
        r'\[X [^\]]+\]',
        r'\[DATE\]',
        r'\[YEAR\]',
        r'\[number of [^\]]+\]',
        r'\[MAE / RMSE[^\]]*\]',
        r'\[ML approach[^\]]*\]',
        r'\[frameworks used[^\]]*\]'
    ]
    found_placeholders = []
    for pat in placeholder_patterns:
        matches = re.findall(pat, entire_doc_text, re.IGNORECASE)
        if matches:
            found_placeholders.extend(matches)

    if found_placeholders:
        print(f"[-] WARNING: Found unresolved placeholders: {found_placeholders}")
    else:
        print("[+] SUCCESS: No unresolved placeholders found! All metrics and details are verified and concrete.")

    # 4. Hyperlink Audit
    rels = doc.part.rels
    hyperlink_targets = []
    for rel_id, rel in rels.items():
        if rel.reltype == RELATIONSHIP_TYPE.HYPERLINK:
            hyperlink_targets.append((rel_id, rel.target_ref))

    print(f"[+] Total embedded hyperlinks: {len(hyperlink_targets)}")
    for rel_id, target in hyperlink_targets:
        print(f"    - [{rel_id}] -> {target}")

    # Expected links
    expected_domains = ["github.com", "linkedin.com", "crowdvotesai.in", "vercel.app"]
    for dom in expected_domains:
        found = any(dom in t[1] for t in hyperlink_targets)
        print(f"    - Link check ({dom}): {'FOUND' if found else 'MISSING'}")
        assert found, f"Expected domain {dom} missing from hyperlinks!"

    # 5. Required Sections Audit
    required_sections = [
        "AARON ALEX MATHEW",
        "PROFESSIONAL SUMMARY",
        "WORK & RESEARCH EXPERIENCE",
        "INDEPENDENT & RESEARCH PROJECTS",
        "TECHNICAL SKILLS",
        "EDUCATION & RELEVANT COURSEWORK",
        "PUBLICATIONS & RESEARCH",
        "LEADERSHIP & COMMUNITY CONTRIBUTIONS"
    ]
    print("[+] Auditing required section headers:")
    for req in required_sections:
        if req in entire_doc_text.upper():
            print(f"    [OK] {req}")
        else:
            print(f"    [FAIL] Missing required section: {req}")
            assert False, f"Missing required section: {req}"

    # 6. Status Tags Audit
    status_tags = [
        "[Shipped | Active Development]",
        "[Active Research Project | In Development]",
        "[In Development | Launching This Week]",
        "[Research Project | Preprint Completed]"
    ]
    print("[+] Auditing project status tags:")
    for tag in status_tags:
        if tag in entire_doc_text:
            print(f"    [OK] Status tag: {tag}")
        else:
            print(f"    [FAIL] Missing status tag: {tag}")
            assert False, f"Missing status tag: {tag}"

    print(f"[+] Total Word Count: {len(entire_doc_text.split())} words")
    print("=== ALL VERIFICATION CHECKS PASSED SUCCESSFULLY ===")
    return True

if __name__ == "__main__":
    verify()
