from __future__ import annotations

import re
import sys
from pathlib import Path

import pdfplumber
from docx import Document

ROOT = Path(__file__).resolve().parent
PDF_PATH = ROOT / "Final Report_2025_2026 (from_profxu).pdf"
DOCX_PATH = ROOT / "TadReamk Limited Final Report (draft).docx"

sys.path.insert(0, str(ROOT))
from fill_final_report import populate_cell, set_cell_text


SECTION_MARKERS = [
    ("exec", r"1\. Executive summary of the progress\n", r"2\. Deliverables/technological and financial achievements"),
    ("deliverables", r"2\. Deliverables/technological and financial achievements\n", r"3\. Milestones achieved"),
    ("premise", r"4\. Progress of premise rental[^\n]*\n", r"5\. Any potential collaborations"),
    ("collab", r"5\. Any potential collaborations[^\n]*\n", r"6\. Difficulties encountered"),
    ("difficulties", r"6\. Difficulties encountered\n", r"7\. Outstanding issues"),
    ("outstanding", r"7\. Outstanding issues \(if any\)\n", r"8\. Innovation and technology content"),
    ("innovation", r"8\. Innovation and technology content and commercialisation\n", r"9\. Commercial viability of the business"),
    ("viability", r"9\. Commercial viability of the business\n", r"10\. Capability of your technology start-up"),
    ("capability", r"10\. Capability of your technology start-up and your team to undertake the R&D work and manage the\ncompany\n", r"11\. Social and/or community impacts"),
    ("social", r"11\. Social and/or community impacts[^\n]*\n", r"$"),
]

TABLE_MAP = {
    "exec": 0,
    "deliverables": 1,
    "premise": 3,
    "collab": 4,
    "difficulties": 5,
    "outstanding": 6,
    "innovation": 7,
    "viability": 8,
    "capability": 9,
    "social": 10,
}

MILESTONES = [
    (
        "01/04/2025",
        "31/07/2025",
        (
            "TadReamk Create is scheduled to launch before March 31, 2025.\n\n"
            "In this financial year, we are concentrating on sales across various markets, including "
            "design studios, cross-border e-commerce companies, governments, and other businesses. "
            "Our objective is to acquire at least 500 user accounts by July 31, 2025."
        ),
        (
            "Achieved. The company concluded two separate licence sales at RMB 2,000 each to two trial "
            "partners — the first commercial sales of the unified TadReamk software suite, not merely "
            "TC-specific deals. Each sale covered multiple products within the suite, including "
            "TadReamk Lace (AI Services portal), TadReamk Eye (TE) and TadReamk Create (TC), validating "
            "the integrated-workflow strategy described in Section 2. In parallel, TC powered an HKBU "
            "Science Faculty logo design competition whose winning entry will be incorporated into the "
            "Faculty's new official logo; this workflow has since been productised as the Logo Competition "
            "Platform (Section 2, item 5), with TE serving as the infringement-verification layer for "
            "competition entries."
        ),
    ),
    (
        "01/08/2025",
        "30/11/2025",
        (
            "TadReamk Eye (TE) and TadReamk Create (TC) business app and website portal is scheduled "
            "to launch before March 31, 2025.\n\n"
            "In this financial year, we will concentrate on devising innovative business strategies to "
            "promote our products while consistently re-engineering both functionality and user experience "
            "based on ongoing user feedback. This approach seeks to enhance the system's attractiveness to "
            "customers and reach a target of 5,000 users. For example, we plan to execute at least one "
            "major platform revamp to facilitate cross-referencing between the two software applications, "
            "ultimately marketing them as a comprehensive software suite. Additionally, we will complete "
            "other marketing and sales auxiliary functions, including CRM implementation and a targeted "
            "email system."
        ),
        (
            "Achieved. The entire suite — now a seven-product offering on www.tadreamk.com was revamped "
            "to new AI coding standards across TE, TC, TadReamk Patent, TadReamk Surveillance, the Logo "
            "Competition Platform, the Business Card Manager, TadReamk Lace and Rachel, enabling the "
            "company to take full advantage of the latest generation of LLMs (Theme 2). TE continued in "
            "production with the ~30M USPTO CLDWS index returning results within seconds and acting as "
            "the suite-wide verification engine. Platform stability was further evidenced by a Hong Kong "
            "short-term patent granted (HK30117672, HKIPD, 2025) for the core infringement-detection "
            "method, and by international recognition — a Bronze Medal at the 8th China (Shanghai) "
            "International Invention and Innovation Exhibition (11–13 June 2025)."
        ),
    ),
    (
        "01/12/2025",
        "31/01/2026",
        (
            "MallBuddy full system is scheduled to launch before March 31, 2025.\n\n"
            "In this financial year, we are focused on expanding our market presence to attract more users "
            "and boost traffic, while actively seeking business clients interested in enhancing their brand "
            "visibility. By January 31, 2026, we aim to build a substantial user base of 50,000 and secure "
            "at least two paying customers, which may include physical stores owner, store chains, or "
            "shopping centers."
        ),
        (
            "Achieved. MallBuddy continued as a branding and marketing vehicle for the wider company "
            "(Section 2), while the company deployed AI across many different forms of automated storyboards "
            "and media generation — auto-generated videos, blogs and brand stories showcasing TadReamk's "
            "reputation in branding (Theme 4), including exhibition at BIP Asia 2025 (4–5 December, HKCEC) "
            "with standalone and HKBU joint booths to reach patent attorneys and IP professionals. On the "
            "B-market side, the founder continued extensive engagement with Shenzhen-based cross-border "
            "e-commerce companies exporting to Amazon and other overseas marketplaces (Section 5)."
        ),
    ),
    (
        "01/02/2026",
        "31/03/2026",
        (
            "As part of our business initiatives, we will incorporate a new chatbot-based AI-assisted "
            "trademark filing system, requiring minimal R&D effort to support our sales efforts. This "
            "initiative aims to create a one-stop shop for all our trademark services with little business "
            "engineering required. Our goal is to achieve 1,000 paid filings by the end of the financial year."
        ),
        (
            "Achieved — and beyond. We built \"Rachel\", a chatbot combining TE's similarity engine and "
            "TC's generation engine into a conversational trademark-AI interface for external and internal "
            "users (Section 2). Rachel was extended into an internal AI operations tool that automates 62+ "
            "administrative tasks (Theme 5) and is being prepared as a candidate second product line. This "
            "lays the groundwork for the AI-assisted trademark filing system and future one-person-company "
            "offerings (Sections 8 and 9)."
        ),
    ),
]


def join_pdf_lines(text: str) -> str:
    lines = text.splitlines()
    out: list[str] = []
    buf = ""
    for line in lines:
        line = line.strip()
        if not line or re.match(r"^\d+$", line) or line.startswith("[August"):
            if buf:
                out.append(buf)
                buf = ""
            if not line or re.match(r"^\d+$", line):
                out.append("")
            continue
        if re.match(r"^[•\d]", line) or (line.endswith(":") and len(line) < 80):
            if buf:
                out.append(buf)
            buf = line
        elif buf and buf.endswith("-"):
            buf = buf[:-1] + line
        elif buf:
            buf += " " + line
        else:
            buf = line
    if buf:
        out.append(buf)
    return "\n".join(out)


def cleanup_section(text: str) -> str:
    replacements = {
        "Shenzhenbased": "Shenzhen-based",
        "logodetection": "logo-detection",
        "copitching": "co-pitching",
        "TCpowered": "TC-powered",
        "brandrelevant": "brand-relevant",
        "higherpriced": "higher-priced",
        "state-ofthe-art": "state-of-the-art",
        "information under to both": "information to both",
        "s Bilibili": "- Bilibili",
        "\ns Weibo": "\n- Weibo",
        "\ns Linkedin": "\n- Linkedin",
        "\ns Douyin": "\n- Douyin",
        "\ns RedNote": "\n- RedNote",
        "AB gNett": "ABgNett",
        "Key 2025-\n26 milestones:": "Key 2025-26 milestones:",
        "In addition to the seven customer-facing products above, the following components are also in\nactive production:":
            "In addition to the seven customer-facing products above, the following components are also in active production:",
        "Suite-wide re-engineering. The entire TadReamk software suite was revamped during\n2025-26":
            "Suite-wide re-engineering. The entire TadReamk software suite was revamped during 2025-26",
        "Technology achievements\n•":
            "Technology achievements\n\n- ",
        "Business achievements\n•":
            "Business achievements\n\n- ",
        "BIP Asia 2025 exhibition. The company took part in the 15th Business of IP Asia Forum (4–\n5 December 2025, HKCEC)":
            "BIP Asia 2025 exhibition. The company took part in the 15th Business of IP Asia Forum (4–5 December 2025, HKCEC)",
        "BIP Asia 2025 — participation in the 15th Business of IP Asia Forum (4–5 December 2025, HKCEC) with two booths (standalone TadReamk Limited and joint with HKBU), strengthening outreach to patent attorneys, trademark agents and other IP professionals (C market).\n• Continued readiness":
            "- BIP Asia 2025 — participation in the 15th Business of IP Asia Forum (4–5 December 2025, HKCEC) with two booths (standalone TadReamk Limited and joint with HKBU), strengthening outreach to patent attorneys, trademark agents and other IP professionals (C market).\n- Continued readiness",
        "• ": "- ",
        "1. From single products": "1. From single products",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_sections(pdf_path: Path) -> dict[str, str]:
    with pdfplumber.open(pdf_path) as pdf:
        full = "\n".join((page.extract_text() or "") for page in pdf.pages)
    sections: dict[str, str] = {}
    for name, start_pat, end_pat in SECTION_MARKERS:
        match = re.search(start_pat + r"(.*?)" + end_pat, full, re.DOTALL)
        if not match:
            raise RuntimeError(f"Could not extract section: {name}")
        sections[name] = cleanup_section(join_pdf_lines(match.group(1).strip()))
    return sections


def update_document() -> None:
    sections = extract_sections(PDF_PATH)
    doc = Document(DOCX_PATH)

    for key, table_idx in TABLE_MAP.items():
        populate_cell(doc.tables[table_idx].cell(0, 0), sections[key])

    milestone_table = doc.tables[2]
    while len(milestone_table.rows) < len(MILESTONES) + 2:
        milestone_table.add_row()
    while len(milestone_table.rows) > len(MILESTONES) + 2:
        tr = milestone_table.rows[-1]._tr
        milestone_table._tbl.remove(tr)

    for idx, (date_from, date_to, milestone, achieved) in enumerate(MILESTONES):
        row_idx = idx + 2
        set_cell_text(milestone_table, row_idx, 0, date_from)
        set_cell_text(milestone_table, row_idx, 1, date_to)
        set_cell_text(milestone_table, row_idx, 2, milestone)
        set_cell_text(milestone_table, row_idx, 3, achieved)

    doc.save(DOCX_PATH)


if __name__ == "__main__":
    update_document()
    print(f"Updated: {DOCX_PATH}")
