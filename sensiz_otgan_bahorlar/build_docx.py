#!/usr/bin/env python3
"""Build the 150-page Word edition of “Sensiz o‘tgan bahorlar”.

The script uses only Python's standard library. It creates a valid OOXML .docx
archive and divides every chapter into exactly six manually paginated pages.
"""

from __future__ import annotations

import re
import zipfile
from datetime import datetime, timezone
from html import escape
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CHAPTER_DIR = ROOT / "boblar"
OUTPUT = ROOT.parent / "Sensiz_otgan_bahorlar_150_bet.docx"
PAGES_PER_CHAPTER = 6
FRONT_PAGES = 6
EXPECTED_CHAPTERS = 24
EXPECTED_PAGES = FRONT_PAGES + EXPECTED_CHAPTERS * PAGES_PER_CHAPTER

CHAPTER_FILES = [
    "01_siyohga_tushgan_bodom_guli.md",
    "02_kok_ipga_boglangan_kalit.md",
    "03_yomgir_yopgan_bozor.md",
    "04_atlas_orasidagi_maktub.md",
    "05_qarzdor_hovlining_jimligi.md",
    "06_singan_qalam_uchidagi_ism.md",
    "07_darvozadan_chiqmagan_xayr.md",
    "08_karvon_ortidan_qolgan_romol.md",
    "09_quruq_muhrning_izi.md",
    "10_oq_romol_ostidagi_vada.md",
    "11_yarim_oqilgan_maktub.md",
    "12_bir_kun_kechikkan_otliq.md",
    "13_begona_uyda_birinchi_tong.md",
    "14_qaytarib_olingan_eshik.md",
    "15_beshinchi_bahorning_qori.md",
    "16_devor_ortidagi_alifbo.md",
    "17_ombordagi_kuygan_daftar.md",
    "18_otasiga_qarshi_turgan_ogil.md",
    "19_qaytarilgan_kok_xat.md",
    "20_sakkiz_yil_jim_turgan_eshik.md",
    "21_hovuz_boyidagi_ikki_soya.md",
    "22_bosh_qolgan_nikoh_sandigi.md",
    "23_qizlar_uchun_tikilgan_daftarlar.md",
    "24_siyohdan_omon_qolgan_gul.md",
]

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_SPACE = "xml:space=\"preserve\""


def word_count(text: str) -> int:
    return len(re.findall(r"\b[\w‘’ʻʼ'-]+\b", text, flags=re.UNICODE))


def read_chapter(path: Path) -> tuple[str, list[str]]:
    text = path.read_text(encoding="utf-8").strip()
    blocks = [b.strip() for b in re.split(r"\n\s*\n", text) if b.strip()]
    if not blocks or not blocks[0].startswith("# "):
        raise ValueError(f"Chapter heading is missing: {path}")
    title = blocks.pop(0)[2:].strip()
    paragraphs: list[str] = []
    for block in blocks:
        if block == "* * *":
            paragraphs.append(block)
        else:
            paragraphs.extend(line.strip() for line in block.splitlines() if line.strip())
    return title, paragraphs


def layout_units(text: str) -> int:
    """Approximate rendered lines, including short-paragraph line loss."""
    if text == "* * *":
        return 1
    plain = strip_markdown(text)
    return max(1, (len(plain) + 81) // 82)


def partition_paragraphs(paragraphs: list[str], parts: int = 6) -> list[list[str]]:
    """Split at paragraph boundaries; balance visual lines, not only words."""
    weights = [0.14] + [(1.0 - 0.14) / (parts - 1)] * (parts - 1)
    counts = [layout_units(p) for p in paragraphs]
    total = sum(counts)
    targets = []
    running = 0.0
    for weight in weights[:-1]:
        running += weight
        targets.append(total * running)

    chunks: list[list[str]] = [[]]
    cumulative = 0
    target_index = 0
    for idx, (paragraph, count) in enumerate(zip(paragraphs, counts)):
        remaining_paragraphs = len(paragraphs) - idx
        remaining_chunks = parts - len(chunks)
        should_break = (
            target_index < len(targets)
            and cumulative >= targets[target_index]
            and remaining_paragraphs > remaining_chunks
            and chunks[-1]
            and chunks[-1][-1] != "* * *"
            and paragraph != "* * *"
        )
        if should_break:
            chunks.append([])
            target_index += 1
        chunks[-1].append(paragraph)
        cumulative += count

    while len(chunks) < parts:
        largest = max(range(len(chunks)), key=lambda i: len(chunks[i]))
        chunk = chunks[largest]
        split_at = max(1, len(chunk) // 2)
        while split_at < len(chunk) - 1 and chunk[split_at] == "* * *":
            split_at += 1
        chunks[largest : largest + 1] = [chunk[:split_at], chunk[split_at:]]
    if len(chunks) != parts or any(not chunk for chunk in chunks):
        raise ValueError("Could not create six non-empty chapter pages")
    return chunks


def strip_markdown(text: str) -> str:
    return text.replace("**", "").replace("*", "")


def run_xml(text: str, *, bold: bool = False, italic: bool = False, size: int | None = None) -> str:
    props = []
    if bold:
        props.append("<w:b/>")
    if italic:
        props.append("<w:i/>")
    if size:
        props.append(f"<w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/>")
    rpr = f"<w:rPr>{''.join(props)}</w:rPr>" if props else ""
    return f"<w:r>{rpr}<w:t {XML_SPACE}>{escape(text)}</w:t></w:r>"


def inline_runs(text: str) -> str:
    # Basic Markdown bold support used by the letter headings in the manuscript.
    pieces = re.split(r"(\*\*.*?\*\*)", text)
    runs = []
    for piece in pieces:
        if not piece:
            continue
        if piece.startswith("**") and piece.endswith("**"):
            runs.append(run_xml(piece[2:-2], bold=True))
        else:
            runs.append(run_xml(piece))
    return "".join(runs)


def paragraph_xml(
    text: str = "",
    *,
    style: str = "BodyText",
    align: str | None = None,
    indent: bool = True,
    keep_next: bool = False,
    break_after: bool = False,
) -> str:
    ppr = [f"<w:pStyle w:val=\"{style}\"/>"]
    if align:
        ppr.append(f"<w:jc w:val=\"{align}\"/>")
    if indent:
        ppr.append("<w:ind w:firstLine=\"397\"/>")
    else:
        ppr.append("<w:ind w:firstLine=\"0\"/>")
    if keep_next:
        ppr.append("<w:keepNext/>")
    body = inline_runs(text) if text else "<w:r><w:t></w:t></w:r>"
    if break_after:
        body += "<w:r><w:br w:type=\"page\"/></w:r>"
    return f"<w:p><w:pPr>{''.join(ppr)}</w:pPr>{body}</w:p>"


def display_paragraph(text: str, size: int = 28, *, bold: bool = False, italic: bool = False, after: int = 180) -> str:
    return (
        "<w:p><w:pPr><w:jc w:val=\"center\"/><w:spacing w:after=\""
        + str(after)
        + "\"/><w:ind w:firstLine=\"0\"/></w:pPr>"
        + run_xml(text, bold=bold, italic=italic, size=size)
        + "</w:p>"
    )


def blank_lines(count: int) -> str:
    return "".join(paragraph_xml("", indent=False) for _ in range(count))


def page_break() -> str:
    return paragraph_xml("", indent=False, break_after=True)


def make_front_pages(chapter_titles: list[str]) -> list[str]:
    pages: list[str] = []

    pages.append(
        blank_lines(7)
        + display_paragraph("SENSIZ O‘TGAN BAHORLAR", size=40, bold=True, after=320)
        + display_paragraph("Tarixiy-maishiy sevgi romani", size=26, italic=True, after=520)
        + blank_lines(8)
        + display_paragraph("Behruz", size=28, bold=True)
        + display_paragraph("Buxoro voqealari, 1872–1881", size=20)
    )

    pages.append(
        blank_lines(4)
        + display_paragraph("ROMAN HAQIDA", size=28, bold=True, after=360)
        + paragraph_xml(
            "1870-yillar Buxorosida yosh xattot Samandar otasi nomiga yopishgan isnodni yuvish uchun "
            "Karmana yo‘liga chiqadi. U qaytguncha atlasdo‘z Mohinurning oilasi qarz, kasallik va el gapi "
            "orasida qoladi. Bir ko‘k ipli maktub manziliga yetmaydi. Bir qalbaki muhr esa uch odamning "
            "hayotini yillar davomida boshqa yo‘lga buradi."
        )
        + paragraph_xml(
            "Roman muhabbatni faqat visol bilan emas, tanlov, g‘urur, sukut va kechikkan rostlik bilan "
            "o‘lchaydi. Samandar, Mohinur va Tohirning har biri birovni asrayman deb uning o‘rniga qaror "
            "qiladi. Bahorlar qaytadi, lekin odamlar o‘sha holicha qolmaydi."
        )
        + blank_lines(5)
        + display_paragraph("“Aytilmagan gap ham ba’zan uzoq yo‘l yuradi.”", size=24, italic=True)
    )

    pages.append(
        blank_lines(3)
        + display_paragraph("TARIXIY QAYD", size=28, bold=True, after=300)
        + paragraph_xml(
            "Roman badiiy to‘qimadir. Samandar, Mohinur, Tohir va ularning oilalari tarixiy shaxslar emas. "
            "Voqealar Amir Muzaffar hukmronligi davrida, Buxoro amirligi 1868-yildan keyin Rossiya "
            "protektoratiga aylangan tarixiy fonda kechadi. Siyosiy tarix qahramonlar hayotiga faqat savdo, "
            "yo‘l, hujjat va shahar muhitidagi o‘zgarishlar orqali kiradi."
        )
        + paragraph_xml(
            "Buxoroning savdo gumbazlari, mahalla-guzarlari, karvonsaroylari va qadimiy shahar tuzilishi "
            "UNESCO World Heritage Centre hamda Library of Congress materiallaridagi umumiy ma’lumotlar "
            "asosida tasvirlandi. Aniq diniy iqtiboslar va tekshirilmagan huquqiy hukmlar ataylab ishlatilmadi."
        )
        + paragraph_xml("Asosiy manbalar:", indent=False)
        + paragraph_xml("• UNESCO World Heritage Centre: whc.unesco.org/en/list/602", indent=False)
        + paragraph_xml("• Library of Congress: hdl.loc.gov/loc.amed/amedpllc.00406943264", indent=False)
        + paragraph_xml("• Library of Congress, The Empire That Was Russia: loc.gov/exhibits/empire/ethnic.html", indent=False)
    )

    def toc_page(start: int, end: int, label: str) -> str:
        content = blank_lines(2) + display_paragraph(label, size=28, bold=True, after=260)
        for idx in range(start, end):
            page = FRONT_PAGES + idx * PAGES_PER_CHAPTER + 1
            title = strip_markdown(chapter_titles[idx])
            content += paragraph_xml(f"{title}  ···  {page}", style="TOC1", indent=False)
        return content

    pages.append(toc_page(0, 12, "MUNDARIJA · I"))
    pages.append(toc_page(12, 24, "MUNDARIJA · II"))

    pages.append(
        blank_lines(5)
        + display_paragraph("TO‘RT FASL ORASIDA", size=30, bold=True, after=420)
        + paragraph_xml("I qism · Siyoh va bodom guli · 1872-yil", style="Subtitle", align="center", indent=False)
        + paragraph_xml("II qism · Yo‘l va yetib bormagan xat · 1872–1874-yillar", style="Subtitle", align="center", indent=False)
        + paragraph_xml("III qism · Begona uy va qaytarilgan eshik · 1874–1878-yillar", style="Subtitle", align="center", indent=False)
        + paragraph_xml("IV qism · Rostlikka qaytgan bahor · 1878–1881-yillar", style="Subtitle", align="center", indent=False)
        + blank_lines(5)
        + display_paragraph("Buxoro", size=22, italic=True)
    )
    return pages


def build_document_pages() -> tuple[list[str], list[str], int]:
    chapter_data = [read_chapter(CHAPTER_DIR / name) for name in CHAPTER_FILES]
    titles = [title for title, _ in chapter_data]
    pages = make_front_pages(titles)
    total_words = 0

    for title, paragraphs in chapter_data:
        total_words += sum(word_count(p) for p in paragraphs)
        chunks = partition_paragraphs(paragraphs, PAGES_PER_CHAPTER)
        for page_index, chunk in enumerate(chunks):
            content = ""
            if page_index == 0:
                content += paragraph_xml(title, style="Heading1", align="center", indent=False, keep_next=True)
            for paragraph in chunk:
                if paragraph == "* * *":
                    content += paragraph_xml("*  *  *", style="SceneBreak", align="center", indent=False)
                elif paragraph.startswith("—") or paragraph.startswith("–"):
                    content += paragraph_xml(paragraph, indent=False)
                else:
                    content += paragraph_xml(paragraph)
            pages.append(content)

    if len(pages) != EXPECTED_PAGES:
        raise AssertionError(f"Expected {EXPECTED_PAGES} pages, built {len(pages)}")
    return pages, titles, total_words


def styles_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W_NS}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="uz-Latn-UZ"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/><w:jc w:val="both"/><w:widowControl/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
  <w:style w:type="paragraph" w:styleId="BodyText"><w:name w:val="Body Text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/><w:jc w:val="both"/><w:widowControl/></w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="BodyText"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="0" w:after="220"/><w:jc w:val="center"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:sz w:val="30"/><w:szCs w:val="30"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="SceneBreak"><w:name w:val="Scene Break"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="100" w:after="100"/><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="100"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9072"/></w:tabs></w:pPr><w:rPr><w:sz w:val="23"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="180"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:sz w:val="24"/></w:rPr></w:style>
</w:styles>'''


def document_xml(pages: list[str]) -> str:
    body_parts = []
    for idx, content in enumerate(pages):
        body_parts.append(content)
        if idx < len(pages) - 1:
            body_parts.append(page_break())
    section = '''<w:sectPr>
      <w:footerReference w:type="default" r:id="rId2"/>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1276" w:header="567" w:footer="567" w:gutter="0"/>
      <w:pgNumType w:start="1"/>
      <w:cols w:space="708"/>
      <w:docGrid w:linePitch="276"/>
    </w:sectPr>'''
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W_NS}" xmlns:r="{R_NS}"><w:body>{''.join(body_parts)}{section}</w:body></w:document>'''


def footer_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="{W_NS}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText {XML_SPACE}> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>'''


def write_docx() -> tuple[int, int]:
    pages, _, total_words = build_document_pages()
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")

    parts = {
        "[Content_Types].xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>''',
        "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>''',
        "word/_rels/document.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>''',
        "word/document.xml": document_xml(pages),
        "word/styles.xml": styles_xml(),
        "word/footer1.xml": footer_xml(),
        "word/settings.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="{W_NS}"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:evenAndOddHeaders/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>''',
        "docProps/core.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Sensiz o‘tgan bahorlar</dc:title><dc:subject>Tarixiy-maishiy sevgi romani</dc:subject><dc:creator>Behruz</dc:creator><cp:lastModifiedBy>Kiro</cp:lastModifiedBy><dc:description>XIX asr Buxorosida kechadigan 24 bobli sevgi romani</dc:description><dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified></cp:coreProperties>''',
        "docProps/app.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Kiro OOXML Builder</Application><Pages>{EXPECTED_PAGES}</Pages><Words>{total_words}</Words><Paragraphs>0</Paragraphs><Company></Company><AppVersion>1.0</AppVersion></Properties>''',
    }

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OUTPUT, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content.encode("utf-8"))
    return len(pages), total_words


if __name__ == "__main__":
    page_count, total_words = write_docx()
    print(f"Created: {OUTPUT}")
    print(f"Manual pages: {page_count}")
    print(f"Manuscript words: {total_words}")
