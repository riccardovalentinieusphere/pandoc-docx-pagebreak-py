#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
pandoc-docx-pagebreakpy
Gestisce pagebreak, toc e title da blocchi commentati
Compatibile con blocchi multilinea tipo <!--- ... --->
"""

import panflute as pf
import re


class DocxPagebreak(object):
    pagebreak = pf.RawBlock(
        "<w:p><w:r><w:br w:type=\"page\" /></w:r></w:p>",
        format="openxml"
    )
    sectionbreak = pf.RawBlock(
        "<w:p><w:pPr><w:sectPr><w:type w:val=\"nextPage\" /></w:sectPr></w:pPr></w:p>",
        format="openxml"
    )
    toc = pf.RawBlock(r"""
<w:sdt>
    <w:sdtContent xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:p>
            <w:r>
                <w:fldChar w:fldCharType="begin" w:dirty="true" />
                <w:instrText xml:space="preserve">TOC \o "1-3" \h \z \u</w:instrText>
                <w:fldChar w:fldCharType="separate" />
                <w:fldChar w:fldCharType="end" />
            </w:r>
        </w:p>
    </w:sdtContent>
</w:sdt>
""", format="openxml")

    def __init__(self):
        self.title = None

    def action(self, elem, doc):
        if isinstance(elem, pf.RawBlock):
            text = elem.text.strip()
            
            # Gestione \newpage
            if text == "<!--\\newpage-->":
                if doc.format == "docx":
                    pf.debug("Page Break")
                    return self.pagebreak
            
            # Gestione \toc
            elif text == "<!--\\toc-->":
                if doc.format == "docx":
                    pf.debug("Table of Contents")
                    para = [pf.Para(pf.Str("Table"), pf.Space(), pf.Str("of"), pf.Space(), pf.Str("Contents"))]
                    div = pf.Div(*para, attributes={"custom-style": "TOC Heading"})
                    return [div, self.toc]
            
            # Gestione commenti generici per il titolo
            elif text.startswith("<!") and text.endswith(">"):
                if "title:" in text:
                    pf.debug("Title block detected")
                    match = re.search(r"title:\s*(.+)", text)
                    if match:
                        self.title = match.group(1).strip()
                    return []  # Rimuove il blocco commentato
            
    return elem


    def finalize(self, doc):
        if self.title:
            header = pf.Header(pf.Str(self.title), level=1)
            doc.content.insert(0, header)


def main(doc=None):
    dp = DocxPagebreak()
    return pf.run_filter(dp.action, prepare=None, finalize=dp.finalize, doc=doc)


if __name__ == "__main__":
    main()
