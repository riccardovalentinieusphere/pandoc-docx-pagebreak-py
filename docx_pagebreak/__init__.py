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
    # Definizione dei RawBlock per page break, section break e TOC
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
        # Solo RawBlock per gestire i commenti speciali
        if isinstance(elem, pf.RawBlock):
            text = elem.text.strip()

            # 1) Gestione del comando <!--\newpage-->
            if text == "<!--\\newpage-->":
                if doc.format == "docx":
                    pf.debug("Page Break")
                    return self.pagebreak

            # 2) Gestione del comando <!--\toc-->
            elif text == "<!--\\toc-->":
                if doc.format == "docx":
                    pf.debug("Costruzione Indice")

                    # Intestazione “Indice” con stile TOC Heading
                    heading = pf.Div(
                        pf.Para(pf.Str("Indice")),
                        attributes={"custom-style": "TOC Heading"}
                    )

                    # Inserisco il campo TOC OpenXML;
                    # Word applicherà automaticamente TOC 1, TOC 2, … per l’indentazione
                    return [heading, self.toc]

            # 3) Gestione dei blocchi commentati per il titolo
            elif text.startswith("<!") and text.endswith(">"):
                if "title:" in text:
                    pf.debug("Title block detected")
                    match = re.search(r"title:\s*(.+)", text)
                    if match:
                        self.title = match.group(1).strip()
                    # Rimuove il blocco commentato dal flusso
                    return []

        # Rimozione di header H1 in output DOCX (se non gestiti diversamente)
        elif isinstance(elem, pf.Header):
            if doc.format == "docx" and elem.level == 1:
                pf.debug("Removing H1 in docx")
                return []

        # Tutti gli altri elementi passano inalterati
        return elem

    def finalize(self, doc):
        # Se è stato specificato un titolo tramite commento, lo inietto in testa
        if self.title:
            para = pf.Para(pf.Str(self.title))
            styled_para = pf.Div(para, attributes={"custom-style": "Title"})
            doc.content.insert(0, styled_para)

def main(doc=None):
    dp = DocxPagebreak()
    return pf.run_filter(dp.action, prepare=None, finalize=dp.finalize, doc=doc)

if __name__ == "__main__":
    main()
