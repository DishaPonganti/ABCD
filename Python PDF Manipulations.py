# Python 3.6.4 |Anaconda custom (64-bit)| (default, Mar 12 2018, 20:20:50) 
# Created by DishaPonganti on 3/13/2019

import fitz
from fixtures import *


#open the file and get the total page count
def get_page_count(PDFDoc):
    doc = fitz.open(PDFDoc)
    if doc:
        total_page_count = doc.pageCount
        return total_page_count

    doc.close()
    return -1


#Add highlight to all the keyword instances in the pdf document
def add_highlight(PDFDoc, keyword):

    total_page_count = get_page_count(PDFDoc)
    doc = fitz.open(PDFDoc)
    for n in range(total_page_count):
        page = doc[n]
        # search the keyword and return the instances
        keyword_instances = page.searchFor(keyword, quads=True, hit_max=5000)
        for k in keyword_instances:
            highlight = page.addHighlightAnnot(k)
            doc.saveIncr()

#Add weblink to all the keyword instances in the pdf document
def add_weblink(PDFDoc, keyword, link):

    total_page_count = get_page_count(PDFDoc)
    doc = fitz.open(PDFDoc)
    for n in range(total_page_count):
        page = doc[n]
        # search the keyword and return the instances
        keyword_instances = page.searchFor(keyword, hit_max=5000)
        for k in keyword_instances:
            highlight = page.addHighlightAnnot(k)
            highlight.setColors({"stroke": lightcoral})
            highlight.update()
            var_weblink = page.insertLink(
                {'kind': 2, 'from': fitz.Rect(k.x0, k.y0, k.x1, k.y1), 'uri': link})
            doc.saveIncr()


#Add sticky note to all the keyword instances in the pdf document
def add_sticky_note(PDFDoc, keyword, note):

    total_page_count = get_page_count(PDFDoc)
    doc = fitz.open(PDFDoc)
    for n in range(total_page_count):
        page = doc[n]
        # search the keyword and return the instances
        keyword_instances = page.searchFor(keyword, hit_max=5000)
        for k in keyword_instances:
            highlight = page.addHighlightAnnot(k)
            highlight.setColors({"stroke": greenyellow})
            highlight.update()
            annot = page.addTextAnnot(text=keyword, point=(page.rect.x1 - 30, k.y0))
            annot.setColors({"stroke": greenyellow})
            info = annot.info
            info["title"] = "DishaPonganti"
            info["content"] = note + keyword
            annot.setInfo(info)
            annot.update(fontsize=8)
            doc.saveIncr()

#Add comment to all the keyword instances in the pdf document
def add_comment(PDFDoc, keyword, note):

    total_page_count = get_page_count(PDFDoc)
    doc = fitz.open(PDFDoc)
    for n in range(total_page_count):
        page = doc[n]
        # search the keyword and return the instances
        keyword_instances = page.searchFor(keyword, hit_max=5000)
        for k in keyword_instances:
            highlight = page.addHighlightAnnot(k)
            hinfo = highlight.info
            hinfo['subject'] = "id" + keyword
            highlight.setInfo(hinfo)
            highlight.update()
            doc.saveIncr()

        annot = page.firstAnnot

        total_ant = len(keyword_instances)

        while annot:
            if annot.info['subject'] == "id" + keyword:
                while total_ant > 0:
                    if annot.info['subject'] == "id" + keyword:
                        annot.setBorder({"dashes": [3]})
                        annot.setColors({"stroke": pink3, "fill": (0.75, 0.8, 0.95)})
                        info = annot.info
                        info["title"] = "Comment"
                        info["content"] = comment
                        annot.setInfo(info)
                        r = annot.rect
                        r.x1 = r.x0 + r.width * 1.2
                        r.y1 = r.y0 + r.height * 1.2
                        annot.setRect(r)
                        annot.update()
                        doc.saveIncr()
                        total_ant = total_ant - 1
                    annot = annot.next
            else:
                annot = annot.next

        doc.saveIncr()


if __name__ == '__main__':

    add_highlight(file_name, highlight_keyword)
    add_weblink(file_name, weblink_keyword, link_uri_address)
    add_sticky_note(file_name, sticky_note_keyword, note)
    add_comment(file_name, comment_keyword, comment)