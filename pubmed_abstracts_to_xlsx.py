# pubmed_abstracts_to_xml
#
# Take the pubmed_result.txt file containing a dump of abstracts from a pubmed search and process the entries for excel

import xml.etree.ElementTree as ET
import openpyxl as PYXL


## -------------------------------------------------------------#
#
# initialize worksheet
#
def init_ws():
    wb = PYXL.Workbook()

    # grab the active worksheet
    ws = wb.active

    ws['A1'] = 'PMID'           # PMID
    ws['B1'] = 'DOI'            # ELocationID
    ws['C1'] = 'JournalTitle'   # Title
    ws['D1'] = 'ArticleTitle'   # ArticleTitle
    ws['E1'] = 'Abstract'       # AbstractText

    return (wb, ws)

## -------------------------------------------------------------#
#
# process article fields 
#
def process_article(field):
    print('found Article')
    for elemArticle in field:
        print(elemArticle.tag)
        if elemArticle.tag == 'ArticleTitle':
            print('found ArticleTitle')
            cell = 'D' + str(entry_index)
            print('for cell ', cell)
            print(elemArticle.text)
            ws[cell] = elemArticle.text
        if elemArticle.tag == 'Abstract':
            print('found Abstract')
            for elemAbstract in elemArticle:
                if elemAbstract.tag == 'AbstractText':
                    cell = 'E' + str(entry_index)
                    print('for cell ', cell)
                    print(elemAbstract.text)
                    ws[cell] = elemAbstract.text

#
# process pmid fields 
#
def process_pmid(field):
    print('found PMID')
    cell = 'A' + str(entry_index)
    ws[cell] = field.text


## -------------------------------------------------------------#
#
# main program
(wb, ws) = init_ws()

#pathToXML = 'pubmed_result.xml'

#tree = ET.parse(r'pathToXML')
tree = ET.parse(r'pubmed_result.xml')
root = tree.getroot()       # PubmedArticleSet

entry_index = 1 # start populating after header

for PubmedArticle in root:
    print('found PubmedArticle')
    entry_index = entry_index + 1
    for elemPubmedArticle in PubmedArticle:
        if elemPubmedArticle.tag == 'MedlineCitation':
            print('found MedlineCitation')
            for field in elemPubmedArticle:
                if field.tag == 'PMID':
                    process_pmid(field)
                if field.tag == 'Article':
                    process_article(field)


# Save the file
wb.save("pubmed_result.xlsx")
