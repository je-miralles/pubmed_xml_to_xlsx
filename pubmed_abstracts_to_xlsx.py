# pubmed_abstracts_to_xlsx
#
# Take the pubmed_result.txt file containing a dump of abstracts from a pubmed search and process the entries for excel

import xml.etree.ElementTree as ET
import openpyxl as PYXL


## -------------------------------------------------------------#
#
# initialize worksheet
#
def init_worksheet():
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
def process_article(Article):
    #print('found Article')
    for ArticleField in Article:
        #print(ArticleField.tag)
        if ArticleField.tag == 'ArticleTitle':
            #print('found ArticleTitle')
            Cell = 'D' + str(EntryIndex)
            #print('for cell ', cell)
            #print(ArticleField.text)
            ws[Cell] = ArticleField.text
        if ArticleField.tag == 'Abstract':
            for elemAbstract in ArticleField:
                if elemAbstract.tag == 'AbstractText':
                    Cell = 'E' + str(EntryIndex)
                    ws[Cell] = elemAbstract.text

#
# process pmid fields 
#
def process_pmid(PMIDField):
    Cell = 'A' + str(EntryIndex)
    ws[Cell] = PMIDField.text


## -------------------------------------------------------------#
#
# main program
#
(wb, ws) = init_worksheet()

#pathToXML = 'pubmed_result.xml'

#tree = ET.parse(r'pathToXML')
tree = ET.parse(r'pubmed_result.xml')
root = tree.getroot()       # PubmedArticleSet

EntryIndex = 1 # start populating after header

for PubmedArticle in root:
    print('found PubmedArticle')
    EntryIndex = EntryIndex + 1
    for PubmedArticleField in PubmedArticle:
        if PubmedArticleField.tag == 'MedlineCitation':
            print('found MedlineCitation')
            for SubField in PubmedArticleField:
                if SubField.tag == 'PMID':
                    process_pmid(SubField)
                if SubField.tag == 'Article':
                    process_article(SubField)


# Save the file
wb.save("pubmed_result.xlsx")
