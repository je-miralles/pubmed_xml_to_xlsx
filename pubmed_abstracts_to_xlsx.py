# pubmed_abstracts_to_xlsx
#
#  Take the pubmed_result.txt file containing a dump of abstracts from a pubmed
# search and process the entries for excel

import xml.etree.ElementTree as ET
import openpyxl as PYXL


## -------------------------------------------------------------#
#
# initialize worksheet
#
def init_worksheet():
    workbook = PYXL.Workbook()

    # grab the active worksheet
    worksheet = workbook.active

    worksheet['A1'] = 'PMID'           # PMID
    worksheet['B1'] = 'DOI'            # ELocationID
    worksheet['C1'] = 'JournalTitle'   # Title
    worksheet['D1'] = 'ArticleTitle'   # ArticleTitle
    worksheet['E1'] = 'Abstract'       # AbstractText

    return (workbook, worksheet)

## -------------------------------------------------------------#

#
# process article fields 
#
def process_article(worksheet, Article, EntryIndex):
    #print('found Article')
    for ArticleField in Article:
        #print(ArticleField.tag)
        if ArticleField.tag == 'ArticleTitle':
            #print('found ArticleTitle')
            Cell = 'D' + str(EntryIndex)
            #print('for cell ', cell)
            #print(ArticleField.text)
            worksheet[Cell] = ArticleField.text
        if ArticleField.tag == 'Abstract':
            for elemAbstract in ArticleField:
                if elemAbstract.tag == 'AbstractText':
                    Cell = 'E' + str(EntryIndex)
                    worksheet[Cell] = elemAbstract.text

#
# process pmid fields 
#
def process_pmid(worksheet, PMIDField, EntryIndex):
    Cell = 'A' + str(EntryIndex)
    worksheet[Cell] = PMIDField.text


## -------------------------------------------------------------#
#
# main program
#
def pm_xml2xlsx(infile, outfile):
    (workbook, worksheet) = init_worksheet()

    pathToXML = infile
    pathToXLSX = outfile

    tree = ET.parse(pathToXML)
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
                        process_pmid(worksheet, SubField, EntryIndex)
                    if SubField.tag == 'Article':
                        process_article(worksheet, SubField, EntryIndex)


    # Save the file
    workbook.save(pathToXLSX)

if __name__ == '__main__':
    print('all good.')
    # ...run automated tests...

