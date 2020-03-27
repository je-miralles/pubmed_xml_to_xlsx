# pubmed_abstracts_to_xlsx
#
#  Take the pubmed_result.txt file containing a dump of abstracts from a pubmed
# search and process the entries for excel

import xml.etree.ElementTree as ET
import openpyxl as PYXL
import logging

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

#TODO :
#
# Replace this construction:
#        if ArticleField.tag == 'Abstract':
#              for elemAbstract in ArticleField:
#                  if elemAbstract.tag == 'AbstractText':
#                      Cell = 'E' + str(EntryIndex)
#                      worksheet[Cell] = elemAbstract.text
# with some relevant built-in search method.

#
# process article fields 
#
def process_article(worksheet, Article, EntryIndex):
    #logging.debug("found Article")
    for ArticleField in Article:
        #logging.debug(ArticleField.tag)
        if ArticleField.tag == 'ArticleTitle':
            #logging.debug("found ArticleTitle")
            Cell = 'D' + str(EntryIndex)
            #logging.debug("for cell {}".format(cell))
            #logging.debug(ArticleField.text)
            worksheet[Cell] = ArticleField.text
        if ArticleField.tag == 'Abstract':
            for elemAbstract in ArticleField:
                if elemAbstract.tag == 'AbstractText':
                    Cell = 'E' + str(EntryIndex)
                    worksheet[Cell] = elemAbstract.text
        # if ArticleField.tag == 'Journal':
        #     for elemJournal in ArticleField:
        #         if elemAbstract.tag == 'AbstractText':
        #             Cell = 'E' + str(EntryIndex)
        #             worksheet[Cell] = elemAbstract.text

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
def pm_xml2xlsx(infile, outfile, debug):
    if debug == "DEBUG":
        logging.basicConfig(level=logging.DEBUG)

    (workbook, worksheet) = init_worksheet()

    pathToXML = infile
    pathToXLSX = outfile

    tree = ET.parse(pathToXML)
    root = tree.getroot()       # PubmedArticleSet

    EntryIndex = 1 # start populating after header

    for PubmedArticle in root.iter('PubmedArticle'):
        PubmedArticleField = PubmedArticle.find('MedlineCitation')
        #logging.debug("found PubmedArticle: {} (${})".format(self.name, self.price))
        logging.debug("found PubmedArticle")
        EntryIndex = EntryIndex + 1
        # for PubmedArticleField in PubmedArticle:
        #     if PubmedArticleField.tag == 'MedlineCitation':
        logging.debug('found MedlineCitation')
        for SubField in PubmedArticleField:
            if SubField.tag == 'PMID':
                process_pmid(worksheet, SubField, EntryIndex)
            if SubField.tag == 'Article':
                process_article(worksheet, SubField, EntryIndex)


    # Save the file
    workbook.save(pathToXLSX)

if __name__ == '__main__':
    logging.debug('all good.')
    # ...run automated tests...

