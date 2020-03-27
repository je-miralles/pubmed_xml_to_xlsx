# pubmed_abstracts_to_xlsx
#
#  Take the pubmed_result.txt file containing a dump of abstracts from a pubmed
# search and process the entries for excel

import xml.etree.ElementTree as ET
import openpyxl as PYXL
import openpyxl.styles as PYXL_S
import logging

## -------------------------------------------------------------#
#
# initialize worksheet
#
def init_worksheet():
    workbook = PYXL.Workbook()

    # grab the active worksheet
    worksheet = workbook.active
    worksheet.page_setup.fitToHeight = 1
    worksheet.column_dimensions['A'].width = 50
    worksheet.column_dimensions['B'].width = 15
    worksheet.column_dimensions['C'].width = 70
    worksheet.column_dimensions['D'].width = 20

    worksheet['A1'] = 'ArticleTitle'   # ArticleTitle
    worksheet['B1'] = 'Author'         # Author
    worksheet['C1'] = 'Abstract'       # AbstractText
    worksheet['D1'] = 'JournalTitle'   # Title
    worksheet['E1'] = 'PMID'           # PMID
    worksheet['F1'] = 'DOI'            # ELocationID

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
    logging.debug('found Article')

    try:
        Cell = 'D' + str(EntryIndex)
        worksheet[Cell].alignment = PYXL_S.Alignment(horizontal='left',
                                                     vertical='top',
                                                     wrap_text=True,
                                                     shrink_to_fit=False)
        worksheet[Cell] = "{}, {}, {} {} ".format(
            Article.find('Journal').find('Title').text,
            Article.find('Journal').find('JournalIssue').find('PubDate').find('Year').text,
            Article.find('Journal').find('JournalIssue').find('PubDate').find('Month').text,
            Article.find('Journal').find('JournalIssue').find('PubDate').find('Day').text)
    except AttributeError:
        worksheet[Cell] = 'NA'
        logging.debug('did not find Journal.Title') 

    try:
        Cell = 'A' + str(EntryIndex)
        worksheet[Cell].alignment = PYXL_S.Alignment(horizontal='left',
                                                     vertical='top',
                                                     wrap_text=True,
                                                     shrink_to_fit=False)
        worksheet[Cell] = Article.find('ArticleTitle').text
    except AttributeError:
        worksheet[Cell] = 'NA'
        logging.debug('did not find ArticleTitle')

    try:
        Cell = 'B' + str(EntryIndex)
        worksheet[Cell].alignment = PYXL_S.Alignment(horizontal='left',
                                                     vertical='top',
                                                     wrap_text=True,
                                                     shrink_to_fit=False)
        worksheet[Cell] = "{}, {}".format(
            Article.find('AuthorList').find('Author').find('LastName').text,
            Article.find('AuthorList').find('Author').find('Initials').text)
    except AttributeError:
        worksheet[Cell] = 'NA'
        logging.debug('did not find ArticleTitle')

    try:
        Cell = 'C' + str(EntryIndex)
        worksheet[Cell].alignment = PYXL_S.Alignment(horizontal='left',
                                                     vertical='top',
                                                     wrap_text=True,
                                                     shrink_to_fit=False)
        worksheet[Cell] = Article.find('Abstract').find('AbstractText').text
    except AttributeError:
        worksheet[Cell] = 'NA'
        logging.debug('did not find AbstractText')

#
# process pmid fields 
#
def process_pmid(worksheet, PMIDField, EntryIndex):
    Cell = 'E' + str(EntryIndex)
    worksheet[Cell].alignment = PYXL_S.Alignment(horizontal='left',
                                                 vertical='top',
                                                 wrap_text=True,
                                                 shrink_to_fit=False)
    worksheet[Cell] = PMIDField.text

#
#logging.debug("found PubmedArticle: {} (${})".format(self.name, self.price))
#logging.debug(ArticleField.tag)
#

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
    worksheet.row_dimensions[EntryIndex].height = 40

    for PubmedArticle in root.iter('PubmedArticle'):
        EntryIndex = EntryIndex + 1
        worksheet.row_dimensions[EntryIndex].height = 40

        MedlineCitation = PubmedArticle.find('MedlineCitation')
        logging.debug('found MedlineCitation')
        process_pmid(worksheet, MedlineCitation.find('PMID'), EntryIndex)
        process_article(worksheet, MedlineCitation.find('Article'), EntryIndex)


    # Save the file
    workbook.save(pathToXLSX)

if __name__ == '__main__':
    logging.debug('all good.')
    # ...run automated tests...

