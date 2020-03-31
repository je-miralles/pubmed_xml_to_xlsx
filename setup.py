from setuptools import setup, find_packages
setup(
    name="pubmed_xml_to_xlsx",
    version="0.1",
    packages=find_packages(),
    # metadata to display on PyPI
    author="Juan Emilio Miralles",
    author_email="jemiralle@gmail.com",
    description="Simple tool to process a PubMed query XML output and store the relevant fields as XLSX.",
    keywords="PubMed Search Query Convert XML XLSX",
    url="https://github.com/je-miralles/pubmed_xml_to_xlsx",
)