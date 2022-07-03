import spms_parser
import print_functions
from docx import Document

# Get information from SPMS
sessions, papers, authors = spms_parser.get_all("get_xml/XML/spms.xml")

# Create document object
document = Document()
document.settings.odd_and_even_pages_header_footer = True

# Add sections to document
print_functions.print_program(document, sessions)
print_functions.print_abstracts(document, sessions)
print_functions.print_authors(document, authors)

# save document
document.save('OUTPUT/example.docx')
