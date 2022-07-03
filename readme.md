# Requirements
- Perl (for JPSP scripts)
- Python 3

# Instructions

### Get SPMS Data

Supply a `conference.config` in the get_xml directory. (See JPSP guide)

Run `perl get_xml/spmsread.pl` to get the latest spms.xml

## Generate word file

Run `pip install -r requirements.txt` to install required python packages

Run `python main.py` to generate an `abstract_booklet.docx` in the `OUTPUT` directory

# Special Thanks

- Volker Schaa for the SPMS read scripts