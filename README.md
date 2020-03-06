# Nmap to Docx

This tool converts a given nmap xml file to a more readable format for reports.

### Prerequisites

```text
python-docx (tested with 0.8.10)
bs4
```

### Usage

```text
usage: format_nmap.py [-h] infile outfile

Process Nmap XML file and produce docx table

positional arguments:
  infile      Input file (e.g. target.xml)
  outfile     Output file (e.g. document.docx)

  optional arguments:
    -h, --help  show this help message and exit
```

