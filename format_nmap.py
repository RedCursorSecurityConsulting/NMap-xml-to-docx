from docx import Document
from docx.oxml.shared import OxmlElement,  qn
from docx.shared import Inches
from bs4 import BeautifulSoup

def shade_cell(cell, shade):
    tcPr = cell._tc.get_or_add_tcPr()
    tcVAlign = OxmlElement("w:shd")
    tcVAlign.set(qn("w:fill"), shade)
    tcPr.append(tcVAlign)

file_handle = open('./output.xml')

xml = BeautifulSoup(file_handle, features="lxml")

def add_if_exists(output, dictionairy, key):
    if dictionairy.has_attr(key):
        output.append(dictionairy[key])
    else:
        output.append("")

hosts = []

for host in xml.findAll("host"):
    #Need to ask gordon about what he wants done if the host is down.
    if host.status["state"] == "down":
        continue

    addresses = []

    for address in host.findAll('address'):
        addresses.append("{} ({})".format(address["addr"], address["addrtype"]))

    hostnames = []

    for hostname in host.hostnames.findAll('hostname'):
        hostnames.append("{} ({})".format(hostname["name"], hostname["type"]))

    ports_table = []
    for port in host.ports.findAll('port'):

        port_table = []

        # Port | Port | State | Service | Reason | Product | Version | Extra info

        add_if_exists(port_table, port, "portid")
        add_if_exists(port_table, port, "protocol")
        add_if_exists(port_table, port.state, "state")
        add_if_exists(port_table, port.service, "name")
        add_if_exists(port_table, port.service, "product")
        add_if_exists(port_table, port.service, "version")

        ports_table.append(port_table)

    port_message = "The {} ports scanned but not shown above are in state: {}".format(host.extraports["count"], host.extraports["state"])

    hosts.append([addresses, hostnames, ports_table, port_message])

file_handle.close()

document = Document('./table-template.docx')

document.add_heading('Nmap Results', 1)

document.add_paragraph()

for host in hosts:
    table = document.add_table(rows=1, cols=1)

    addresses_table = table.rows[0].cells[0].add_table(rows=1 + len(host[0]), cols=1)
    addresses_table.rows[0].cells[0].text = "Addresses"

    for idx, hostname in enumerate(host[0]):
        addresses_table.rows[1 + idx].cells[0].text = hostname

    hostname_table = table.rows[0].cells[0].add_table(rows=1 + len(host[1]), cols=1)
    hostname_table.rows[0].cells[0].text = "Hostnames"

    for idx, hostname in enumerate(host[1]):
        hostname_table.rows[1 + idx].cells[0].text = hostname

    port_table = table.rows[0].cells[0].add_table(rows=2 + len(host[2]), cols=6)

    port_table.rows[0].cells[0].merge(port_table.rows[0].cells[1])
    port_table.rows[0].cells[0].text = "Port"
    port_table.rows[0].cells[2].text = "State"
    port_table.rows[0].cells[3].text = "Service"
    port_table.rows[0].cells[4].text = "Product"
    port_table.rows[0].cells[5].text = "Version"

    for i in range(1,6):
        port_table.rows[-1].cells[0].merge(port_table.rows[-1].cells[i])

    port_table.rows[-1].cells[0].text = host[3]


    for i, port_list in enumerate(host[2]):
        color = "#00000"

        if port_list[2] == "open":
            color = "#EAF1DD"
        elif port_list[2] == "closed":
            color = "#F2DBDB"

        for j, port_info in enumerate(port_list):
            shade_cell(port_table.rows[1+i].cells[j], color)
            port_table.rows[1 + i].cells[j].text = port_info

    #Fix Styling
    for paragraph in table.rows[0].cells[0].paragraphs:
        paragraph.style = document.styles["NoSpacing"]

    addresses_table.style = document.styles["RedCursor"]
    hostname_table.style = document.styles["RedCursor"]
    port_table.style = document.styles["RedCursor"]

    document.add_paragraph()

document.save('./formatted-nmap.docx')
