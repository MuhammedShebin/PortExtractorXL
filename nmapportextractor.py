import argparse
import xml.etree.ElementTree as ET
from openpyxl import Workbook

def extract_nmap_results(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    results = []

    for host in root.findall('host'):
        ip = host.find('address').attrib['addr']
        for port in host.findall('.//port'):
            port_number = port.attrib['portid']
            service_elem = port.find('service')
            if service_elem is not None:
                service = service_elem.attrib.get('name', 'Unknown')
            else:
                service = 'Unknown'
            state = port.find('state').attrib['state']
            if state == 'open':
                results.append((ip, port_number, service, state))

    return results

def write_to_excel(results, output_file):
    wb = Workbook()
    ws = wb.active
    ws.append(['IP', 'Port', 'Service', 'State'])

    for ip, port, service, state in results:
        ws.append([ip, port, service, state])

    wb.save(output_file)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract IP, port, service, and state from Nmap XML output and write to Excel.")
    parser.add_argument("xml_file", help="Path to the Nmap XML output file")
    parser.add_argument("-o", "--output", help="Output Excel file path", default="nmap_results.xlsx")
    args = parser.parse_args()

    nmap_xml_file = args.xml_file
    excel_output_file = args.output

    extracted_results = extract_nmap_results(nmap_xml_file)
    write_to_excel(extracted_results, excel_output_file)
    print("Results written to", excel_output_file)