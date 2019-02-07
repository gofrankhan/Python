#!/usr/bin/env python

from common_util import *
from xml.etree.ElementTree import Element, SubElement
from xml.etree.ElementTree import tostring
from xml.dom.minidom import parseString




def create_xml_folder():
   if not os.path.isdir(XML_FOLDER_PATH):
       os.mkdir(XML_FOLDER_PATH)

def create_summary_xml():
    report_folder_path = XML_FOLDER_PATH + os.path.sep

    summary = Element('summary')
    SubElement(summary, 'total_house_rent').text = str(HOUSE_RENT)
    SubElement(summary, 'total_bazar').text = str(TOTAL_BAZAR)
    SubElement(summary, 'total_bua_bill').text = str(TOTAL_BUA_BILL_5 + AYWAN_BUA_BILL)
    SubElement(summary, 'total_net_bill').text = str(TOTAL_NET_BILL)
    SubElement(summary, 'total_pani_bill').text = str(TOTAL_PANI_BILL)
    SubElement(summary, 'total_electricity_bill').text = str(TOTAL_ELECTRICITY_BILL)
    SubElement(summary, 'total_meal').text = str(TOTAL_MEAL)
    SubElement(summary, 'meal_rate').text = str(MEAL_RATE)
    members = Element('members')
    i = 0
    for m in NAMES:
        child = Element(m)
        SubElement(child, 'house_rent').text = str(HOUSE_RENT/5)
        SubElement(child, 'bua_bill').text = str(TOTAL_BUA_BILL_5/5)
        SubElement(child, 'net_bill').text = str(TOTAL_NET_BILL/5)
        SubElement(child, 'pani_bill').text = str(TOTAL_PANI_BILL/5)
        SubElement(child, 'electricity_bill').text = str(TOTAL_ELECTRICITY_BILL/5)
        SubElement(child, 'meal_cost').text = str(MEAL_COST[i])
        SubElement(child, 'payable').text = str(PAYABLE[i])
        SubElement(child, 'bazar').text = str(BAZAR[i])
        SubElement(child, 'settlement').text = str(PAYABLE[i]-BAZAR[i])
        members.append(child)
        i += 1
    summary.append(members)

    xml = tostring(summary)
    dom = parseString(xml)
    report_file = os.path.join(report_folder_path, 'meal.xml')
    with open(report_file, 'w') as f:
        f.write(dom.toprettyxml())