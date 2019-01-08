#!/usr/bin/env python

import xml.etree.ElementTree as ET
import os
import sys

def parse_xml():
	tree = ET.parse('country_data.xml')
	root = tree.getroot()
	print('root is :' + root.tag)
	
	#access all child of root
	for child in root:
		print(child.tag, child.attrib)
	
	#access child by index
	print(root[0][1].text)
	
	#example of any child search
	#find all neighbor by Element.iter 
	for neighbor in root.iter('neighbor'):
		print(neighbor.attrib)
	
	#example of finding child right away
	#find all country fo root by Element.findall()
	for child in root.findall('country'):
		name = child.get('name')
		rank = child.find('rank').text
		print(name, rank)
		
	#Write xml file
	for rank in root.iter('rank'):
		new_rank = int(rank.text) + 1
		rank.text = str(new_rank)
		rank.set('updated', 'yes')
	tree.write('new_country_xml.xml')
	
	#remove an element by condition
	#and create new xml fiel
	for country in root.findall('country'):
		rank = int(country.find('rank').text)
		if rank > 50:
			root.remove(country)
	tree.write('after_remove.xml')
	
def main():
	parse_xml()
	
	
if __name__=='__main__':
	sys.exit(main())
