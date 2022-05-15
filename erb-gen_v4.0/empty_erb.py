#!/usr/bin/env python3

####################################################
# Empty ERB Python Script
# Designed and Written by Salvador Melendez
# WARNING! Any changes made to this file can
#          damage the functionality of the script
####################################################

import os
import sys
import glob
import xml.etree.ElementTree as ET

cwd = os.getcwd()
#SETUP ERB FOLDER
folder_name = 'erb'
if os.path.exists(folder_name):
    msg = 'rm -r ' + folder_name + '/'
    os.system(msg)
    os.makedirs(folder_name)
else:
    os.makedirs(folder_name)
#CREATE NEW XML FILE - EMPTY
xml_file = cwd + '/' + folder_name + '/findings.xml'
def indent(elem, level=0):
    i = "\n" + level*"    "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "    "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level+1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i
#CREATE FILE STRUCTURE
root = ET.Element('data')
#WRITING XML
indent(root)
tree = ET.ElementTree(root)
tree.write(xml_file, encoding='utf-8', xml_declaration=True)

