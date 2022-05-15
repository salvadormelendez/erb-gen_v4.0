#!/usr/bin/env python3

####################################################
# FRIC Parser Python Script
# Designed and Written by Salvador Melendez
# WARNING! Any changes made to this file can
#          damage the functionality of the script
####################################################

import os
import sys
import glob
import xml.etree.ElementTree as ET

if(len(sys.argv)>1):
    fric_folder = sys.argv[1]
    if 'fric_export_' in fric_folder and os.path.exists(fric_folder):
        cwd = os.getcwd()
        #SETUP ERB FOLDER
        folder_name = 'erb'
        if os.path.exists(folder_name):
            msg = 'rm -r ' + folder_name + '/'
            os.system(msg)
            os.makedirs(folder_name)
        else:
            os.makedirs(folder_name)
        #EXTRACT FILES FROM FRIC FOLDER TO ERB FOLDER
        raw_folders = next(os.walk(fric_folder))[1]
        raw_folders.sort()
        source_folder = cwd + '/' + folder_name + '/'
        for i in range(len(raw_folders)):
            msg = 'mkdir ' + source_folder + str(i) + "/"
            os.system(msg)
            msg = "cp -r " + fric_folder + "/'" + raw_folders[i] + "'" + "/* " + source_folder + str(i) + "/"
            os.system(msg)
        #WORK WITH ERB FINDINGS FOLDERS
        dict_chars = {"\\n":"\n", '\\\\':'\\', '%2f':'/', '&amp;':'&', '&quot;':'"', '&#039;':"'", '&lt;':'<', '&gt;':'>'}
        titles = []
        hosts = []
        issues = []
        postures = []
        mitigations = []
        include_mitigations = []
        screenshots = []
        raw_folders = next(os.walk(source_folder))[1]
        int_raw_folders = [int(x) for x in raw_folders]
        int_raw_folders.sort()
        for i in int_raw_folders:
            working_file = source_folder + str(i) + "/*.txt"
            for text_file in glob.iglob(working_file):
                with open(text_file, 'r') as pf:
                    content = pf.readlines()
                    f_text = ''.join(content)
                    #GET FINDING NAME
                    text = f_text[f_text.find('"finding": "'):f_text.find('"_id": "')].split('"')
                    for j in dict_chars:
                        text[-2] = text[-2].replace(j, dict_chars[j])
                    titles.append(text[-2])
                    #GET HOSTS
                    text = f_text[f_text.find('"system": "'):f_text.find('"date": "')].split('"')
                    hosts.append(text[-2])
                    #GET ISSUES
                    text = f_text[f_text.find('"notes": "'):f_text.find('"system": "')].split('"')
                    for j in dict_chars:
                        text[-2] = text[-2].replace(j, dict_chars[j])
                    if text[-2] == '':
                        desc = 'NO DESCRIPTION FOUND FOR THIS FINDING...'
                    else:
                        desc = text[-2]
                    issues.append(desc)
                    postures.append('NEARSIDER')
                    mitigations.append('NO MITIGATION FOUND FOR THIS FINDING...')
                    include_mitigations.append('no')
            msg = 'rm ' + working_file
            os.system(msg)
            msg = source_folder + str(i) + '/'
            ss_list = os.listdir(msg)
            ss_list.sort()
            screenshots.append(ss_list)
        #CREATE NEW XML FILE
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
        #CREATE FINDING
        for i in range(len(int_raw_folders)):
            xml_finding = ET.SubElement(root, 'finding')
            xml_finding.set('uid', str(i))
            xml_folder = ET.SubElement(xml_finding, 'folder')
            xml_folder.text = cwd + '/erb/' + str(i) + '/'
            xml_active = ET.SubElement(xml_finding, 'active')
            xml_active.text = str(1)
            xml_rank = ET.SubElement(xml_finding, 'rank')
            xml_rank.text = str(i)
            xml_title = ET.SubElement(xml_finding, 'title')
            xml_title.text = titles[i]
            xml_hosts = ET.SubElement(xml_finding, 'hosts')
            xml_hosts.text = hosts[i]
            xml_issues = ET.SubElement(xml_finding, 'issues')
            xml_issues.text = issues[i]
            xml_posture = ET.SubElement(xml_finding, 'posture')
            xml_posture.text = postures[i]
            xml_mitigation = ET.SubElement(xml_finding, 'mitigation')
            xml_mitigation.text = mitigations[i]
            xml_include_mitigation = ET.SubElement(xml_finding, 'include_mitigation')
            xml_include_mitigation.text = include_mitigations[i]
            xml_screenshots = ET.SubElement(xml_finding, 'screenshots')
            xml_screenshots.text = str(screenshots[i])
        #WRITING XML
        indent(root)
        tree = ET.ElementTree(root)
        tree.write(xml_file, encoding='utf-8', xml_declaration=True)
    else:
        print('FRIC folder does NOT exist... Try again!')
else:
    print('Missing FRIC folder to parse... Usage: ./fric_parser.py fric_export_*')
