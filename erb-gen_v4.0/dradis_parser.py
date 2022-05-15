#!/usr/bin/env python3

####################################################
# Dradis Parser Python Script
# Designed and Written by Salvador Melendez
# WARNING! Any changes made to this file can
#          damage the functionality of the script
####################################################

import os
import sys
import re
import xml.etree.ElementTree as ET

folder_dict = {}

if(len(sys.argv)>1):
    dradis_zip_raw = sys.argv[1]
    dradis_zip = dradis_zip_raw.replace('(', '\(')
    dradis_zip = dradis_zip.replace(')', '\)')
    dradis_zip = dradis_zip.replace(' ', '\ ')
    if 'dradis-' in dradis_zip and os.path.isfile(dradis_zip_raw):
        cwd = os.getcwd()
        #SETUP ERB FOLDER
        folder_name = 'erb'
        if os.path.exists(folder_name):
            msg = 'rm -r ' + folder_name + '/'
            os.system(msg)
        else:
            os.makedirs(folder_name)
        #EXTRACT FILES FROM DRADIS ZIP TO ERB FOLDER
        msg = 'unzip ' + dradis_zip + ' -d ' + folder_name + '/' + ' >/dev/null 2>&1'
        os.system(msg)
        erb_folder = cwd + '/' + folder_name + '/'
        raw_folders = next(os.walk(erb_folder))[1]
        raw_folders.sort()
        for i in range(len(raw_folders)):
            folder_dict[raw_folders[i]] = str(i)
            msg = 'mv ' + erb_folder + '/' + raw_folders[i] + '/ ' + folder_name + '/' + str(i)
            os.system(msg)

        #COPY DRADIS XML TO ERB - WORKING FILE
        dradis_xml = cwd + '/' + folder_name + '/dradis-repository.xml'
        tree = ET.parse(dradis_xml)
        root = tree.getroot()

        #EXTRACT DATA FROM DRADIS XML
        def parse_data(string, start, end):
            tmp = string[string.find(start)+len(start):string.find(end)]
            tmp = tmp.replace('\n', '')
            tmp = ''.join(tmp)
            tmp = tmp.split('#[')
            return tmp[0]
        #'issues'
        titles = {}
        issue_nums = []
        issues = {}
        postures = {}
        mitigations = {}
        for i in root.iter('issue'):
            issue_id = str(i[0].text)
            issue_nums.append(issue_id)
            string = i[2].text
            #TITLE
            if '#[Title]#' in string:
                start = '#[Title]#'
                end = '#[Status]#'
                titles[issue_id] = parse_data(string, start, end)
            else:
                titles[issue_id] = ''
            #ISSUE
            if '#[Description]#' in string:
                start = '#[Description]#'
                end = '#[Posture]#'
                issues[issue_id] = parse_data(string, start, end)
            else:
                issues[issue_id] = ''
            #POSTURE
            if '#[Posture]#' in string:
                start = '#[Posture]#'
                end = '#[Confidentiality]#'
                postures[issue_id] = parse_data(string, start, end)
            else:
                postures[issue_id] = 'NEARSIDER'
            #MITIGATION
            if '#[ShortMitigation]#' in string:
                start = '#[ShortMitigation]#'
                end = '#[Mitigation]#'
                mitigations[issue_id] = parse_data(string, start, end)
            else:
                mitigations[issue_id] = ''

        find_hosts = {}
        find_screenshots = {}
        folder_ids = {}
        for i in issue_nums:
            find_hosts[i] = ''
            find_screenshots[i] = []
            folder_ids[i] = ''
        #'nodes'
        issue_ids = []
        host = ''
        ip_ports = ''
        for i in root.iter('issue-id'):
            issue_ids.append(i.text)
        num_nodes = len(issue_ids)
        count = 0
        for num, i in enumerate(root.iter('content')):
            if '#[HostName]#' in i.text:
                node_id = issue_ids[count]
                string = i.text
                #HOSTS
                if '#[HostName]#' in string:
                    start = '#[HostName]#'
                    end = '#[IPAndPortList]#'
                    host = parse_data(string, start, end)
                else:
                    host = ''
                #IP AND PORT
                if '#[IPAndPortList]#' in string:
                    start = '#[IPAndPortList]#'
                    end = '#[FindingArtifacts]#'
                    ip_ports = parse_data(string, start, end)
                else:
                    ip_ports = ''
                #SCREENSHOTS
                if '#[FindingArtifacts]#' in string:
                    artifacts = string.split('#[FindingArtifacts]#')
                    tmp = re.findall(r'!/(.+?)!', artifacts[1])
                    #SCREENSHOT NAMES
                    screenshots = []
                    for i in tmp:
                        aux = i.split('/')
                        pic_name = aux[-1]
                        if '(' in pic_name and ')' in pic_name and pic_name[-1] == ')':
                            slen = len(pic_name)
                            for j in range(slen-1,-1,-1):
                                if pic_name[j] == '(':
                                    pic_name = pic_name[0:j]
                        screenshots.append(pic_name)
                    find_screenshots[node_id] = screenshots
                    #FOLDER ID
                    if tmp:
                        if '/' in tmp[0]:
                            aux = tmp[0].split('/')
                            folder_ids[node_id] = aux[4]
                        else:
                            folder_ids[node_id] = ''
                    else:
                        folder_ids[node_id] = ''
                else:
                    find_screenshots[node_id] = ''
            #CONCATENATE HOST(S) AND IP(S)
            if host != '' and ip_ports != '':
                host_ip = host + ' / ' + ip_ports
            elif host != '' and ip_ports == '':
                host_ip = host
            elif host == '' and ip_ports != '':
                host_ip = ip_ports
            else:
                host_ip = ''
            find_hosts[node_id] = host_ip
            count+=1

        #CREATE NEW XML FILE
        xml_file = erb_folder + 'findings.xml'
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
        for i in range(len(issue_nums)):
            xml_finding = ET.SubElement(root, 'finding')
            xml_finding.set('uid', str(i))
            xml_folder = ET.SubElement(xml_finding, 'folder')
            if folder_ids[issue_nums[i]]:
                xml_folder.text = cwd + '/erb/' + folder_dict[folder_ids[issue_nums[i]]] + '/'
            else:
                xml_folder.text = ''
            xml_active = ET.SubElement(xml_finding, 'active')
            xml_active.text = str(1)
            xml_rank = ET.SubElement(xml_finding, 'rank')
            xml_rank.text = str(i)
            xml_title = ET.SubElement(xml_finding, 'title')
            xml_title.text = titles[issue_nums[i]]
            xml_hosts = ET.SubElement(xml_finding, 'hosts')
            xml_hosts.text = find_hosts[issue_nums[i]]
            xml_issues = ET.SubElement(xml_finding, 'issues')
            xml_issues.text = issues[issue_nums[i]]
            xml_posture = ET.SubElement(xml_finding, 'posture')
            xml_posture.text = postures[issue_nums[i]]
            xml_mitigation = ET.SubElement(xml_finding, 'mitigation')
            xml_mitigation.text = mitigations[issue_nums[i]]
            xml_include_mitigation = ET.SubElement(xml_finding, 'include_mitigation')
            xml_include_mitigation.text = 'yes'
            xml_screenshots = ET.SubElement(xml_finding, 'screenshots')
            xml_screenshots.text = str(find_screenshots[issue_nums[i]])
        #WRITING XML
        indent(root)
        tree = ET.ElementTree(root)
        tree.write(xml_file, encoding='utf-8', xml_declaration=True)
        #REMOVE ORIGINAL DRADIS XML
        msg = 'rm ' + dradis_xml
        os.system(msg)
    else:
        print('.zip file is invalid... Try again!')
else:
    print('Missing .zip file to parse... Usage: ./dradis_parser.py dradis-export.zip')
