#! python3

import re
import docx
import requests
import os
import sys
import xml.etree.ElementTree as ET

foundAFile = False

files = os.listdir("./")

if len(sys.argv) == 2:
    files = [sys.argv[1]]

for file in files:
    if file.endswith(".doc") or file.endswith(".docx") and not file.startswith("~$"):
        foundAFile = True

        print("Analysing file " + file)
        doc = docx.Document(file)

        goodLinks = []
        badLinks = []

        print("Starting tests.")

        for paragraph in doc.paragraphs:
            text = paragraph.text

            # test your regex here:
            # https://www.regextester.com/93652
            
            urlPattern = '((?:\<(?:[^\<\>]*(?:www|http:|https:){1}[^\<\>]*(?!http)[^\<\>]*\>))|(?:http[s]?:\/\/(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+))'

            root = ET.fromstring(paragraph._element.xml)
            
            fullText = ''

            for e in root.itertext():
                fullText += e

            multiple_match_links = re.findall(urlPattern, fullText)

            links = []

            for link in multiple_match_links:
                strippedLink = link.strip()

                forbiddenChars = [' ','<', '>', '\n']
                for char in forbiddenChars:
                    strippedLink = strippedLink.replace(char, '')

                if strippedLink not in links:
                    links.append(strippedLink)


            if(len(links) > 0):
                for link in links:
                    if link:
                        try:
                            print(" Testing URL: " + link)
                            r = requests.head(link)
                            # print('t ', link, " with code: ", r.status_code)
                            if r.status_code < 400:
                                goodLinks.append(link)
                            else:
                                badLinks.append(link)
                        except:
                            badLinks.append(link)

        print("-> Working URLs: ")
        for link in goodLinks:
            print(link)
            
        print("-> Not Working URLs: ")
        for link in badLinks:
            print(link)

if not foundAFile:
    print("Drag and drop the desired word document to the file linkTester.exe")

input("Press Enter to continue...")
