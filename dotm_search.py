#!/usr/bin/env python
"""
Given a directory path, this searches all files in the path for a given text string 
within the 'word/document.xml' section of a MSWord .dotm file.
"""

import sys
import os
import docx2txt

def read_file(file):
    """This function reads the file and returns the text"""
    text = docx2txt.process(file)
    return text



def search_text(dir, search_text):
    """Searches through each file, if file contains a 'search_text' string, 
    returns +- 40 chars from 'search_text'"""

    dotm_files_list = []
    files_with_search_text = []

    for root, dirs, files in os.walk(dir):
        for file in files:
            if file.endswith(".dotm"):
                dotm_files_list.append(os.path.join(root, file))
        
        
    for file in dotm_files_list:
        new_file = read_file(file)
        if search_text in new_file:
            files_with_search_text.append(file)


    for file in files_with_search_text:
        text = read_file(file)
        tmp = text.split()
        text = ' '.join(tmp)
        index_of_char = text.find(search_text)
        print "Match found in file {}".format(file)
        print ' '
        print "Found Text: " + text[index_of_char-40:index_of_char] + text[index_of_char:index_of_char + 41]
        print '-------------------'
    


def main():
    if len(sys.argv) != 4:
        print 'usage: ./dotm_search.py {--dir} directory text'
        sys.exit(1)

    

    option = sys.argv[1]
    directory = sys.argv[2]
    text_to_search = sys.argv[3]
    if option == '--dir':
        search_text(directory, text_to_search)
    else:
        print 'unknown option: ' + option
        sys.exit(1)

if __name__ == "__main__":
    main()