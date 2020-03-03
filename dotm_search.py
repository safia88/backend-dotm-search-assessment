#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Given a directory path, search all files in the path for a given text string
within the 'word/document.xml' section of a MSWord .dotm file.
"""
__author__ = "Safia Ali"

import os
import argparse
import zipfile


def create_parser():
    parser = argparse.ArgumentParser(
        description='Search given text string in dotm file')
    parser.add_argument(
        '--dir', help='Directory of .dotm files to scan', default='.')
    parser.add_argument(
        'text', help='Given text that you want to search for')
    return parser


def search_file(filename, text):
    if not zipfile.is_zipfile(filename):
        print("Error: It's not a zipfile")
        return 0
    with zipfile.ZipFile(filename) as zip:
        with zip.open('word/document.xml') as doc:
            for line in doc:
                indexOfSearchText = line.find(text)
                if indexOfSearchText >= 0:
                    print("Match found in file: {}".format(filename))
                    print("Text Context: " +
                          line[indexOfSearchText-40:indexOfSearchText+40]+" ")
                    return True
            return False


def main():
    # raise NotImplementedError("Your awesome code begins here!")
    parser = create_parser()
    total_file_matches = 0
    total_file_searches = 0
    user_args = parser.parse_args()
    all_file_list = os.listdir(user_args.dir)
    print("Searching " + user_args.dir + " for text " + user_args.text)
    for file in all_file_list:
        if not file.endswith(".dotm"):
            print("File is not a .dotm: " + file)
            continue
        else:
            total_file_searches += 1
            file_path = os.path.join(user_args.dir, file)
            if search_file(file_path, user_args.text):
                total_file_matches += 1
    print("Total dotm files searched: " + str(total_file_searches))
    print("Total dotm files matched: " + str(total_file_matches))


if __name__ == '__main__':
    main()
