#!/usr/bin/python3
# -*- coding: utf-8 -*-
import os
import re

def strtr(s, repl):
    pattern = '|'.join(map(re.escape, sorted(repl, key=len, reverse=True)))
    return re.sub(pattern, lambda m: repl[m.group()], s)

def strtr_re(s, patterns):
    for pattern, repl in patterns.items():
        s = re.sub(pattern, repl, s)
    return s

zguni = {
    'ဳ': 'ု', 'ဴ': 'ူ', '္': '်', '်': 'ျ', 'ျ': 'ြ', 'ြ': 'ွ', 'ွ': 'ှ', 'ၚ': 'ါ်',
    'ၠ': '္က', 'ၡ': '္ခ', 'ၢ': '္ဂ', 'ၣ': '္ဃ', 'ၤ': 'င်္', 'ၥ': '္စ', 'ၦ': '္ဆ', 'ၧ': '္ဆ',
    'ၨ': '္ဇ', 'ၩ': '္ဈ', 'ၪ': 'ဉ', 'ၫ': 'ည', 'ၬ': '္ဋ', 'ၭ': '္ဌ', 'ၮ': 'ဍ္ဍ', 'ၯ': 'ဍ္ဎ',
    'ၰ': '္ဏ', 'ၱ': '္တ', 'ၲ': '္တ', 'ၳ': '္ထ', 'ၴ': '္ထ', 'ၵ': '္ဒ', 'ၶ': '္ဓ', 'ၷ': '္န',
    'ၷ': '္ပ', 'ၸ': '္ပ', 'ၹ': '္ဖ', 'ၺ': '္ဗ', 'ၻ': '္ဘ', 'ၼ': '္မ', 'ၽ': 'ျ', 'ၾ': 'ြ',
    'ၿ': 'ြ', 'ႀ': 'ြ', 'ႁ': 'ြ', 'ႂ': 'ြ', 'ႃ': 'ြ', 'ႄ': 'ြ', 'ႅ': '္လ', 'ႆ': 'ဿ',
    'သ္သ': 'ဿ', 'ႇ': 'ှ', 'ႈ': 'ှု', 'ႉ': 'ှူ', 'ႊ': 'ွှ', 'ႏ': 'န', '႐': 'ရ', '႑': 'ဏ္ဍ',
    '႒': 'ဋ္ဌ', '႓': '္ဘ', '႔': '့', '႕': '့', '႗': 'ဋ္ဋ', '၈ၤ': 'ဂင်္', 'ဧ။္': '၏', 'ဧ၊္': '၏',
    '၄င္း': '၎င်း', '၎': '၎င်း', '၎င္း': '၎င်း', 'ေ၀': 'ေဝ', 'ေ၇': 'ေရ', 'ေ၈': 'ေဂ', 'စ်': 'ဈ',
    'ဥာ': 'ဉာ', 'ဥ္': 'ဉ်', 'ၾသ': 'ဩ', 'ေၾသာ္': 'ဪ'
}

zgunicorrect = {
    '\\s+္': '္', '([က-အ])(င်္)': '\\2\\1', '(ေ)([က-အ၀၈၇]{1}္[က-အ၀၈၇]{1})': '\\2\\1', 
    '([ေြ]{1,2})([က-အ၀၈၇]{1})': '\\2\\1', '(ေ)([ျြွှ]+)': '\\2\\1', '(ှ)(ျ)': '\\2\\1', 
    '(ံ)([ုူ])': '\\2\\1', '([ုူ])([ိီ])': '\\2\\1', '(ော)(္[က-အ])': '\\2\\1', 
    '(ဲ)(ွ)': '\\2\\1'
}

def zawgyi_to_unicode(text):
    converted = strtr(text, zguni)
    converted = strtr_re(converted, zgunicorrect)
    return converted

def rename_files_in_directory(directory_path):
    for filename in os.listdir(directory_path):
        file_path = os.path.join(directory_path, filename)
        if os.path.isfile(file_path):
            new_filename = zawgyi_to_unicode(filename)
            new_file_path = os.path.join(directory_path, new_filename)
            os.rename(file_path, new_file_path)
            print(f"Renamed {filename} to {new_filename}")

def main():
    directory_path = input("Enter the path to the directory: ")
    if os.path.isdir(directory_path):
        rename_files_in_directory(directory_path)
    else:
        print("Invalid directory path")

if __name__ == "__main__":
    main()
