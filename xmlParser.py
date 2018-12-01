import xml.etree.ElementTree as ET
from spellchecker import SpellChecker
import time
import string
import os

trial = {'l':'t', 'H':'ll'}
trialKeys = trial.keys()
currency = ["$"]

spell  = SpellChecker()
spell.word_frequency.load_text_file('./words.txt')

def hasNumbers(inputString):
    return any(char.isdigit() for char in inputString)

def checkCurr(inputString):
    inputString = inputString.split(',')
    for i in inputString:
        if not i.isdigit():
            return False
    return True

def checkDate(inputString):
    return True

def checkErrors(word, n):
    # print(word)
    correction = spell.correction(word.lower())
    if correction.lower() != word.lower():
        # print(spell.candidates(correction))
        w = word
        for i in range(n, len(word)):
            if word[i] in trialKeys:
                word = word[:i] + trial[word[i]] + word[i+1:]
                flag, newWord = checkErrors(word, i+1)
                if flag:
                    return True, newWord
                else:
                    word = w
    else:
        return True, word
    return False, word

def spellCheck(word):
    tempWord = word
    startPunc = ""
    for i in range(len(word)):
        if word[i] in string.punctuation:
            startPunc = startPunc+word[i]
            tempWord = tempWord[1:]
        else:
            break

    endPunc = ""
    for i in range(len(word)-1, -1, -1):
        if word[i] in string.punctuation:
            endPunc = word[i]+endPunc
            tempWord = tempWord[:-1]
        else:
            break

    title = tempWord.istitle()
    upper = tempWord.isupper()
    word = tempWord
    # word = word.lower()
    
    flag, word = checkErrors(word, 0)
                
    # candidates = spell.candidates(word)
    # print(candidates)
    if title:
        word = word.title()
    elif upper:
        word = word.upper()
    return flag, startPunc + word + endPunc

def parseXML(path, xml):
    fileName = xml
    tree  = ET.parse(path + xml)
    root = tree.getroot()

    # for page in root:
    #     for child in page:
    #         # if child.text:
    #         print(child.text)
    #         for b in child:
    #             print(b)

    tempChild = None
    dollarFlag = False
    prevChild = None
    
    toEdit = None
    # for child in root.iter('text'):
    for page in root:
        for child in page:
            if child.tag != 'text':
                continue
            text = None
            if child.text:
                text = child.text
                toEdit = child
            else:
                for b in child:
                    text = b.text
                    toEdit = b
                    break
            if not text or text == " ":
                page.remove(child)
                continue
            text = text.split()
            # print(text)
            for i, t in enumerate(text):
                # print(t)
                if t in currency:
                    continue
                elif t == "s" or t == "S":
                    text[i] = "$"
                elif hasNumbers(t):
                    if ',' in t:
                        if checkCurr(t):
                            pass
                            # print("Currency\n")
                    elif '/' in t:
                        if checkDate(t):
                            pass
                            # print("Date\n")
                else:
                    correction, correct = spellCheck(t)
                    if correction:
                        # print("CORRECTION: ", correct)
                        text[i] = correct
            toEdit.text = ' '.join(text)
            prevChild = child
            # break
    # os.makedirs("./" + fileName)
    tree.write(path + "new_" + fileName)

parseXML("XML/A7FX7D8H/", "A7FX7D8H.pdf.xml")
