import os
from PlayActions import send_keys as send
import pandas as pd
import pygetwindow as gw
from TASFiles import TASFileReferences

TAS = TASFileReferences().references

        
def createNewPart(partNumber, description, partClass, partType, mfg = "", mfgNumber = "", vendor = "", vendorNumber = "", specs = "", closeINB = True):
    openEnterInventory(closeINB)
    
    res = enterPartInfo(partNumber, description, partClass, partType)
    
    if res == 1: # part already exists
        return
    
    
    if mfg != "":
        enterMfgInfo(mfg, mfgNumber)
        
    if vendor != "":
        enterVendorInfo(vendor, vendorNumber)
    
    if specs != "":
        addSpecs(specs)
        
    saveFile(closeINB)

def openEnterInventory(closeINB = True):
    window = gw.getWindowsWithTitle("IN-B")
    if len(window) > 0:
        window[0].activate()
        
        if closeINB:
            send([1, "alt x", 1])
        else:
            send([4])
            return
        
    openTASProgram("INB")
        
def enterPartInfo(partNumber, description, partClass, partType) -> int:
    send([partNumber, "enter 2", 3])
    
    pd.DataFrame(columns="696d207469726564".split(" ")).to_clipboard(index=False) 
    send(["ctrl c"])
    
    clipboard = pd.read_clipboard()
    
    print("'" + clipboard.columns[0] + "'", "696d207469726564", clipboard.columns[0] == "696d207469726564")
    
    if clipboard.columns[0] != "696d207469726564":  # theres already a part in there
        send(["alt x", 1])
        return 1
        
    desc = segmentizeSentence(description, 30)
    
    if len(desc) > 2:
        send([description[:30], "enter", description[30:], "enter"])
    elif len(desc) == 2:
        send([desc[0], "enter", desc[1], "enter"])
    else:
        send([description, "enter 2"])
    
    send([partClass, "enter 3"])
    
    send([partType, "enter 23"])
    
    send([partNumber, "enter 2", 1, "enter", "alt s", "enter", 0.5, "enter", 17, "alt x", 0.5])
    
    return 0
    
def enterMfgInfo(mfg, mfgNumber):
    send([ "alt v", 1, "alt a",1])
    send([mfg, "enter", mfgNumber])
    send(["alt s", "enter", 1, "alt x", 1])
    
def enterVendorInfo(vendor, vendorNumber):
    send(["alt v", 1, "alt a", 1])
    send([vendor, "enter", vendorNumber])
    send(["alt s", "enter", 1, "alt x", 1])
    
def addSpecs(specs):
    send(["alt p", 1])
    
    sentences = segmentizeSentence(specs, 30)
    
        
    for sentence in sentences:
        send([sentence, "enter"])
        
    send(["alt x", 1])
    
def saveFile(closeINB):
    send(["alt s", "enter", 5])
    
    if closeINB:
        send(["alt x", 1])
    
    
def quitProgram():
    os._exit(0)
    
def openTASProgram(program):
    send(["focus EVO ~ ERP", 1, "alt m z u a", 1, TAS[program][0], "enter", TAS[program][1]])
    
def abreviateWords(text):
    abbreviations = { "assembly" : "ASSY",
                     "stainless steel" : "STN STL",
                     "ss" : "STN STL",
                     "s.s." : "STN STL",
                     "station" : "STA",
                     "<MOD-DIAM>" : "DIA"
                     }
    
    for word in abbreviations:
        text = text.upper().replace(word, abbreviations[word])
        
    return text

def segmentizeSentence(text, length):
    sentences = []
    
    sents = text.split("\n")
    
    for sent in sents:
        words = sent.split(" ")
        
        sentence = ""
        for word in words:
            if len(sentence + word) > length:
                sentences.append(sentence)
                sentence = word
            else:
                sentence += word + " "
            
        sentences.append(sentence)
        
    return sentences


def isNone(text):
    return text.replace(" ", "").upper().replace("NONE", "").replace("\n", "") == ""
