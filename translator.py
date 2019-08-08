from xml.dom import minidom
import csv
import pandas as pd
import sys
import os
import xml.etree.ElementTree as ET
from os import path
from nltk.corpus import wordnet

global dictionary

def make_Dict(additionalDictionaryFile):
    dictionary={}
    dictionary.clear()
    mydoc = minidom.parse('en-es.xml')
    scientific_dictionary=csv.reader(open("scientific_dictionary.csv","r"))
    englishword=mydoc.getElementsByTagName('c')
    spanishword=mydoc.getElementsByTagName('d')
    for eng,spa in zip(englishword, spanishword):
        try:
            dictionary[((eng.firstChild.data).lower())]=(spa.firstChild.data).lower()
        except AttributeError:
            dictionary[((eng.firstChild.data).lower())]=''
    for term in scientific_dictionary:
        dictionary[term[0]]=term[1]
    if additionalDictionaryFile is not None:
        if additionalDictionaryFile.split(".")[1]=="xml":
            xmlTree = ET.parse(additionalDictionaryFile)
            elemList=[]
            for elem in xmlTree.iter():
                elemList.append(elem.tag)
            elemList=list(set(elemList))
            additionalDictionary_xml=minidom.parse(additionalDictionaryFile)
            engWords=addtionalDictionary_xml.getElementsByTagName(elemList[0])
            spaWords=addtionalDictionary_xml.getElementsByTagName(elemList[1])
            for eng,spa in zip(engWords, spaWords):
                try:
                    dictionary[((eng.firstChild.data).lower())]=(spa.firstChild.data).lower()
                except AttributeError:
                    dictionary[((eng.firstChild.data).lower())]='' 
        elif additionalDictionaryFile.split(".")[1]=="csv":
            additionalDictionary_csv=csv.reader(open(additionalDictionaryFile,"r"))
            for term in additionalDictionary_csv:
                dictionary[term[0]]=term[1]
        else:
            print("The additional dictionaries must be in .csv or .xml format.")
    return dictionary

def convert_Excel(inFile):
    fileName=inFile.split(".")[0]

    if inFile.find(".xls")!=-1 or inFile.find(".xlsx")!=-1:
        xl=pd.ExcelFile(inFile)
        sheets=xl.sheet_names
        for sheet in sheets:
            df = pd.read_excel(inFile,
                sheet_name=sheet, 
                usecols="A:C", 
                keep_default_na='TRUE', 
                na_values=['NULL'])
            with open((fileName+".csv"), 'a') as inFile_csv:
                df.to_csv(inFile_csv, index=False, header=False)
        inFile_csv.close()
        try:
            file_object = open((fileName+".csv"), 'r')
            lines = csv.reader(file_object, delimiter=',', quotechar='"')
            flag = 0
            data=[]
            for line in lines:
                if line == []:
                    flag=1
                    continue
                else:
                    data.append(line)
            file_object.close()
            if flag==1: #if blank line is present in file
                file_object = open((fileName+".csv"), 'w')
                for line in data:
                    str1 = ','.join(line)
                    file_object.write(str1+"\n")
                file_object.close() 
        except Exception:
            print (e)
        file_object=open((fileName+".csv"), 'r')
    return fileName, inFile

def translate(i, dictionary, translated, dictionaryUsed, keyIDs):
    if dictionary.get(i).find(",")==-1:
        if dictionary.get(i).find("{")!=-1:
            translatedWord=dictionary.get(i)[:dictionary.get(i).find("{")]
        else:
            translatedWord=dictionary.get(i)
        translated.append(translatedWord)
        if keyIDs.index(i)<=keyIDs.index("Abies siberica oil"):
            dictionaryUsed.append("Github")
        else:
            dictionaryUsed.append("Mexican Laws")
    else:
        if dictionary.get(i).split(",")[0].find("{")!=-1:
            translatedWord=dictionary.get(i)[:dictionary.get(i).find("{")]
        else:
            translatedWord=dictionary.get(i).split(",")[0]
        translated.append(translatedWord)
        if keyIDs.index(i)<=keyIDs.index("Abies siberica oil"):
            dictionaryUsed.append("Github")
        else:
            dictionaryUsed.append("Mexican Laws")
            
def main(excel_to_csv, outFile_csv, additionalDictionary=None):
    if path.exists("translations.csv"):
        os.remove("translations.csv")
    fileName, inFile=convert_Excel(excel_to_csv)
    dictionary=make_Dict(additionalDictionary)
    elements_txt=open("elements.txt","r")
    f=(open(outFile_csv,"w"))
    outFile=csv.writer(f, lineterminator="\n")
    outFile.writerow(["English Terms","Spanish Terms","Dictionaries Used","Synonyms Used","Failed Words"])

    elements=[]
    for i in elements_txt:
        elements.append((i.lower())[:-1])
        elements.append(((i.lower())[:-1])+"ii")

    keyIDs=[]
    for i in dictionary.keys():
        keyIDs.append(i)
    correct=0
    incorrect=0
    with open((fileName+".csv")) as file_object:
        terms_csv = csv.reader(file_object)
        for row in terms_csv:
            translated=[]
            wordslist=[]
            dictionaryUsed=[]
            synonymsUsed=[]
            failed=[]
            englishTerms=''
            spanishTerms=''
            dictionaryString=''
            failedWords=''
            synonymString=''
            if row[0].find(":")==-1:
                wordslist.append(row[0])
            else:
                separated=row[0].split(':')
                index=1
                while index<len(separated):
                    separated[index]=" "
                    index=index+2
                for i in separated:
                    if i != " ":
                        wordslist.append(i)
            for i in wordslist:
                if i.find("-")==0:
                    i=i[1:]
                if i.find("ï")==0:
                    i=i[3:]
                if (i in elements):
                        correct=correct+1
                        translated.append(i)
                        dictionaryUsed.append("Chemical Element")
                elif (dictionary.get(i)!="") and i.isalpha() and i.islower() and (i != "mdash") and (i != "middot") and (i != "ndash") and len(i)>2:
                    if (i in dictionary):
                        correct=correct+1
                        translate(i, dictionary, translated, dictionaryUsed, keyIDs)
                    else:
                        synonyms=[]
                        for syn in wordnet.synsets(i):
                            for l in syn.lemmas():
                                synonyms.append(l.name())
                        if len(synonyms)>0:
                            for synonym in synonyms:
                                if (synonym in dictionary):
                                    translate(synonym, dictionary, translated, dictionaryUsed, keyIDs)
                                    synonymsUsed.append(i+"->"+synonym)
                                    break
                        
                        else:
                            incorrect=incorrect+1
                            translated.append(i)
                            dictionaryUsed.append("None")
                            failed.append(i)
                else:
                    translated.append(i)
                    dictionaryUsed.append("None")
                    failed.append(i)
            for eng in wordslist:
                if eng.find("-")==0:
                    eng=eng[1:]
                if eng.find("ï")==0:
                    eng=eng[3:]
                englishTerms=englishTerms+eng+", "
            for i in translated:
                spanishTerms=spanishTerms+i+", "
            for i in dictionaryUsed:
                dictionaryString=dictionaryString+i+", "
            for i in failed:
                failedWords=failedWords+i+", "
            for i in synonymsUsed:
                synonymString=synonymString+i+", "
            outFile.writerow([englishTerms[:-2], spanishTerms[:-2], dictionaryString[:-2], synonymString[:-2], failedWords[:-2]])
    percent=(correct/(correct+incorrect))*100
    print("Percent of terms in " + inFile+ " translated successfully: " + str(round(percent, 2)) + "%.")
    f.close()
    file_object.close()
    if inFile.find(".xls")!=-1 or inFile.find(".xlsx")!=-1:
       os.remove(fileName+".csv")

if __name__ == "__main__":
    main("jacob terms.csv", "translations.csv")
