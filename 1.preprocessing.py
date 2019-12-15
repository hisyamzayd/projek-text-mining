import re
import nltk
from Sastrawi.Stemmer.StemmerFactory import StemmerFactory
import xlsxwriter
import numpy as np

#Melakukan parsing dari file data latih menjadi dokumen.
#1 file data latih berisi 10 kalimat
#1 kalimat sama dengan 1 dokumen
def parsing(namaFile) :
    dokumen = []

    for i in namaFile :
        kelas = i.replace('data_latih/', '')
        kelas = kelas.replace('.txt', '')
        doc = open(i)
        hasil = doc.read().split('\n')
        for x in hasil:
            dokumen.append([kelas, x])

    return dokumen

#Mengubah semua huruf menjadi huruf kecil dari seluruh dokumen 
def caseFolding(dokumen) :
    result = []

    for i in dokumen :
        result.append([i[0], i[1].lower()])

    return result

#Menghilangkan angka dari seluruh dokumen
def removeNumber(dokumen) :
    numberPattern = '[0-9]'

    result = []

    for i in dokumen :
        result.append([i[0], re.sub(numberPattern, ' ', i[1])])

    return result

#Menghilangkan tanda baca dari seluruh dokumen
def removePunc(dokumen) :
    puncPattern = '[!”#$%&’()*+,-./:;<=>?@[\]^_`{|}~]'

    result = []

    for i in dokumen :
        result.append([i[0], re.sub(puncPattern, ' ', i[1])])

    return result

#Menghilangkan spasi berlebih dari seluruh dokumen
def removeSpace(dokumen) :
    result = []

    for i in dokumen :
        result.append([i[0], i[1].strip()])

    return result

#Membuat kata-kata dari setiap dokumen menjadi bentuk token
#1 kata sama dengan 1 token
def tokenization(dokumen) :
    result = []

    for i in dokumen :
        result.append([i[0], nltk.tokenize.word_tokenize(i[1])])

    return result 

#Menghilangkan kata kata yang dinilai tidak penting sebagai sebuah token. 
#Menggunakan daftar stopword dari library nltk 
def stopwordRemoval(dokumen) :
    stopwordList = set(nltk.corpus.stopwords.words('indonesian'))

    result = []

    for doc in dokumen :
        result1 = []
        for each in doc[1] :
            if (each not in stopwordList) :
                result1.append(each)
        result.append([doc[0], result1])

    return result  

#Mengubah kata-kata dari setiap dokumen menjadi kata dasar
def findStem(dokumen) :
    factory = StemmerFactory()
    stemmer = factory.create_stemmer()

    result = []

    for i in dokumen :
        result1 = []
        for each in i[1] :
            result1.append(stemmer.stem(each))
        result.append([i[0], result1])

    return result

#Membuat susunan array baru dan
#menghitung kemunculan setiap kata 
#untuk persiapan diubah menjadi file .xlsx
def makeWordList(dokumen) :
    wordList = [dokumen[0][1][0]]

    for i in dokumen :
        for each in i[1] :
            if (each not in wordList) :
                wordList.append(each)

    resultWord = [['wordList', wordList]]

    for doc in dokumen :
        eachdoc = []
        for word in wordList :
            count = 0
            for each in doc[1] :
                if (word == each) :
                    count += 1
            eachdoc.append(count)
        resultWord.append([doc[0], eachdoc])
    
    return resultWord

#Membuat file .xlsx
def makeExcel(resultWord, namaFile) :
    wordNumpyArray = np.array(resultWord)

    array = []

    for each in wordNumpyArray :
        array1 = []
        array1.append(each[0])
        for i in each[1] :
            array1.append(i)
        array.append(array1)

    excelArray = np.vstack(array)
    workbook = xlsxwriter.Workbook(namaFile)
    worksheet = workbook.add_worksheet()

    row = 0

    for col, data in enumerate(excelArray) :
        worksheet.write_column(row, col, data)

    workbook.close()
    
#Fungsi Preprocessing secara keseluruhan
def preprocessing(dokumen, namaFile) :
    parsed = parsing(dokumen)
    foldedCase = caseFolding(parsed)
    numberRemoved = removeNumber(foldedCase)
    puncRemoved = removePunc(numberRemoved)
    spaceRemoved = removeSpace(puncRemoved)
    tokened = tokenization(spaceRemoved)
    stopped = stopwordRemoval(tokened)
    stemmed = findStem(stopped)
    wordList = makeWordList(stemmed)
    makeExcel(wordList, namaFile)
    print(wordList)

if __name__ == "__main__" :
    file = ['data_latih/Korupsi.txt', 'data_latih/MotorListrik.txt', 'data_latih/Populasi.txt']
    preprocessing(file, "preprocessing.xlsx")

