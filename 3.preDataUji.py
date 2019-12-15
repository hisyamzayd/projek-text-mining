import xlrd
import re
import nltk
from Sastrawi.Stemmer.StemmerFactory import StemmerFactory
import numpy as np
import xlsxwriter
import math
from xlutils.copy import copy

#Mengambil file "tf-idf.xlsx" kemudian 
#diubah menjadi array dalam python
def get_tf_idf_doc(namaFile) :
    excelWorkBook = xlrd.open_workbook(namaFile)
    sheet = excelWorkBook.sheet_by_index(0)

    result = []

    #ambil word list dan idf saja
    for i in range(sheet.ncols) :
        eachClass = []
        eachValue = []
        if (i==0 or i==sheet.ncols-1) :
            for j in range(sheet.nrows) :
                if (j==0 or i==0) :
                    value = sheet.cell_value(j,i)
                else :
                    value = float(sheet.cell_value(j,i))
                    if (value == 0) :
                        value = 0
                if (j==0) :
                    eachClass.append(value)
                else :
                    eachValue.append(value)
            eachClass.append(eachValue)
            result.append(eachClass)
    
    return result

#Mengubah semua huruf menjadi huruf kecil dari seluruh dokumen uji
def casefold_data_uji(namaFile) :
    nama = "Data Uji"
    doc = open(namaFile).read()
    result = [nama, doc.lower()]
    return result

#Menghilangkan angka dari seluruh dokumen uji
def remove_number_uji(dokumen) :
    numberPattern = '[0-9]'
    result = [dokumen[0], re.sub(numberPattern, ' ', dokumen[1])]
    return result

#Menghilangkan tanda baca dari seluruh dokumen uji
def remove_punc_uji(dokumen) :
    puncPattern = '[!”#$%&’()*+,-./:;<=>?@[\]^_`{|}~]'
    result = [dokumen[0], re.sub(puncPattern, ' ', dokumen[1])]
    return result

#Menghilangkan spasi berlebih dari seluruh dokumen uji
def remove_space(dokumen) :
    result = [dokumen[0], dokumen[1].strip()]
    return result

#Membuat kata-kata dari setiap dokumen uji menjadi bentuk token
#1 kata sama dengan 1 token
def tokenization_uji(dokumen) :
    result = [dokumen[0], nltk.tokenize.word_tokenize(dokumen[1])]
    return result

#Menghilangkan kata kata yang dinilai tidak penting sebagai sebuah token. 
#Menggunakan daftar stopword dari library nltk 
def stopword_removal_uji(dokumen) :
    stopwordList = set(nltk.corpus.stopwords.words('indonesian'))

    word = []
    for each in dokumen[1] :
        if (each not in stopwordList) :
            word.append(each)
    result = [dokumen[0], word]
    return result

#Mengubah kata-kata dari setiap dokumen uji menjadi kata dasar
def stemming_uji(dokumen) :
    factory = StemmerFactory()
    stemmer = factory.create_stemmer()

    word = []
    for each in dokumen[1] :
        word.append(stemmer.stem(each))
    result = [dokumen[0], word]
    return result

#Menghitung kemunculan setiap kata
def count_word(dokumen, dataTfIdf) :
    wordList = dataTfIdf[0][1]

    hasil = []
    for word in wordList :
        count = 0
        for each in dokumen[1] :
            if (word == each) :
                count += 1
        hasil.append(count)
    result = [dokumen[0], hasil]

    return result

#Menghitung nilai log data uji dengan data "tf-idf.xlsx"
def hitung_log_uji(dokumen) :
    hasilLog = []
    for each in dokumen[1] :
        if (each == 0) :
            hasilLog.append(0)
        else :
            hasil = 1+math.log10(each)
            hasilLog.append(round(hasil, 3))
    result = [dokumen[0], hasilLog]
    return result

#Menghitung nilai tf-idf data uji
def uji_kali_idf(dokumen, word_and_idf) :
    hasilKaliIdf = []
    for index, item in enumerate(word_and_idf[1][1]) :
        hasil = item*dokumen[1][index]
        if (hasil == 0) :
            hasilKaliIdf.append(0)
        else :
            hasilKaliIdf.append(round(hasil, 3))
    result = [dokumen[0], hasilKaliIdf]
    return result

#Menggabungkan file "tf-idf.xlsx" dengan array hasil
#perhitungan data uji dan membuat file .xlsx
def combine(dokumen, namaFile, namaFileOutput) :
    rb = xlrd.open_workbook(namaFile)
    sheet = rb.sheet_by_index(0)

    wb = copy(rb)
    sheet_wb = wb.get_sheet(0)

    #ubah array dokumen menjadi satu dimensi
    dokumen_1_dimensi = []
    dokumen_1_dimensi.append(dokumen[0])
    for each in dokumen[1] :
        dokumen_1_dimensi.append(each)

    #gabung file
    for index, item in enumerate(dokumen_1_dimensi) :
        sheet_wb.write(index, sheet.ncols-1, item)

    wb.save(namaFileOutput)

#Fungsi Preprocessing data uji secara keseluruhan
def pre_data_uji(nama_file_uji, nama_file_tf_idf, nama_file_output) :
    word_and_idf = get_tf_idf_doc(nama_file_tf_idf)

    hasil_case_fold = casefold_data_uji(nama_file_uji)
    hasil_remove_numb = remove_number_uji(hasil_case_fold)
    hasil_remove_punc = remove_punc_uji(hasil_remove_numb)
    hasil_remove_space = remove_space(hasil_remove_punc)
    hasil_tokenization = tokenization_uji(hasil_remove_space)
    hasil_word_removal = stopword_removal_uji(hasil_tokenization)
    hasil_stemming = stemming_uji(hasil_word_removal)
    hasil_count = count_word(hasil_stemming, word_and_idf)
    hasil_hitung_log = hitung_log_uji(hasil_count)
    hasil_uji_kali_idf = uji_kali_idf(hasil_hitung_log, word_and_idf)

    combine(hasil_uji_kali_idf, nama_file_tf_idf, nama_file_output)
    print(hasil_uji_kali_idf)

if __name__ == "__main__" :
    pre_data_uji("DataUji.txt", "tf-idf.xlsx", "tf-idf-with-data-uji.xlsx")
