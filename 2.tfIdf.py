import xlrd
import xlsxwriter
import numpy as np
import math

#Mengambil file "preprocessing.xlsx" kemudian 
#diubah menjadi array dalam python
def prepareDoc(namaFile) :
    excelWorkBook = xlrd.open_workbook(namaFile)
    sheet = excelWorkBook.sheet_by_index(0)

    result = []

    for i in range(sheet.ncols) :
        eachClass = []
        eachValue = []
        for j in range(sheet.nrows) :
            if (j==0 or i==0) :
                value = sheet.cell_value(j,i)
            else :
                value = float(sheet.cell_value(j,i))
                if (value == 0) :
                    value = 0
            if (j == 0) :
                eachClass.append(value)
            else :
                eachValue.append(value)
        eachClass.append(eachValue)
        result.append(eachClass)
    
    return result

#Menghitung DF dari setiap dokumen untuk setiap kata
def hitungDF(dokumen) :
    array = []

    for index, item in enumerate(dokumen) :
        if (index > 0) :
            array.append(item[1])

    result = [sum(x) for x in zip(*array)]
    dokumen.append(['df', result])

    return dokumen

#Menghitung IDF dari DF
def hitungIDF(dokumen) :
    panjang = len(dokumen)
    array = dokumen[panjang-1][1]

    result =  []

    for each in array :
        if (each != 0) :
            hasilLog = math.log(panjang-2)/each
            result.append(round(hasilLog, 3))
        else :
            result.append(0)

    dokumen[panjang-1][0] = "idf"
    dokumen[panjang-1][1] = result
    return dokumen

#Menghitung nilai Log
def hitungLog(dokumen) :
    result = []

    result.append([dokumen[0][0], dokumen[0][1]])
    panjang = len(dokumen)

    for index, item in enumerate(dokumen) :
        if (index > 0 and index != panjang-1) :
            array = []
            for each in item[1] :
                if (each == 0) :
                    array.append(0)
                elif (each == 1) :
                    array.append(1)
                else :
                    hasilLog = 1+math.log10(each)
                    array.append(round(hasilLog, 3))
            result.append([item[0], array])

    result.append([dokumen[panjang-1][0], dokumen[panjang-1][1]])
    return result

#Menghitung TF-IDF atau nilai akhir per kata per dokumen
def hitungWeight(dokumen) :
    panjang = len(dokumen)
    arrayIDF = dokumen[panjang-1][1]

    result = []

    result.append([dokumen[0][0], dokumen[0][1]])
    
    for index, item in enumerate(dokumen) :
        if (index > 0 and index != panjang-1) :
            array = []
            for index, each in enumerate(item[1]) :
                hasilKali = each*arrayIDF[index]
                if (hasilKali == 0) :
                    array.append(0)
                else :
                    array.append(round(hasilKali, 3))
            result.append([item[0], array])
    
    result.append([dokumen[panjang-1][0], dokumen[panjang-1][1]])
    return result

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

#Fungsi TF-IDF secara keseluruhan
def hitungTFIDF(dokumenExcel, namaFile) :
    doc = prepareDoc(dokumenExcel)
    hasilDF = hitungDF(doc)
    hasilIDF = hitungIDF(hasilDF)
    hasilLog = hitungLog(hasilIDF)
    hasilWeight = hitungWeight(hasilLog)
    makeExcel(hasilWeight, namaFile)
    print(hasilWeight)

if __name__ == "__main__" :
    file = "preprocessing.xlsx"
    hitungTFIDF(file, "tf-idf.xlsx")
