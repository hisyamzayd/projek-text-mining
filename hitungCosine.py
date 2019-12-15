# coding: utf-8
import xlrd

#Get xlsx file, turn into array
def prepareDoc(nama_file) :
    excelWorkBook = xlrd.open_workbook(nama_file)
    sheet = excelWorkBook.sheet_by_index(0)

    result = []

    for i in range(sheet.ncols) :
        eachClass = []
        eachValue = []
        if (i!=0) :
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

def hitungCosine(dokumen) :
    dataUji = dokumen[len(dokumen)-1][1]
    result = []

    for index, item in enumerate(dokumen) :
        if (index != len(dokumen)-1) :
            eachClass = []
            eachValue = 0
            for j in range(len(dataUji)) :
                value = round(dataUji[j] * item[1][j], 3)
                if (value == 0) :
                    value = 0
                eachValue += value
            eachClass.append(item[0])
            eachClass.append(round(eachValue, 3))
            result.append(eachClass)
    
    return result

def minmaxCosine(dokumen) :
    check = True
    while check :
        check = False
        for index, item in enumerate(dokumen) :
            if (index != len(dokumen)-1) :
                before = dokumen[index]
                after = dokumen[index+1]
                if (before[1] < after[1]) :
                    check = True
                    transisi = after
                    dokumen[index+1] = before
                    dokumen[index] = transisi
    
    return dokumen

def cropWithK(dokumen, nilai_k) :
    result = []
    for i in range(nilai_k) :
        result.append([dokumen[i][0], dokumen[i][1]])
    return result

def findBestClass(dokumen) :
    #list all class
    allClasses = []
    allClasses.append(dokumen[0][0])
    for index, item in enumerate(dokumen) :
        if (dokumen[index][0] not in allClasses) :
            allClasses.append(dokumen[index][0])

    valueClasses = []
    for i in allClasses :
        value = 0
        for index, item in enumerate(dokumen) :
            if (i==dokumen[index][0]) :
                value += 1
        valueClasses.append(value)

    #combine
    result = []
    for index, item in enumerate(allClasses) :
        result.append([allClasses[index], valueClasses[index]])

    #urutkan
    check = True
    while check :
        check = False
        for index, item in enumerate(result) :
            if (index != len(result)-1) :
                before = result[index]
                after = result[index+1]
                if (before[1] < after[1]) :
                    check = True
                    transisi = after
                    result[index+1] = before
                    result[index] = transisi
    
    return result

def hitungCosineAll(nama_file) :
    idf_and_uji_data = prepareDoc(nama_file)
    hasil_cosine = hitungCosine(idf_and_uji_data)
    hasil_min_max = minmaxCosine(hasil_cosine)
    hasil_crop = cropWithK(hasil_min_max, 3)
    hasil_find_best = findBestClass(hasil_crop)
    print("Nilai K = 3, hasil crop data : ")
    print(hasil_crop)
    print()
    print("Kelompokkan sesuai kelas, hasil :")
    print(hasil_find_best)
    print()
    print("Maka data uji termasuk kelas : " + hasil_find_best[0][0])

if __name__ == "__main__" :
    hitungCosineAll("tf-idf-with-data-uji.xlsx")
