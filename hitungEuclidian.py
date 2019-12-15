import xlrd
import math

#Mengambil file "tf-idf-with-data-uji.xlsx" kemudian 
#diubah menjadi array dalam python
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

#Menghitung nilai euclidian dari dokumen uji dengan 
#setiap dokumen latih
def hitungEuclidian(dokumen) :
    dataUji = dokumen[len(dokumen)-1][1]
    result = []

    for index, item in enumerate(dokumen) :
        if (index != len(dokumen)-1) :
            eachClass = []
            eachValue = 0
            for j in range(len(dataUji)) :
                each = dataUji[j]-item[1][j]
                value = round(math.pow(each, 2), 3)
                if (value == 0) :
                    value = 0
                eachValue += value
            eachValue = math.sqrt(eachValue)
            eachClass.append(item[0])
            eachClass.append(round(eachValue, 3))
            result.append(eachClass)
    
    return result

#Mengurutkan hasil dokumen dengan 
#nilai terkecil ke nilai terbesar
def minmaxEuclidian(dokumen) :
    check = True
    while check :
        check = False
        for index, item in enumerate(dokumen) :
            if (index != len(dokumen)-1) :
                before = dokumen[index]
                after = dokumen[index+1]
                if (before[1] > after[1]) :
                    check = True
                    transisi = after
                    dokumen[index+1] = before
                    dokumen[index] = transisi
    
    return dokumen

#Mengambil k dokumen
def cropWithK(dokumen, nilai_k) :
    result = []
    for i in range(nilai_k) :
        result.append([dokumen[i][0], dokumen[i][1]])
    return result

#Menggabungkan dokumen dengan kelas yang sama
#kemudian diurutkan dengan kemunculan terbanyak
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

    #combine dokumen dengan kelas yang sama
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

#Fungsi menghitung nilai euclidian secara keseluruhan
def hitungEuclidianAll(nama_file) :
    idf_and_uji_data = prepareDoc(nama_file)
    hasil_euclidian = hitungEuclidian(idf_and_uji_data)
    hasil_min_max = minmaxEuclidian(hasil_euclidian)
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
    hitungEuclidianAll("tf-idf-with-data-uji.xlsx")
