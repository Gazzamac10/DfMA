import os
import xlrd

dir = "C:\\Users\\mccarthyg\\PycharmProjects\\DfMA"

def SaveoutTxtFile(data,Pathout):
    f = open(Pathout,"w")
    for item in data:
        f.write("%s\n" % item)
    f.close()

def ifequal(b, a):
    indiceslist = []
    for lA in a:
        counter = 0
        for lB in b:
            if (lA == lB):
                indiceslist.append(counter)
            counter += 1
    return indiceslist

def getrange(li, ra):
    return [li[item] for item in ra]

def repeat(a, n):
    list = []
    for i in range(n):
        list.append(a)
    return list


CFCHSName = "CFCHS"
CFCHSExcelSheet = dir+"\\CFCHS-secpropsdimsprops-EC3(UKNA)-UK-8-10-2017.xlsx"
CFCHSindexes = [1,3,4,5,6,7]

def CHScreatetextfile(name, ExcelSheet,indexofparam):
    workbook = xlrd.open_workbook(ExcelSheet)
    sheet = workbook.sheet_by_index(0)

    columns = []
    for colx in range(sheet.ncols):
        columns.append(sheet.col_values(colx))
    rows = []
    for rowx in range(sheet.nrows):
        rows.append(sheet.row_values(rowx))

    row1 = columns[0]
    row2 = columns[1]
    row3 = columns[2]
    row2stp = row2[10:len(row2)]

    row2len = len([item for item in row2stp if item.strip()])
    colrange = range(10, row2len + 10)

    filter = [row3[item] for item in colrange]

    Sections = [row1[item] for item in colrange]
    UniqueSections = [item for item in Sections if item.strip()]

    series = ifequal(Sections, UniqueSections)

    v = [x[0] - x[1] for x in zip(series[1:], series[:-1])]
    v.append(row2len - series[-1])

    namelistoflists = []
    for i in range(len(UniqueSections)):
        namelistoflists.append(repeat(UniqueSections[i], v[i]))

    sectionliststart = [j for i in namelistoflists for j in i]
    sectionlistend = [row2[item]for item in colrange]

    truncols1 = []
    for i in range(len(columns)):
        truncols1.append(getrange(columns[i], colrange))

    truncols = [truncols1[item] for item in indexofparam]

    Sectioname = []
    for i in range(len(sectionliststart)):
        Sectioname.append([name + sectionliststart[i] +"x"+ sectionlistend[i]])

    j = [item for item in map(list, zip(*truncols))]
    j = [Sectioname[i]+[sectionliststart[i]]+j[i]+Sectioname[i]for i in range(len(Sectioname))]

    return j

a = CHScreatetextfile(CFCHSName,CFCHSExcelSheet,CFCHSindexes)

print (a)