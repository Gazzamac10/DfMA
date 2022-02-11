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

AecomFamilyPrefix = "_ACM_S_SBM_"
specialprefix = "SP_"

"""CHS"""
CFCHSName = "CFCHS"
CFCHSExcelSheet = dir+"\\CFCHS-secpropsdimsprops-EC3(UKNA)-UK-8-10-2017.xlsx"

CFCHStemplateline = ",Diameter##SECTION_PROPERTY##MILLIMETERS,Wall Nominal Thickness##SECTION_PROPERTY##MILLIMETERS," \
                     "Wall Design Thickness##SECTION_PROPERTY##MILLIMETERS," \
                     "Section Area##SECTION_AREA##SQUARE_CENTIMETERS," \
                     "Perimeter##SURFACE_AREA##SQUARE_METERS_PER_METER," \
                     "Nominal Weight##WEIGHT_PER_UNIT_LENGTH##KILOGRAMS_FORCE_PER_METER," \
                     "Moment of Inertia strong axis##MOMENT_OF_INERTIA##CENTIMETERS_TO_THE_FOURTH_POWER," \
                     "Elastic Modulus strong axis##SECTION_MODULUS##CUBIC_CENTIMETERS," \
                     "Plastic Modulus strong axis##SECTION_MODULUS##CUBIC_CENTIMETERS," \
                     "Torsional Moment of Inertia##MOMENT_OF_INERTIA##CENTIMETERS_TO_THE_FOURTH_POWER," \
                     "Torsional Modulus##SECTION_MODULUS##CUBIC_CENTIMETERS,Section Name Key##other##"
CFCHSindexes = [1,1,4,12,3,6,8,9,10,11]


def CHScreatetextfile(name,ExcelSheet,templateline,indexofparam):

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

    def ifequal(b, a):
        indiceslist = []
        for lA in a:
            counter = 0
            for lB in b:
                if (lA == lB):
                    indiceslist.append(counter)
                counter += 1
        return indiceslist

    series = ifequal(Sections, UniqueSections)

    v = [x[0] - x[1] for x in zip(series[1:], series[:-1])]
    v.append(row2len - series[-1])

    def repeat(a, n):
        list = []
        for i in range(n):
            list.append(a)
        return list

    namelistoflists = []
    for i in range(len(UniqueSections)):
        namelistoflists.append(repeat(UniqueSections[i], v[i]))

    sectionliststart = [j for i in namelistoflists for j in i]
    sectionlistend = [row2[item]for item in colrange]

    def getrange(li, ra):
        return [li[item] for item in ra]

    truncols1 = []
    for i in range(len(columns)):
        truncols1.append(getrange(columns[i], colrange))

    truncols = [truncols1[item] for item in indexofparam]

    Sectioname = []
    for i in range(len(sectionliststart)):
        Sectioname.append([name + sectionliststart[i] +"x"+ sectionlistend[i]])

    j = [item for item in map(list, zip(*truncols))]
    j = [Sectioname[i]+[sectionliststart[i]]+j[i]+Sectioname[i]for i in range(len(Sectioname))]

    fileoutstandard = [templateline]
    fileoutspecials = [templateline]
    for i in range(len(j)):
        if len(filter[i]) > 0:
            fileoutspecials.append(",".join(j[i]))
        else:
            fileoutstandard.append(",".join(j[i]))
    return fileoutstandard,fileoutspecials






CFCHSnamelist = [CFCHSName]
CFCHSexcelsheetlist = [CFCHSExcelSheet]
CFCHStemplatelist = [CFCHStemplateline]
CFCHSindexlist = [CFCHSindexes]


CFCHSnewpaths = [dir + "\\" + AecomFamilyPrefix + str(CFCHSnamelist[i] + ".txt")for i in range(len(CFCHSnamelist))]
CFCHSnewlist = [CHScreatetextfile(CFCHSnamelist[i],CFCHSexcelsheetlist[i],CFCHStemplatelist[i],CFCHSindexlist[i])for i in range(len(CFCHSnamelist))]


print (CFCHSexcelsheetlist)

##for i in range(len(CFCHSnewlist)):
   # SaveoutTxtFile(CFCHSnewlist[i][0],CFCHSnewpaths[i])
