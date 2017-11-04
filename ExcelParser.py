import xlrd, xlwt

facultyDetailsDictionary = {}
i = 0

def sublist(oldList, start, end):
    courseAndStartSection = oldList[start]
    if end == -1:
        return [courseAndStartSection]
    #sectionEnd =
    pass

def formatCourseAndSectionList(courseAndSectionList):
    new_list = courseAndSectionList[:]
    start = end = -1 #Range variables for same course continuous sections
    for idx, val in enumerate(courseAndSectionList):
        if idx == len(courseAndSectionList) - 1: #Ignore last element as this doesn't have a next element to pair with.
            continue
        section = int(val.split('/')[1])
        nextSection = int(courseAndSectionList[idx+1].split('/')[1])

        if nextSection - 1 == section: # if sections are continuous
            if end == -1:
                start = idx
            end = idx + 1
        else:
            new_list.append(sublist(courseAndSectionList, start, end)) #Append range for same course continuous sections
            end = -1
            start = idx

    return new_list

def outputData():
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('test')
    row = 0
    for key in facultyDetailsDictionary:
        sheet.write(row, 0, key)
        sheet.write(row, 1, ', '.join(facultyDetailsDictionary[key]))
        row = row + 1
    workbook.save('output1.xls')

def addNewEntryInDictionary(profName, courseAndSection):
    if profName in facultyDetailsDictionary.keys():
        listForProf = facultyDetailsDictionary[profName]
        listForProf.append(courseAndSection)
    else:
        facultyDetailsDictionary[profName] = [courseAndSection]

# MAIN STARTS HERE
book = xlrd.open_workbook('Spring 2018 Schedule -Lois.xlsx')
# Read input data from excel into dictionary facultyDetails
for sheet in book.sheets():
    for row in range(sheet.nrows):
        i = i+1 # Row number
        try:
            profName = sheet.row(row)[3].value
            if profName in ['Instructor', '']: #Ignore column headers
                continue
            courseAndSection =  str(sheet.row(row)[1].value).split('-')[1] + "/" + str(int(sheet.row(row)[0].value))
            addNewEntryInDictionary(profName, courseAndSection)
        except:
            print('Check row' + str(i)) #Print rows with errors and not added to dictionary

outputData()


print('Bianca Harper\'s List', facultyDetailsDictionary['Bianca Harper'])
print(formatCourseAndSectionList(facultyDetailsDictionary['Bianca Harper']))

