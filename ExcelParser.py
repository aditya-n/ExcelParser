import xlrd, xlwt

facultyDetailsDictionary = {}
i = 0

def getRangeElement(oldList, start, end):
    courseAndStartSection = oldList[start]
    if end == -1:
        return courseAndStartSection
    else:
        end_section = int(oldList[end].split('/')[1])%1000;
        range_element = courseAndStartSection + '-' + str(end_section);
        return range_element

def formatCourseAndSectionList(courseAndSectionList):
    new_list = []
    start = end = -1 #Range variables for same course continuous sections
    for idx, val in enumerate(courseAndSectionList):
        if idx == len(courseAndSectionList) - 1: #Ignore last element as this doesn't have a next element to pair with.
            new_list.append(getRangeElement(courseAndSectionList, start, end))
            return new_list
        section = int(val.split('/')[1])
        nextSection = int(courseAndSectionList[idx+1].split('/')[1])

        if nextSection - 1 == section: # if sections are continuous
            if end == -1:
                start = idx
            end = idx + 1
        else:
            new_list.append(getRangeElement(courseAndSectionList, start, end)) #Append range for same course continuous sections
            end = -1
            start = idx

def outputData():
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('test')
    row = 0
    for key in facultyDetailsDictionary:
        courseAndSectionListForEachProf = formatCourseAndSectionList(facultyDetailsDictionary[key])
        for entry in courseAndSectionListForEachProf:
            sheet.write(row, 0, key)
            sheet.write(row, 1, entry)
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

prof2Check = 'Alla Branzburg'
print('Prof\'s List', facultyDetailsDictionary[prof2Check])
print(formatCourseAndSectionList(facultyDetailsDictionary[prof2Check]))

