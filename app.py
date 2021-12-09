import enum
from os import name
import easyocr
import re
from openpyxl import Workbook


def imgToexcel(imageName):

    IMAGE_PATH = imageName

    reader = easyocr.Reader(['en'], gpu=False)
    result = reader.readtext(IMAGE_PATH, detail=0)

    myPreciousData = result

    # This function will generate [['6-Mar-06', '27-Mar-06', '12-Mar-81'], ['6-Dec-04', '27-Dec-04', '13-Dec-79'],]

    def listOfListExcelDateEntryGenStepOne(mydataList):
        # print(len(mydataList))

        listOfListWithEachEntry = []
        dateRepetitionTracker = 0
        EntryRepetitionTracker = 0
        # Track that if code joints and make 10 digits number then it won't be conflict by phone numbers
        phoneConflictWithCodeTracker = 0

        for index, element in enumerate(mydataList):

            # It will help to avoid nextElement error
            if len(mydataList) - 1 == index:
                continue

            # Getting the next element and previous element
            previousElement = mydataList[index - 1]
            nextElement = mydataList[index + 1]

            # It will add a new list of entry when dateRepetitionTracker is 0
            if (dateRepetitionTracker == 0):

                listOfListWithEachEntry.append([])
            # 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1

            # Date pattern
            pattern1 = re.compile("\d\d-\w\w\w-\d\d")  # 33-Mar-43
            pattern2 = re.compile("\d-\w\w\w-\d\d")  # 5-Jul-34
            pattern3 = re.compile("\d-\w\w\w\d")  # 6-Mar6
            pattern4 = re.compile("\d-\w\w\w-\d")  # 4-mar-3
            pattern5 = re.compile("\d\w\w\w-\d")  # 6mar-3
            pattern6 = re.compile("-\w\w\w-\d")  # -Dec-06
            pattern7 = re.compile("\d\d-\w\w\w- \d\d")  # 10-May- 12
            pattern8 = re.compile("\d\d\w\w\w-\d\d")  # 24Aug-07
            pattern9 = re.compile("\d-\w\w-\d")  # 4-Ap-0?
            pattern10 = re.compile("\d\d-\w\w\w-\d")  # 13-Apr-0?
            # pattern11 = re.compile("\w\w\w-\d\d")  # May-84

            # if date pattern matches then +1 dateRepetitionTracker9
            if pattern1.match(element) or pattern2.match(element) or pattern3.match(element) or pattern4.match(element) or pattern5.match(element) or pattern6.match(element) or pattern7.match(element) or pattern8.match(element) or pattern9.match(element) or pattern10.match(element):
                dateRepetitionTracker = dateRepetitionTracker + 1  # tracker
                listOfListWithEachEntry[EntryRepetitionTracker].append(
                    element)  # adding to list

            # 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1

            # This is the end statement
            if dateRepetitionTracker == 3 and (pattern1.match(nextElement) or pattern2.match(nextElement) or pattern3.match(nextElement) or pattern4.match(nextElement) or pattern5.match(nextElement) or pattern6.match(nextElement) or pattern7.match(nextElement) or pattern8.match(nextElement) or pattern9.match(nextElement) or pattern10.match(nextElement)):

                # pattern1.findall()
                dateRepetitionTracker = 0
                EntryRepetitionTracker = EntryRepetitionTracker + 1

        return listOfListWithEachEntry

    # This function will give [['6-Mar-06', '27-Mar-06', '12-Mar-81', '3917', '270457'], ['6-Dec-04', '27-Dec-04', '13-Dec-79', '387', '270411'],]

    def ledgerFolioGenStepTwo(mydataList, stepOneList):
        # pattern1 = re.compile("\d\d\d\d\d\d\d\d\d\d")  # 3942270482
        pattern2 = re.compile("\d\d\d")  # 3960
        # pattern3 = re.compile("\d\d\d\d\d\d") # 270500

        # Temp List to hold the extracted data from down loop
        tempData = []

        # Loop for data extraction like 2034 and 343255
        for i in mydataList:
            if pattern2.search(i):
                integerFilter = int(re.search(r'\d+', i).group())
                if len(str(integerFilter)) > 2 and not (str(i).strip()[0] == '9' or str(i).strip()[0] == '8') and (str(i).strip()[0] == '2' or str(i).strip()[0] == '3'):
                    tempData.append(str(integerFilter))

        # This list contain fresh data
        freshList = []
        # Fine extraction from list
        for index, element in enumerate(tempData):
            if len(tempData) - 1 == index:
                continue

            previousElement = tempData[index - 1]
            nextElement = tempData[index + 1]

            if len(str(element)) == 6:
                freshList.append(previousElement)
                freshList.append(element)

        # print(len(freshList))
        # print(len(stepOneList))
        for i in range(0, len(freshList), 2):
            if i//2 >= len(stepOneList) - 1:
                continue
            stepOneList[i//2].append(freshList[i])
            stepOneList[i//2].append(freshList[i + 1])

        return stepOneList

    # print(ledgerFolioGenStepTwo(myPreciousData, listOfListExcelDateEntryGenStepOne(myPreciousData)))

    # This function will give ['6-Mar-06', '27-Mar-06', '12-Mar-81', '3917', '270457', 'CLOSED'], ['6-Dec-04', '27-Dec-04', '13-Dec-79', '387', '270411', 'CLOSED'],

    def openCloseGenStepThree(mydataList, stepTwoList):
        tempList = []  # A list that holds closed and open
        # Criteria
        a = 'CLOSED'
        b = 'OPEN'
        c = 'closed'
        d = 'open'

        # Loop to add open closed data in templist
        for i in mydataList:
            if a in i or c in i:
                tempList.append("CLOSED")
            elif b in i or d in i:
                tempList.append("OPEN")

        # # A list with previous data
        # allDataList = []

        for i, ele in enumerate(stepTwoList):
            if len(stepTwoList) - 1 == i:
                continue

            stepTwoList[i].append(tempList[i])

        return stepTwoList

    # haha = openCloseGen(myPreciousData, ledgerFolioGenStepTwo(myPreciousData, listOfListExcelDateEntryGenStepOne(myPreciousData)))
    # print(haha)

    # This function will generate [['6-Mar-06', '27-Mar-06', '12-Mar-81', '3917', '270457', 'CLOSED', 'HARL'], ['6-Dec-04', '27-Dec-04', '13-Dec-79', '387', '270411', 'CLOSED', 'POONAN PANDIT'],

    def getTheNameStepFour(mydataList, stepThreeList):
        # Pattern
        namePattern = re.compile('[a-z]')

        # This list holds ['HARL', 'POONAN PANDIT', 'ESHWANT NAGAR', 'OONAN NAGULA']
        listOfNames = []

        # This loop generates the element of above list
        for index, i in enumerate(mydataList):
            if len(mydataList) - 1 == index:
                continue
            previousElement = mydataList[index - 1]
            nextElement = mydataList[index + 1]
            if namePattern.search(i) and not '-' in i and not '@' in i and not ',' in i and not '.' in i and not ';' in i and not '"' in i and not "'" in i and not "=" in i and not "ale" in i:
                if previousElement[0] == '2' or previousElement[0] == '3':
                    listOfNames.append(str(i).upper())

        # This loop create involved elements
        for i, element in enumerate(stepThreeList):
            if len(listOfNames) - 1 <= i:
                continue
            stepThreeList[i].append(listOfNames[i])

        return stepThreeList

    # haha = getTheName(myPreciousData, openCloseGenStepThree(myPreciousData, ledgerFolioGenStepTwo(myPreciousData, listOfListExcelDateEntryGenStepOne(myPreciousData))))
    # print(haha)

    # This function will generate [['6-Mar-06', '27-Mar-06', '12-Mar-81', '3917', '270457', 'CLOSED', 'HARL', 'phanu_chowdary@yahoo_'], ['6-Dec-04', '27-Dec-04', '13-Dec-79', '387', '270411', 'CLOSED', 'POONAN PANDIT', 'oonampandit_1 3@rediffm &il.com 9323568824'],]

    def getEmailStepFive(mydataList, stepFourList):
        # Pattern
        emailPattern = re.compile('[a-z0-9^@A-Z]')

        # temp holding emails
        tempEmails = []

        # This loop generates that list
        for i in mydataList:
            if emailPattern.search(i) and "@" in i:
                tempEmails.append(str(i).lower())

        # This loop integrate all things together
        for i, element in enumerate(stepFourList):
            if len(tempEmails) - 1 <= i:
                continue

            stepFourList[i].append(tempEmails[i])

        return stepFourList

    # haha = getEmailStepFive(myPreciousData, getTheNameStepFour(myPreciousData, openCloseGenStepThree(myPreciousData, ledgerFolioGenStepTwo(myPreciousData, listOfListExcelDateEntryGenStepOne(myPreciousData)))))
    # print(haha)

    # This function will give [['6-Mar-06', '27-Mar-06', '12-Mar-81', '3917', '270457', 'CLOSED', 'HARL', 'phanu_chowdary@yahoo_', '8404275483', '9948345998'], ['6-Dec-04', '27-Dec-04', '13-Dec-79', '387', '270411', 'CLOSED', 'POONAN PANDIT', 'oonampandit_1 3@rediffm &il.com 9323568824', '9920545466'],

    def getPhoneNumberStepSix(mydataList, stepFiveList):
        # Pattern
        numberPattern = re.compile('[0-9]')

        # Tracker
        closedOpenTrackerRepetition = -1

        # List of list that holds numbers [['8404275483', '9948345998'], ['9920545466'], ]
        listOfListOfNumbers = []

        # This loop generates the above list items
        for i in mydataList:

            a = 'CLOSED'
            b = 'OPEN'
            c = 'closed'
            d = 'open'

            # This thing to track the index of list
            if a in i or c in i or b in i or d in i:
                closedOpenTrackerRepetition = closedOpenTrackerRepetition + 1
                listOfListOfNumbers.append([])

            if numberPattern.search(i) and (len(i) == 10 or len(i) == 12) and not i.count('-') > 1:
                listOfListOfNumbers[closedOpenTrackerRepetition].append(i)

        # This loop to integrate all those things together
        for i, element in enumerate(stepFiveList):
            if len(listOfListOfNumbers) - 1 <= i:
                continue

            for j in listOfListOfNumbers[i]:
                stepFiveList[i].append(j)

        return stepFiveList

    # haha = getPhoneNumberStepSix(myPreciousData, getEmailStepFive(myPreciousData, getTheNameStepFour(myPreciousData, openCloseGenStepThree(myPreciousData, ledgerFolioGenStepTwo(myPreciousData, listOfListExcelDateEntryGenStepOne(myPreciousData))))))
    # print(haha)

    # And Finally This function will extract the address

    def getAddresStepSeven(mydataList, stepSixList):

        # Patterns
        a = 'CLOSED'
        b = 'OPEN'
        c = 'closed'
        d = 'open'
        # Date pattern
        pattern1 = re.compile("\d\d-\w\w\w-\d\d")  # 33-Mar-43
        pattern2 = re.compile("\d-\w\w\w-\d\d")  # 5-Jul-34
        pattern3 = re.compile("\d-\w\w\w\d")  # 6-Mar6
        pattern4 = re.compile("\d-\w\w\w-\d")  # 4-mar-3
        pattern5 = re.compile("\d\w\w\w-\d")  # 6mar-3
        pattern6 = re.compile("-\w\w\w-\d")  # -Dec-06
        pattern7 = re.compile("\d\d-\w\w\w- \d\d")  # 10-May- 12
        pattern8 = re.compile("\d\d\w\w\w-\d\d")  # 24Aug-07
        pattern9 = re.compile("\d-\w\w-\d")  # 4-Ap-0?
        pattern10 = re.compile("\d\d-\w\w\w-\d")  # 13-Apr-0?
        # Tracker
        closedOpenTrackerRepetition = 0

        # List of list that holds numbers [['8404275483', '9948345998'], ['9920545466'], ]
        listOfListOfAddress = [[]]

        # This loop generates the above list items
        for i in mydataList:

            # This thing to track the index of list
            if a in i or c in i or b in i or d in i:
                closedOpenTrackerRepetition = closedOpenTrackerRepetition + 1
                listOfListOfAddress.append([])

            if not 'CLOS' in i and not 'ale' in i and not 'OPEN' in i and not '@' in i and not (len(i) == 10 or len(i) == 12 or len(i) == 4 or len(i) == 6) and not 'lly' in i and not (pattern1.match(i) or pattern2.match(i) or pattern3.match(i) or pattern4.match(i) or pattern5.match(i) or pattern6.match(i) or pattern7.match(i) or pattern8.match(i) or pattern9.match(i) or pattern10.match(i)) and not 'Tech' in i:
                listOfListOfAddress[closedOpenTrackerRepetition].append(i)

        # popping the first unnecessary element
        listOfListOfAddress.pop(0)

        # refined list
        refinedList = []
        # refining elements in list
        for i, element in enumerate(listOfListOfAddress):
            tempHoldAddress = ''

            for j in listOfListOfAddress[i]:
                tempHoldAddress = tempHoldAddress + ' ' + j

            refinedList.append(tempHoldAddress.strip())

        # Integrating all things together
        for i, element in enumerate(stepSixList):
            if len(refinedList) - 1 == i:
                continue

            stepSixList[i].append(refinedList[i])
        return stepSixList

    refinedDataListWithoutDateFormatted = getAddresStepSeven(myPreciousData, getPhoneNumberStepSix(myPreciousData, getEmailStepFive(myPreciousData, getTheNameStepFour(
        myPreciousData, openCloseGenStepThree(myPreciousData, ledgerFolioGenStepTwo(myPreciousData, listOfListExcelDateEntryGenStepOne(myPreciousData)))))))
    # print(haha)

    def dateFormatter(mydataList):
        for i, element in enumerate(mydataList):
            for index, j in enumerate(mydataList[i]):
                if 'jan' in str(j).lower():
                    # ['6','Jan','34'] or [3, jan]
                    tempSplittedList = str(j).split('-')

                    if str(tempSplittedList[-1]).startswith('0') or str(tempSplittedList[-1]).startswith('1') or str(tempSplittedList[-1]).startswith('2'):
                        tempSplittedList[-1] = '20' + tempSplittedList[-1]
                    else:
                        tempSplittedList[-1] = '19' + tempSplittedList[-1]

                    # tempSplittedList[1] = '01'
                    monthPattern = re.compile('[a-z]')
                    for myindex, month in enumerate(tempSplittedList):
                        if monthPattern.search(month):
                            tempSplittedList[myindex] = '01'
                    mydataList[i][index] = '-'.join(tempSplittedList)
                elif 'feb' in str(j).lower():
                    # ['6','Jan','34'] or [3, jan]
                    tempSplittedList = str(j).split('-')

                    if str(tempSplittedList[-1]).startswith('0') or str(tempSplittedList[-1]).startswith('1') or str(tempSplittedList[-1]).startswith('2'):
                        tempSplittedList[-1] = '20' + tempSplittedList[-1]
                    else:
                        tempSplittedList[-1] = '19' + tempSplittedList[-1]

                    # tempSplittedList[1] = '01'
                    monthPattern = re.compile('[a-z]')
                    for myindex, month in enumerate(tempSplittedList):
                        if monthPattern.search(month):
                            tempSplittedList[myindex] = '02'
                    mydataList[i][index] = '-'.join(tempSplittedList)
                elif 'mar' in str(j).lower():
                    # ['6','Jan','34'] or [3, jan]
                    tempSplittedList = str(j).split('-')

                    if str(tempSplittedList[-1]).startswith('0') or str(tempSplittedList[-1]).startswith('1') or str(tempSplittedList[-1]).startswith('2'):
                        tempSplittedList[-1] = '20' + tempSplittedList[-1]
                    else:
                        tempSplittedList[-1] = '19' + tempSplittedList[-1]

                    # tempSplittedList[1] = '01'
                    monthPattern = re.compile('[a-z]')
                    for myindex, month in enumerate(tempSplittedList):
                        if monthPattern.search(month):
                            tempSplittedList[myindex] = '03'
                    mydataList[i][index] = '-'.join(tempSplittedList)
                elif 'apr' in str(j).lower():
                    # ['6','Jan','34'] or [3, jan]
                    tempSplittedList = str(j).split('-')

                    if str(tempSplittedList[-1]).startswith('0') or str(tempSplittedList[-1]).startswith('1') or str(tempSplittedList[-1]).startswith('2'):
                        tempSplittedList[-1] = '20' + tempSplittedList[-1]
                    else:
                        tempSplittedList[-1] = '19' + tempSplittedList[-1]

                    # tempSplittedList[1] = '01'
                    monthPattern = re.compile('[a-z]')
                    for myindex, month in enumerate(tempSplittedList):
                        if monthPattern.search(month):
                            tempSplittedList[myindex] = '04'
                    mydataList[i][index] = '-'.join(tempSplittedList)
                elif 'may' in str(j).lower():
                    # ['6','Jan','34'] or [3, jan]
                    tempSplittedList = str(j).split('-')

                    if str(tempSplittedList[-1]).startswith('0') or str(tempSplittedList[-1]).startswith('1') or str(tempSplittedList[-1]).startswith('2'):
                        tempSplittedList[-1] = '20' + tempSplittedList[-1]
                    else:
                        tempSplittedList[-1] = '19' + tempSplittedList[-1]

                    # tempSplittedList[1] = '01'
                    monthPattern = re.compile('[a-z]')
                    for myindex, month in enumerate(tempSplittedList):
                        if monthPattern.search(month):
                            tempSplittedList[myindex] = '05'
                    mydataList[i][index] = '-'.join(tempSplittedList)
                elif 'jun' in str(j).lower():
                    # ['6','Jan','34'] or [3, jan]
                    tempSplittedList = str(j).split('-')

                    if str(tempSplittedList[-1]).startswith('0') or str(tempSplittedList[-1]).startswith('1') or str(tempSplittedList[-1]).startswith('2'):
                        tempSplittedList[-1] = '20' + tempSplittedList[-1]
                    else:
                        tempSplittedList[-1] = '19' + tempSplittedList[-1]

                    # tempSplittedList[1] = '01'
                    monthPattern = re.compile('[a-z]')
                    for myindex, month in enumerate(tempSplittedList):
                        if monthPattern.search(month):
                            tempSplittedList[myindex] = '06'
                    mydataList[i][index] = '-'.join(tempSplittedList)
                elif 'jul' in str(j).lower():
                    # ['6','Jan','34'] or [3, jan]
                    tempSplittedList = str(j).split('-')

                    if str(tempSplittedList[-1]).startswith('0') or str(tempSplittedList[-1]).startswith('1') or str(tempSplittedList[-1]).startswith('2'):
                        tempSplittedList[-1] = '20' + tempSplittedList[-1]
                    else:
                        tempSplittedList[-1] = '19' + tempSplittedList[-1]

                    # tempSplittedList[1] = '01'
                    monthPattern = re.compile('[a-z]')
                    for myindex, month in enumerate(tempSplittedList):
                        if monthPattern.search(month):
                            tempSplittedList[myindex] = '07'
                    mydataList[i][index] = '-'.join(tempSplittedList)
                elif 'aug' in str(j).lower():
                    # ['6','Jan','34'] or [3, jan]
                    tempSplittedList = str(j).split('-')

                    if str(tempSplittedList[-1]).startswith('0') or str(tempSplittedList[-1]).startswith('1') or str(tempSplittedList[-1]).startswith('2'):
                        tempSplittedList[-1] = '20' + tempSplittedList[-1]
                    else:
                        tempSplittedList[-1] = '19' + tempSplittedList[-1]

                    # tempSplittedList[1] = '01'
                    monthPattern = re.compile('[a-z]')
                    for myindex, month in enumerate(tempSplittedList):
                        if monthPattern.search(month):
                            tempSplittedList[myindex] = '08'
                    mydataList[i][index] = '-'.join(tempSplittedList)
                elif 'sep' in str(j).lower():
                    # ['6','Jan','34'] or [3, jan]
                    tempSplittedList = str(j).split('-')

                    if str(tempSplittedList[-1]).startswith('0') or str(tempSplittedList[-1]).startswith('1') or str(tempSplittedList[-1]).startswith('2'):
                        tempSplittedList[-1] = '20' + tempSplittedList[-1]
                    else:
                        tempSplittedList[-1] = '19' + tempSplittedList[-1]

                    # tempSplittedList[1] = '01'
                    monthPattern = re.compile('[a-z]')
                    for myindex, month in enumerate(tempSplittedList):
                        if monthPattern.search(month):
                            tempSplittedList[myindex] = '09'
                    mydataList[i][index] = '-'.join(tempSplittedList)
                elif 'oct' in str(j).lower():
                    # ['6','Jan','34'] or [3, jan]
                    tempSplittedList = str(j).split('-')

                    if str(tempSplittedList[-1]).startswith('0') or str(tempSplittedList[-1]).startswith('1') or str(tempSplittedList[-1]).startswith('2'):
                        tempSplittedList[-1] = '20' + tempSplittedList[-1]
                    else:
                        tempSplittedList[-1] = '19' + tempSplittedList[-1]

                    # tempSplittedList[1] = '01'
                    monthPattern = re.compile('[a-z]')
                    for myindex, month in enumerate(tempSplittedList):
                        if monthPattern.search(month):
                            tempSplittedList[myindex] = '10'
                    mydataList[i][index] = '-'.join(tempSplittedList)
                elif 'nov' in str(j).lower():
                    # ['6','Jan','34'] or [3, jan]
                    tempSplittedList = str(j).split('-')

                    if str(tempSplittedList[-1]).startswith('0') or str(tempSplittedList[-1]).startswith('1') or str(tempSplittedList[-1]).startswith('2'):
                        tempSplittedList[-1] = '20' + tempSplittedList[-1]
                    else:
                        tempSplittedList[-1] = '19' + tempSplittedList[-1]

                    # tempSplittedList[1] = '01'
                    monthPattern = re.compile('[a-z]')
                    for myindex, month in enumerate(tempSplittedList):
                        if monthPattern.search(month):
                            tempSplittedList[myindex] = '11'
                    mydataList[i][index] = '-'.join(tempSplittedList)
                elif 'dec' in str(j).lower():
                    # ['6','Jan','34'] or [3, jan]
                    tempSplittedList = str(j).split('-')

                    if str(tempSplittedList[-1]).startswith('0') or str(tempSplittedList[-1]).startswith('1') or str(tempSplittedList[-1]).startswith('2'):
                        tempSplittedList[-1] = '20' + tempSplittedList[-1]
                    else:
                        tempSplittedList[-1] = '19' + tempSplittedList[-1]

                    # tempSplittedList[1] = '01'
                    monthPattern = re.compile('[a-z]')
                    for myindex, month in enumerate(tempSplittedList):
                        if monthPattern.search(month):
                            tempSplittedList[myindex] = '12'
                    mydataList[i][index] = '-'.join(tempSplittedList)
                # else:
                #     # ['6','Jan','34'] or [3, jan]
                #     tempSplittedList = str(j).split('-')

                #     if str(tempSplittedList[-1]).startswith('0') or str(tempSplittedList[-1]).startswith('1') or str(tempSplittedList[-1]).startswith('2'):
                #         tempSplittedList[-1] = '20' + tempSplittedList[-1]
                #     else:
                #         tempSplittedList[-1] = '19' + tempSplittedList[-1]

                #     # tempSplittedList[1] = '01'

                #     mydataList[i][index] = '-'.join(tempSplittedList)

        return mydataList

    refinedDataListWithDateFormatted = dateFormatter(
        refinedDataListWithoutDateFormatted)

    # print(refinedDataListWithDateFormatted)

    # THIS IS FOR EXCEL WORK ---- THIS IS FOR EXCEL WORK ---- THIS IS FOR EXCEL WORK

    def excelGen(mydataList):

        wb = Workbook()
        ws = wb.active

        # OUR MAIN ISSUE
        # This loop is adding data to worksheet
        for index, element in enumerate(mydataList):
            temp = mydataList[index]
            listToAdd = ['', '', '', '', '', '', '', '',
                         '', '', '', '', '', '', '', '', '', '']

            for i, e in enumerate(element):
                # print('this loop runs')
                previousElement = element[i - 1]
                # Adding date
                if '-' in e and ('20' in e or '19 in e'):
                    if listToAdd[3] == '':
                        listToAdd[3] = e
                        continue
                    elif listToAdd[4] == '':
                        listToAdd[4] = e
                        continue
                    elif listToAdd[16] == '':
                        listToAdd[16] = ''
                        continue

                # Adding numbers
                if (len(e) == 10 or len(e) == 12):
                    if listToAdd[13] == '':
                        listToAdd[13] = e
                        continue
                    elif listToAdd[14] == '':
                        listToAdd[14] = e
                        continue

                # Adding ledger
                if (len(e) == 3 or len(e) == 4):
                    if listToAdd[1] == '':
                        listToAdd[1] = e
                        continue

                # Adding folio
                if len(e) == 6:
                    if listToAdd[2] == '':
                        listToAdd[2] = e
                        continue

                # Adding closed and open
                if e == "OPEN" or e == 'CLOSED':
                    if listToAdd[5] == '':
                        listToAdd[5] = e
                        continue

                # Adding name
                if previousElement == 'CLOSED' or previousElement == 'OPEN':
                    if listToAdd[6] == '':
                        listToAdd[6] = e
                        continue

                # Adding Email
                if '@' in e:
                    if listToAdd[15] == '':
                        listToAdd[15] = e
                        continue

                # Adding address
                if len(element) - 1 == i:
                    if listToAdd[7] == '':
                        listToAdd[7] = e
                        continue

            # Appending final thing
            ws.append(listToAdd)

        wb.save(IMAGE_PATH + '.xlsx')

    # This thing creates the excel sheet
    excelGen(refinedDataListWithDateFormatted)
