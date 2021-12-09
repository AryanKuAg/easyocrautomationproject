import enum
from os import name
import easyocr
import re
from openpyxl import Workbook

IMAGE_PATH = 'mydata.jpeg'

# reader = easyocr.Reader(['en'], gpu=False)
# result = reader.readtext(IMAGE_PATH, detail = 0)

# for i in result:
#     print(i)

myPreciousData = ['6-Mar-06', '3917', '270457', 'harl', 'chowdary', '27-Mar-06', 'CLOSED', 'phanu_chowdary@yahoo_', 'CmTL', '8404275483', '12-Mar-81', '9948345998', '"8,000 arnually" Hyderabad\' S-', 'cunderabad BEPB Tech', 'Fem ale', '"P HANUMANTHA RAO AMARAM APMS,#304,HNO.8-3-578,', 'YELLAREDDYGUDAAMEERPET,HYD, PIN:S00076', '6-Dec-04', '387', '270411', 'Poonan Pandit', '27-Dec-04', 'CLOSED', 'oonampandit_1 3@rediffm &il.cOm 9323568824', '13-Dec-79', '9920545466', 'HSBC', '"4,17,000 anually"', "Hyderabad' Secunderabad BC Om", 'et ale   "Akash building', 'Flat no', '301, Wing 4', 'eshwant nagar', 'Behing', 'akola Church', 'akola ,Santacruz', 'ea5', 'mumbai', '55"', '16-Nov-06', '3942', '270432', 'oonan nagula', '-Dec-06', 'CLOSED', 'oonamraju@gmail com', '22-Nov-81', '9949062240', "Hyderabad' Secunderabad MBAPGDM", 'Fet ale', '19-Apr-12', '4140', '270680', 'abbaraju venkata sowjanya', '10-May- 12', 'OPEN', 'avsoujanya_2604@yahoo.co.in', '040-40140201', '26-Apr-87', '9985268953', 'TRAINEE FRESHER', "Hyderabad' Secunderabad BEB.Tech", 'Fet ale', '"h-n4-114,flat no302/B`', 'Vijetha Pran eela Pride,', 'Durganagar; Dilsukhnagar, Hyderabad 500060"', '24Aug-07', '3970', '270510', 'Pulijala Arnapurna', '14Sep-07', 'CLOSED', 'anu_337@yahoo.co.in', '30-Aug-82', '9921942912', 'SKM', 'echnologiesPvt Ltd', '"2,04,000 anually', "Hyderabad' Secunderabad BEPB.Tech", 'em ale', '"Sridhar Colony; Karve Nagar,Pune_', '4-Ap-0?', '3956', '270496', 'APARNA D', '25-Apr-07', 'OPEN', 'apafflo', 'duddala@rediff cOt', '9849692325', '10-Apr-82', '9397006551', "Hyderabad' Secunderabad MBAPGDM", 'em ale', '"HNo.l2-11-78, Namalagundu', 'eethaphalmandi, Secunderabad 500061"', '23-Mar-08', '3992', '270532', 'Aparna Mylavarapu', '13-Apr-0?', 'CLOSED', 'aparna', 'mylavarapu@gmail', 'comt', '040-24044087', '30-Mar-83', '9866998318', "Hyderabad' Secunderabad MC APGDCA", 'Fet ale', '"HNO.10-6,SBI COLONY, KOTHAPET % ROADS,HYDERABAD', '500035', 'AP, INDIA"', '14Jun-11', '4109', '270649', 'RAMA DIVYA', '5-Jul-11 CLOSED', 'ram adivya &l 9@gmail com', '20-Jun-36', '9959337238', "Hyderabad' Secunderabad BSc", 'em ale', '"FLAT NO: 401,VIJAYA SUDHA APARTMENTS, BESIDE MG. COLLEGE,NTR NAGAR,HYDERABAD"', '31-Jul-O0', '3712', '270252', 'vudithe parat eshwari', '21-4ug-00', 'parateSha', '2OO@yahoo com', '-Aug-75', '9349611745', 'ICF AI', 'Schod of Inform ation', '"1,99,000 annually"', "Hyderabad' Secunderabad MCAPGDCA", 'Fetale', 'wipro techonologies', '17-Dec-05', '3909', '270449', 'Pavani K', 'JJan-O6 CLOSED', 'pavanik23@rediffmai.com', '040-27173461', '23-Dec-80', '9949973842', 'NIIT', 'Ltd', '"2,24,000 &nually', "Hyderabad' Secunderabad MC APGDCA", 'et ale', '"4-7-8/29_', 'Raghavendra', 'agar, Nacharat, Hyderabad-5O0076"', '17-Sep-08', '4009', '270549', 'phai kum ati', '8-Oct-08', 'OPEN', 'phani naidu@rediffmail cOm', '040-24241555', '24Sep-83', '9948872109', "Hyderabad' Secunderabad MSc", 'Fem ale', '"T3-53, Self Finance €', 'NGO', 'Colory,', 'anasthalipurat,', 'Hyderabad', '500 070.', '19-Apr-1l', '4104', '270644 PRASANTHI YAV ANAMANDHA', '10-May', 'CLOSED', 'prasanthi', 'Vavaflat', 'anda@yahoo cOm', '25-Apr-', '9989872444', "Hyderabad' Secunderabad BEIB Tech", 'Fet ale', '-Oct-08', '4011', '270551', 'priyadarsini paikaray', '28-Oct-08', 'OPEN', 'priyadarsini_paikaray@yahoo', 'in9866458911', '14Oct-83', '9866458911', "'southern rocks", 'mineralsItd,hyderabad"', '"10,000 arnually"', "Hyderabad' Secunderabad BEIB Tech", 'em ale', '"PRIY ADARSINI PAIKARAY,plot no-', '10,anu', 'ladies', 'hosteL SRNagat hyderabad-300038,anchrapradesh"', '15-Apr-0?', '3957', '270497', 'pacmn', 'rokkala', '6-May-0?', 'OPEN', 'psk_padma@yahoo cOm', '21-Apr-82', '9399992835', "Hyderabad' Secunderabad BEB.Tech", 'etaleRPADMA C{O', 'MOSHE LIG B210 ASRAO NAGAR ECIL POST HYDERABAD', '4031', '270571', 'PUSHPA LATHA', '16-May-09', 'OPEN', 'puspha_latha@yahoo com', 'May-84', '9835846033', "Hyderabad' Secunderabad Others", 'em ale', 'HNO 1-7-1/116/1 PRASANTH NAGAR AIWA ISECUNDERABAD', '6-Jul-09 4039', '270579', 'YELLAMMA P', '27-Jul-09', 'OPEN', 'pachipalayatuna@gmail. com', '12-Jul-84', '9963375823', '"2,08,000', "anrally'", "Hyderabad' Secunderabad BEB.Tech", 'em ale', '11-Jan-06', '3911', '270451', 'archana archana', 'eb-06', 'CLOSED', 'archana_+179@yahoo.com', '17-Jan-81', '9985388690', 'Plus Solutions', 'Pvt Ltd "2,05,00 &nually"', 'Hyderabad =', 'cunderabad MC APGDCA', 'Fet ale', '"Fal No.203,Plot No.204 PoliReddy Residence SrinivasNagat colony, Kapra', 'Hyderabad-500062"', '6-Apr-09', '4029', '270569',
                  'kandru radhika', '27-Apr-09', 'OPEN', 'radhikakandru@gn &l.cOm', '12-Apr-84', '9912251881', "Hyderabad' Secunderabad MC APGDCA", 'em ale', 'plotno.20', 'nagar Near mythrivanat Ameerpet Hyderabad', '29-May-08', '3998', '270538', 'reena theegala', '19-Jun-08', 'CLOSED', 'reena', 'theegaa@yahoo coin', '5-Jun-839908196970', 'Bajaj Allianz', '03,000 annually', "Hyderabad' Secunderabad MCAPGDCA", 'etnale   "Plot no-104, Sai praveens', 'dim a residensy Moosapet,hyderabad"', 'OPEN', 'cchnology', 'olory;', 'Apt-09', 'gayathi']


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


print(refinedDataListWithDateFormatted)

# THIS IS FOR EXCEL WORK ---- THIS IS FOR EXCEL WORK ---- THIS IS FOR EXCEL WORK

# def excelGen(mydataList):
#     try:

#         wb = Workbook()
#         ws = wb.active

#         # OUR MAIN ISSUE
#         # This loop is adding data to worksheet
#         for index, element in enumerate(mydataList):
#             temp = mydataList[index]
#             ws.append(['', temp[]])

#     # ws.append([1,3,5,43])

#         wb.save('result' + str(i)+'.xlsx')
#         wb.remove()

#     except Exception as e:
#         print("An exception occurred: ", e)
