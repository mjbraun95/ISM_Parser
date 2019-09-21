from bs4 import BeautifulSoup
import operator
import re
# import ISMDataUpdater

PMI_IndustriesArray = ["Machinery","Computer & Electronic Products","Paper Products",
                    "Apparel, Leather & Allied Products","Printing & Related Support Activities",
                    "Primary Metals","Nonmetallic Mineral Products","Petroleum & Coal Products",
                    "Plastics & Rubber Products","Miscellaneous Manufacturing",
                    "Food, Beverage & Tobacco Products","Furniture & Related Products",
                    "Transportation Equipment","Chemical Products","Fabricated Metal Products",
                    "Electrical Equipment, Appliances & Components","Textile Mills","Wood Products"]
NMI_IndustriesArray = ["Retail Trade","Utilities","Arts, Entertainment Recreation",
                    "Other Services","Healthcare and Social Assistance","Food and Accomodations",
                    "Finance and Insurance","Real Estate, Renting and Leasing",
                    "Transport and Warehouse","Mining","Wholesale","Public Admin",
                    "Professional, Science and Technology Services","Information","Education",
                    "Management","Construction","Agriculture, Forest, Fishing and Hunting"]
PMI_SectorsArray = ["ISM MANUFACTURING","NEW ORDERS","PRODUCTION","EMPLOYMENT","SUPPLIER DELIVERIES","INVENTORIES",
                    "CUSTOMER INVENTORIES","PRICES","BACKLOG OF ORDERS","EXPORTS","IMPORTS"]
NMI_SectorsArray = ["ISM NON-MANUFACTURING","BUSINESS ACTIVITY","NEW ORDERS","EMPLOYMENT","SUPPLIER DELIVERIES","INVENTORIES",
                    "PRICES","BACKLOG OF ORDERS","EXPORTS","IMPORTS","INVENTORY SENTIMENT"]
MonthArray = [ "January", "February", "March", "April", "May", "June", "July", "August", 
                "September", "October", "November", "December" ]

SectorsArray = [""]

debug = True
debug_grabClassArrayDebug = True
debug_grabDate = True

def grabSoup( ISM_Page ):
    ISM_Soup = BeautifulSoup(ISM_Page.content, 'html.parser')
    return ISM_Soup

def grabClassArray( ISM_Soup, findThisClass ):
    thisArray = ISM_Soup.findAll( findThisClass )
    if debug_grabClassArrayDebug == True: 
        print("len({}_Array): {}".format( findThisClass, len(thisArray) ) )
        for index, instance in enumerate( thisArray ):
            print("{}_Array[{}].text: {}".format(findThisClass, index, thisArray[index].text))
    return thisArray


def grabDate( ISM_Soup ):
    reportDate = ''
    strong_Array = grabClassArray( ISM_Soup, "strong" )
    releaseMonth = strong_Array[-2].text
    for monthIndex, month in enumerate( MonthArray ):
        if releaseMonth.find( month ) != -1:
            if debug_grabDate == True:
                print( "Found {} from string {}!".format(month, releaseMonth) )
                print( "releaseMonth[-5:-1]: {}".format(releaseMonth[-5:-1]) )
            reportMonthIndex = monthIndex - 1
            reportMonth = MonthArray[ reportMonthIndex ]
            releaseYear = int(releaseMonth[-5:-1])
            if reportMonth == "December":
                reportYear = releaseYear - 1
            else:
                reportYear = releaseYear
            reportDate = "{} {}".format(reportMonth, str(reportYear))
    return reportDate

def grabType( ISM_Soup ):
    h4_Array = grabClassArray( ISM_Soup, "h4" )

    if h4_Array[0].text.find("PMI") != -1:
        print("PMI page detected!")
        Type = "PMI"
    elif h4_Array[0].text.find("NMI") != -1:
        print("NMI page detected!")
        Type = "NMI"
    else:
        print("ERROR: NOT PMI OR NMI. ASSUMING PMI")
        Type = "PMI"
    return Type

def grabISMcomments(ISM_Soup):
    print("Grabbing ISM Comments...")
    lgi_Array = ISM_Soup.findAll("li",{"class":"list-group-item"})
    ISM_CommentList = []
    x = 6
    while x < len(lgi_Array):
        ISM_CommentList.append(lgi_Array[x-6])
        x += 1

    # print("Data array created from {} to {}!".format(ISM_CommentList[0][0],ISM_CommentList[-1][0]))
    print("Number of notes: {}".format(len(ISM_CommentList)))
    return ISM_CommentList
    # return 0

def filterRankingParagraphs( ISM_Soup ):
    debug_filterRankingParagraphs = True
    p_mb3_Array_Filtered = []
    p_mb3_Array_Unfiltered = []
    p_Array = grabClassArray( ISM_Soup, "p" )
    for mb3 in p_Array:
        p_mb3_Array_Unfiltered.append(mb3.text)
        found = False
        if (mb3.text.find("order") != -1 and mb3.text.find("report") != -1 and mb3.text.find("industries") != -1 and mb3.text.find(";") != -1 ):
            p_mb3_Array_Filtered.append(mb3.text)
            found = True
        if debug_filterRankingParagraphs == True:
            try:
                print("\nIn p_mb3_Array_Filtered: {}\n{}".format( found, str(mb3.text) ))
            except UnicodeEncodeError:
                print("\nIn p_mb3_Array_Filtered: {}\nUnicodeEncodeError, can't print str".format( found ))

    if len(p_mb3_Array_Filtered) != 11:
        print("FATAL ERROR: len(p_mb3_Array_Filtered) = {}, should be 11! :(".format( len(p_mb3_Array_Filtered) ) )
        exit()
    return p_Array, p_mb3_Array_Unfiltered, p_mb3_Array_Filtered

def chooseSectorsAndIndustriesArrays( ISM_Soup ):
    if grabType(ISM_Soup) == "PMI":
        thisSectorsArray = PMI_SectorsArray
        thisIndustriesArray = PMI_IndustriesArray

    elif grabType(ISM_Soup) == "NMI":
        thisSectorsArray = NMI_SectorsArray
        thisIndustriesArray = NMI_IndustriesArray
    return thisSectorsArray, thisIndustriesArray

def grabSortedPosNegRankLists( ISM_Soup, p_mb3_Array_Filtered, thisIndustriesArray ):
    debug_grabSortedPosNegRankLists = True
    for p_mb3_FilteredIndex, p_mb3_FilteredText in enumerate( p_mb3_Array_Filtered ):
        if debug_grabSortedPosNegRankLists == True:
            print("\n{}".format(p_mb3_FilteredText))
        industArray = [m.start() for m in re.finditer( "indust", str(p_mb3_FilteredText) )]
        firstRankMin = industArray[0]
        try:
            secondRankMin = industArray[1]
        except IndexError:
            secondRankMin = -1
        if debug_grabSortedPosNegRankLists == True:
            print("firstRankMin: {}".format(firstRankMin))
            print("secondRankMin: {}".format(secondRankMin))
        positiveRankDictionary = {}
        negativeRankDictionary = {}
        subString1 = p_mb3_FilteredText[firstRankMin:secondRankMin]
        

        for thisIndustry in thisIndustriesArray:
            if subString1.find(thisIndustry) != -1:
                positiveRankDictionary[thisIndustry] = subString1.find( thisIndustry )
                
            subString2 = p_mb3_FilteredText[secondRankMin:]
            if subString2.find(thisIndustry) != -1:
                negativeRankDictionary[thisIndustry] = subString2.find( thisIndustry )
                
        sortedPositiveRankList = sorted( positiveRankDictionary.items(), key=lambda kv: kv[1])
        sortedNegativeRankList = sorted( negativeRankDictionary.items(), key=lambda kv: kv[1])
        if debug_grabSortedPosNegRankLists == True:
            print("sortedPositiveRankList: {}".format( sortedPositiveRankList ) )
            print("sortedNegativeRankList: {}".format( sortedNegativeRankList ) )

    return sortedPositiveRankList, sortedNegativeRankList

def grabRankingListArray( p_mb3_Array_Filtered, thisSectorsArray, thisIndustriesArray, sortedPositiveRankList, sortedNegativeRankList ):
    debug_grabRankingListArray = True
    ISM_PositiveRankingList = []
    ISM_NegativeRankingList = []
    ISM_NeutralRankingList = []
    ISM_PositiveRankingListArray = []
    ISM_NegativeRankingListArray = []
    ISM_NeutralRankingListArray = []
    for p_mb3_FilteredIndex, p_mb3_FilteredText in enumerate( p_mb3_Array_Filtered ):

        # ADD NEUTRAL INDUSTRIES TO ENTIRE ARRAY
        for thisIndustry in thisIndustriesArray:
            ISM_NeutralRankingList.append( thisIndustry )

        # POSITIVE INDUSTRIES
        for index, positiveIndustry in enumerate( sortedPositiveRankList ):
            
            quotes = [m.start() for m in re.finditer( "'", str(positiveIndustry) )]
            posStrL = quotes[0]
            posStrR = quotes[1]
            posStr = str(positiveIndustry)[posStrL+1:posStrR]

            if debug_grabRankingListArray == True:
                print("posStr: {}".format( posStr ) )
            ISM_PositiveRankingList.append( posStr )

            # Removes positives from neutral
            ISM_NeutralRankingList.remove( posStr )

        # NEGATIVE INDUSTRIES
        for index, negativeIndustry in enumerate( sortedNegativeRankList ):
            
            quotes = [m.start() for m in re.finditer( "'", str(negativeIndustry) )]
            negStrL = quotes[0]
            negStrR = quotes[1]
            negStr = str(negativeIndustry)[negStrL+1:negStrR]

            if debug_grabRankingListArray == True:
                print("negStr: {}".format( negStr ) )
            ISM_NegativeRankingList.append( negStr )

            # Removes negatives from neutral
            ISM_NeutralRankingList.remove( negStr )

        ISM_PositiveRankingListArray.append( ISM_PositiveRankingList )
        ISM_NegativeRankingListArray.append( ISM_NegativeRankingList )
        ISM_NeutralRankingListArray.append( ISM_NeutralRankingList )

        ISM_NegativeRankingListArray[p_mb3_FilteredIndex].reverse()

        if debug_grabRankingListArray == True:
            print("----")
            print("ISM_PositiveRankingListArray[-1]: {}".format( ISM_PositiveRankingListArray[-1] ) )
            print("ISM_NegativeRankingListArray[-1]: {}".format( ISM_NegativeRankingListArray[-1] ) )
            print(" ISM_NeutralRankingListArray[-1]: {}".format( ISM_NeutralRankingListArray[-1] ) )
            print("----")

        
    finalRankDictionaryArray = [ dict() for x in range(len(p_mb3_Array_Filtered) ) ] 
    for dictIndex, dictionary in enumerate(finalRankDictionaryArray):
        for positiveIndustryIndex, positiveIndustry in enumerate(ISM_PositiveRankingListArray[dictIndex]):
            dictionary[positiveIndustry] = positiveIndustryIndex + 1
        neutralOffset = len(ISM_PositiveRankingListArray[dictIndex])
        for neutralIndustry in ISM_NeutralRankingListArray[dictIndex]:
            dictionary[neutralIndustry] = neutralOffset + 1
        negativeOffset = neutralOffset + len(ISM_NeutralRankingListArray[dictIndex])
        for negativeIndustryIndex, negativeIndustry in enumerate(ISM_NegativeRankingListArray[dictIndex]):
            dictionary[negativeIndustry] = negativeOffset + negativeIndustryIndex + 1

        go = input("Press enter for dictionary {}".format(thisSectorsArray[dictIndex]))
        print(dictionary)
        print("")
            
            
        
    return finalRankDictionaryArray

def grabISMrankings(ISM_Soup):
    debug_grabISMrankings = True
    print("Grabbing ISM Rankings...")
    ISM_RankingList = []

    # ISM_PositiveRankingListArray = []
    # ISM_NegativeRankingListArray = []
    # ISM_NeutralRankingListArray = []

    thisSectorsArray, thisIndustriesArray = chooseSectorsAndIndustriesArrays( ISM_Soup )



    p_Array, p_mb3_Array_Unfiltered, p_mb3_Array_Filtered = filterRankingParagraphs( ISM_Soup )

    sortedPositiveRankList, sortedNegativeRankList = grabSortedPosNegRankLists( ISM_Soup, p_mb3_Array_Filtered, thisIndustriesArray )

    finalRankDictionaryArray = grabRankingListArray( p_mb3_Array_Filtered, thisSectorsArray, thisIndustriesArray, sortedPositiveRankList, sortedNegativeRankList )
    return finalRankDictionaryArray

if __name__ == "__main__":
    # grabISMcomments("https://www.instituteforsupplymanagement.org/ISMReport/MfgROB.cfm?SSO=1")
    grabISMrankings("https://www.instituteforsupplymanagement.org/ISMReport/MfgROB.cfm?SSO=1")