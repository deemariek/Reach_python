import os
import pandas as pd
import xlrd
from collections import Counter
pd.core.format.header_style = None
import csv
import operator
from collections import Counter

# update filenames for each month's appropriate folder 
fileIn = 'AoO_Round7_Nov2015_Dataset_NonAnonymised.xlsx'
fileSheet = 'AoO_Round7_Nov2015_Dataset_NonA'
fileOut = 'AoO_Round7_Nov2015_subdistrictAgg.csv'

# update each month according to the questions required for the factsheet
aggregateQuestions = ['Displacement/QB001_num_Pre_conf_popul_remained_last_day_pre_m',
                      'Health/QE017_Where_women_delivered_babies',
                      'NFI/QD002_source_electricity_used_most_hours_pre_m',
                      'WASH/QF006_most_way_people_disposed_garbage_pre_m',
                      'Shelter/G_QC001/QC001_most_type_of_housing_this_village_pre_m',
                      'Education/G_QI001/QI001_Primary_school_available_pre_m']

# remove the aggQuestions string so that it's a bit more user friendly
aggregate_Questions = ['Displacement_QB001', 'Health_QE017', 'NFI_QD002', 'WASH_QF006', 'Shelter_QC001', 'Education_QI001']

# responses that shouldn't be included in aggregations
noResponse = ['No IDPs', 'No information', 'No concensus', 'No IDPs last winter', 'NA']

# ################################################

def aggregateToSubDistrict(fileIn, fileSheet):
    """ This function takes a filename and the sheet name as arguments.
    The function aggregates through each SD and aggregated the values for each aggregateQuestion.
    The function returns a list of the aggregated values for each subdistrict.
    """
    # read the data in as a DataFrame
    df = pd.DataFrame()
    data = pd.read_excel(fileIn, fileSheet)
    df = df.append(data) 
    print "{0} loaded".format(fileIn)
    # empty list where the responses will eventually be appended to
    allResponses = []
    # locate the column on which you the script to aggregate, in this case 'Subdistrict_Assessed_location'
    subdistricts = df['Subdistrict_Assessed_location']
    # get list of unique values in the column
    subdistrictList = pd.unique(df.Subdistrict_Assessed_location.ravel())
    # get a Series containing the coverage counts for each subdistrict
    subCounts = subdistricts.value_counts()
    # iterate through the unique subdistrict values
    for sub in subdistrictList:
       # modeResponse is a string to which each subdistricts name, coverage and aggregated values are written to
       modeResponse = "{0},{1},".format(sub, subCounts[sub])
       # iterate through each of the questions of interest and get SD questions and confidence responses
       for aggQ in aggregateQuestions:
           responseList = []
           #confidenceResponse = []
           # a dict to contain SD question and conf responses
           resConfs = {}
           # start to iterate through the dataframe rows
           for index, row in df.iterrows():
              # find rows that relate to our current subdistrict
              if row['Subdistrict_Assessed_location'] == sub:
                   # get the question from the current subdistrict for the current aggQ
                   question = row[aggQ]
                   confidence = row['Conf_'+aggQ]
                   # append values to list
                   responseList.append(str(question))
                   #confidenceResponse.append(str(confidence))
                   # only include responses on which we want to aggregate
                   if question not in noResponse:
                       # add to dict, add value to key if key already there
                       if question not in resConfs:
                           resConfs[question] = [confidence]
                       else:
                           resConfs[question].append(confidence)
                   else:
                        pass
           # only one response - therefore no need to aggregate
           if len(responseList) == 1:
                # check if it's a value we want to include or not
                if responseList[0] in noResponse:
                    modeResponse = modeResponse + "{0},".format('NA')
                else: 
                    modeResponse = modeResponse + "{0},".format(responseList[0])
           # we need to aggregate if there are more two responses for one subdistrict
           elif len(responseList) > 1:
                # check there are responses for that subdistrict
                if resConfs:
                    maxConfs = {}
                    # create new dict with questions and sum of confidences
                    for key, values in resConfs.iteritems():
                        sumValue = sum(values)
                        if key not in maxConfs:
                            maxConfs[key] = [sumValue]
                        else:
                            maxConfs[key].append(sumValue)
                    # get max value from confidence sums
                    maxValue = max(maxConfs.iteritems(), key=operator.itemgetter(1))[1]
                    # a list that will contain the final aggregated value for each subdistrict
                    keys = []
                    for key, values in maxConfs.iteritems():
                        if values == maxValue:
                            keys.append(key)
                        keyStr = "/".join(keys)
                    # we don't report on SDs that have more than two equally reported values
                    # if there are more than two items in the keys list then a value of 'No Concensus' is returned
                    if len(keys) <= 2:
                        modeResponse = modeResponse + "{0},".format(keyStr)
                    # return NC (No Concensus) if there are more than two equal options returned 
                    elif len(keys) > 2:
                        modeResponse = modeResponse + "{0},".format("NC")
                # if there are no responses, write 'NA' modeResponse
                else:
                    modeResponse = modeResponse + "{0},".format('NA')
           else:
                pass
       # the aggregated values for each aggregateQuestion are appended to allResponses after processing each subdistrict
       allResponses.append(modeResponse)
    print "Aggregated to subdistrict"
    # the full list of SDs and aggregated values is returned by the function
    return allResponses

# ################################################

def writeResponsesToCSV(inputData, outputFile):
    """ This function takes a list of data and an output filename as arguments.
    The function writes the data to the output file.
    districtHeading is the variables declared above.
    questionHeading is a string made from the aggregate_Questions list
    districtHeading and questionHeading become the first line of the text
    the inputData is appended to each line after that
    """
    print "Writing aggregations to {0}".format(outputFile)
    districtHeading = 'Subdistrict_Assessed_location'
    questionHeading = ",".join(aggregate_Questions)
    # create the heading line for the line
    heading = "{0},{1},{2}".format(districtHeading,'coverage',questionHeading)
    with open(outputFile, 'wb') as outFile:
        writer = csv.writer((outFile), delimiter=' ', quotechar='|', quoting=csv.QUOTE_MINIMAL)
        # write the heading line to the file
        writer.writerow([heading])
        # write each SD line to the output file
        for row in inputData:
            writer.writerow([row])
    print "Aggregation file saved"

# ################################################

# return a list of SDs and their aggregated values
aggregation = aggregateToSubDistrict(fileIn, fileSheet)

# write the aggregation list to the outut file
writeToFile = writeResponsesToCSV(aggregation, fileOut)
