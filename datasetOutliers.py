import os
import pandas as pd
import xlrd
import csv
import numpy as np
##pd.core.format.header_style = None

# update these file paths every month
filePath = "C:\\REACH\\SYR\\Projects\\13BVJ_AoO\\Activities\\Round7_November2015\\Data Collection\\AoO_Round7_Data\\AoO_Round7_datacleaning"
fileIn = "AoO_Round7_Nov2015_Dataset_Merged_v5_outliers.xlsx"

# csv file containing columns names that have numeric responses
# update this each month in accordance with the master indicator list
# eventually the aim would be to have the script automatically detect numeric responses
numericCols = "AoO_Round7_numericColumns.csv"

# dicts to receive outliers for each country
jorDict = {}
turDict = {}
lbnDict = {}
irqDict = {}

# --------------------------------

class countryOutliers:
    def __init__(self, area, outliers):
        self.area = area
        self.outliers = outliers

    def writeToFile(self):
        with open(os.path.join(filePath, '{0}Outliers.txt'.format(self.area)), 'w') as outlierFile:
            outlierFile.writelines("{0} outliers\n\n".format(self.area))
            for key, values in sorted(self.outliers.iteritems()):
                outlierFile.writelines("{0}\n".format(key))
                for value in values:
                    outlierFile.writelines("{0}\n".format(value))
                outlierFile.writelines("\n")
        print "Outliers for {0} created".format(self.area)

# #############################################

def csvFiletoList(filePath, fileName):
    rowdata = []
    with open(os.path.join(filePath, fileName), 'r') as colFile:     
        reader = csv.reader(colFile)
        for row in reader:
            rowdata.append((row))
    return rowdata

# #############################################

def detectOutliers(numericQs):
    try:
        print "Trying to load dataset..."
        dfIn = pd.DataFrame()
        dataIn = pd.read_excel((os.path.join(filePath, fileIn)), 'ALL_KI_Village_Level_Monitoring')
        dfIn = dfIn.append(dataIn)
        print "Dataset loaded"
        colList = []                # empty array to receive column headings from the loaded questions
        rows = dfIn.shape[0]        # get the number of rows in the dataset
        for nQ in numericQs:        # appending numeric questions to colList array
            colList.append(nQ)
        print "Iterating through {0} columns and isolating numeric columns...".format(len(colList))
        for cL in colList:
            col = cL[0]
            stats = dfIn[col].describe(percentiles=[.25, .75, .98])
            outliers = 0
            #print stats
            outlier98 = stats.loc['98%']    # returns numpy.float64
            outlierMax = stats.loc['max']   # returns numpy.float64
            q1 = stats.loc['25%']           # returns numpy.float64
            q3 = stats.loc['75%']           # returns numpy.float64
            iQRange = q3-q1
            iQRange1_5 = int(iQRange)*1.5
            maxFence = q3+iQRange1_5        # outer minor fence
            minFence = q1-iQRange1_5        # inner minor fence

            for y in range(rows):
                cellValue = dfIn.iloc[y][col]
                #if cellValue <= outlierMax and cellValue >= outlier98:
                if cellValue <= outlierMax and cellValue >= maxFence:    
                    value = "Participant ID {0} is an outlier: {1}".format(dfIn.iloc[y]['Basic/Participant_no'], cellValue) 
                    if cellValue <= outlierMax and cellValue >= outlier98:
                        if dfIn.iloc[y]['Country'] == "JOR":  
                            if col not in jorDict:
                                jorDict[col] = [value]
                            else:
                                jorDict[col].append(value)
                        elif dfIn.iloc[y]['Country'] == "TUR":
                            if col not in turDict:
                                turDict[col] = [value]
                            else:
                                turDict[col].append(value)
                        elif dfIn.iloc[y]['Country'] == "LBN":
                            if col not in lbnDict:
                                lbnDict[col] = [value]
                            else:
                                lbnDict[col].append(value)
                        elif dfIn.iloc[y]['Country'] == "IRQ":
                            if col not in irqDict:
                                irqDict[col] = [value]
                            else:
                                irqDict[col].append(value)
                        else:
                            #pass
                            print "Did not process value {0} for col {1}".format(col, value)   

        print "Outliers determined - writing to respective country files"
        Jordan = countryOutliers("Jordan", jorDict)
        Jordan.writeToFile()

        Turkey = countryOutliers("Turkey", turDict)
        Turkey.writeToFile()

        Lebanon = countryOutliers("Lebanon", lbnDict)
        Lebanon.writeToFile()

        Iraq = countryOutliers("Iraq", irqDict)
        Iraq.writeToFile()  

        print "Outlier processing complete"
    except Exception, e:
        print "Error: {0}".format(e)

# #############################################

# reads numeric columns from a csv file
numericQuestions = csvFiletoList(filePath, numericCols)

# generates outliers for each column and writes them to a file
runOutliers = detectOutliers(numericQuestions)