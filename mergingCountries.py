import os
import pandas as pd
import xlrd
pd.core.format.header_style = None

filePath = 'C:\\REACH\\SYR\\Projects\\13BVJ_AoO\\Activities\\Round7_November2015\\Data Collection\\AoO_Round7_Data\\AoO_Round7_datacleaning'
fileCountry = '{0}_KI_Village_Level_Monitoring_Tool_Nov_Final_coded_NA.xlsx'
outFile = 'AoO_Round7_Nov2015_Dataset_Merged_NA.xlsx'

def mergeCountriesFromXlsx(filePath, fileCountry, outFile):
	appendList = []
	try:
		dfJOR = pd.DataFrame()
		countryCode = 'JOR'
		fileJOR = os.path.join(filePath, fileCountry.format(countryCode))
		dataJOR = pd.read_excel(fileJOR, '{0}_KI_Village_Level_Monitoring'.format(countryCode))
		dfJOR = dfJOR.append(dataJOR)
		#appendList.append(dfJOR)
		print "Jordan loaded"
	except Exception, e:
		print "Jordan not loaded"

	try:	
		dfTUR = pd.DataFrame()
		countryCode = 'TUR'
		fileTUR = os.path.join(filePath, fileCountry.format(countryCode))
		dfTUR = pd.DataFrame()
		dataTUR = pd.read_excel(fileTUR, '{0}_KI_Village_Level_Monitoring'.format(countryCode))
		dfTUR = dfTUR.append(dataTUR)
		appendList.append(dfTUR)
		print "Turkey loaded"
	except Exception, e:
		print "Turkey not loaded"

	try:
		dfIRQ = pd.DataFrame()
		countryCode = 'IRQ'
		fileIRQ = os.path.join(filePath, fileCountry.format(countryCode))
		dfIRQ = pd.DataFrame()
		dataIRQ = pd.read_excel(fileIRQ, '{0}_KI_Village_Level_Monitoring'.format(countryCode))
		dfIRQ = dfIRQ.append(dataIRQ)
		appendList.append(dfIRQ)
		print "Iraq loaded"
	except Exception, e:
		print "Iraq not loaded"

	# try:
	# 	dfLBN = pd.DataFrame()
	# 	countryCode = 'LBN'
	# 	fileLBN = os.path.join(filePath, '{0}_KI_Village_Level_Monitoring_Tool_Nov_Final_output_Test.xlsx'.format(countryCode))
	# 	dfLBN = pd.DataFrame()
	# 	dataLBN = pd.read_excel(fileLBN, '{0}_KI_Village_Level_Monitoring'.format(countryCode))
	# 	dfLBN = dfLBN.append(dataLBN)
	# 	appendList.append(dfLBN)
	# 	print "Lebanon loaded"
	# except Exception, e:
	# 	print "Lebanon not loaded"

	try:
	    print "Appending countries..."
	    mergedCountries = dfJOR.append(appendList)
	    print "Countries appended"
	    mergedFile = os.path.join(filePath, outFile)
	    writer = pd.ExcelWriter(mergedFile)
	    mergedCountries.to_excel(writer, 'ALL_KI_Village_Level_Monitoring', index=False)
	    #mergedCountries.to_excel(writer, 'ALL_KI_Village_Level_Monitoring', index=False)
	    writer.save()
	    print "Merged file saved"

	except ValueError, e:
	        print "Error: {0}".format(e)

# ################################################################


mergeCountries = mergeCountriesFromXlsx(filePath, fileCountry, outFile)