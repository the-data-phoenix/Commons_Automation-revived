import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import internal.GlobalVariable as GlobalVariable

/*This test script:
 - Opens the browser of choice: Chrome, Firefox or Edge
 - Driver opened by Katalon is used in Selenium.
 - Takes the Query from input excel and fetches data from Neo4j database.
   Saves the results from neo4j and application in the same name mentioned in the input excel.
 - Clicks on the Cases button in the Navbar of ICDC's homepage.
 - Clicks on the Filter 'Breed' from left pane
 - Selects the specific check box from 'Breed' filter.
 - Reads the results displayed for the selected filter (from all the pages in UI) and saves in the excel mentioned in Input file
 - Reads Neo4j DB using the query from Input file and saves the data in the excel mentioned in Input file
 - Reads Neo4j excel and Webdata excel as lists and compares the data.
 */
WebUI.closeBrowser()

CustomKeywords.'utilities.TestRunner.RunKatalon'('TC01_CCDI_phs003111_Sex-male_vitalsta-Dead_Sampletumorstat-Unknown.xlsx')
CustomKeywords.'utilities.TestRunner.clickTab'('Bento/Banner/Bento_Warning_Continue_Btn')

WebUI.waitForElementPresent(findTestObject('CCDI/Navbar/Explore_Menu'), 5)
CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/Navbar/Explore_Menu')

WebUI.waitForElementPresent(findTestObject('CCDI/ExplorePage/Filters/Study_Facet/Study_Facet'), 5)
CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/ExplorePage/Filters/Study_Facet/Study_Facet')

WebUI.waitForElementPresent(findTestObject('CCDI/ExplorePage/Filters/Study_Facet/PHS_Accession/PHS_Accession_Ddn'), 5)
CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/ExplorePage/Filters/Study_Facet/PHS_Accession/PHS_Accession_Ddn')

WebUI.waitForElementPresent(findTestObject('CCDI/ExplorePage/Filters/Study_Facet/PHS_Accession/phs003111_Chkbx'), 5)
CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/ExplorePage/Filters/Study_Facet/PHS_Accession/phs003111_Chkbx')

WebUI.waitForElementPresent(findTestObject('CCDI/ExplorePage/Filters/Demographics-Facet/SexAtBirth/Sex_Ddn'), 5)
CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/ExplorePage/Filters/Demographics-Facet/Sex/Sex_Ddn')

WebUI.waitForElementPresent(findTestObject('CCDI/ExplorePage/Filters/Demographics-Facet/SexAtBirth/Male_Chkbx'), 5)
CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/ExplorePage/Filters/Demographics-Facet/Sex/Male_Chkbx')

WebUI.waitForElementPresent(findTestObject('CCDI/ExplorePage/Filters/Diagnosis-Facet/Vital_Status/VitalStatus_Ddn'), 5)
CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/ExplorePage/Filters/Diagnosis-Facet/Vital_Status/VitalStatus_Ddn')

WebUI.waitForElementPresent(findTestObject('CCDI/ExplorePage/Filters/Diagnosis-Facet/Vital_Status/Dead_Chkbx'), 5)
CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/ExplorePage/Filters/Diagnosis-Facet/Vital_Status/Dead_Chkbx')

WebUI.waitForElementPresent(findTestObject('CCDI/ExplorePage/Filters/Samples_Facet/SampleTumorStatus/SampleTumorStatus_Ddn'), 5)
CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/ExplorePage/Filters/Samples_Facet/SampleTumorStatus/SampleTumorStatus_Ddn')

WebUI.waitForElementPresent(findTestObject('CCDI/ExplorePage/Filters/Samples_Facet/SampleTumorStatus/Unknown_Chkbx'), 5)
CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/ExplorePage/Filters/Samples_Facet/SampleTumorStatus/Unknown_Chkbx')


//Read Statbar
CustomKeywords.'utilities.TestRunner.readStatBarCCDIhub'('CCDI/Statbar/Studies_Cnt', 'CCDI/Statbar/Participants_Cnt', 
    'CCDI/Statbar/Samples_Cnt', 'CCDI/Statbar/Files_Cnt')

//Participants tab
WebUI.waitForElementPresent(findTestObject('CCDI/ExplorePage/Participants_ResultsTab'), 5)
CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/ExplorePage/Participants_ResultsTab')
CustomKeywords.'utilities.TestRunner.multiFunction'('CCDI', GlobalVariable.G_StatBar_Participants, 'CCDI/ExplorePage/CCDI_Participants_Tbl',
	'CCDI/ExplorePage/CCDI_Participants_TblHdr', 'CCDI/ExplorePage/CCDI_Participants_TblNextBtn', GlobalVariable.G_WebTabnameParticipants,	
	'TsvDataParticipants', GlobalVariable.G_QueryParticipantsTab)


//clicking the Diagnosis tab
WebUI.waitForElementPresent(findTestObject('CCDI/ExplorePage/Diagnosis_ResultsTab'), 5)
CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/ExplorePage/Diagnosis_ResultsTab')
CustomKeywords.'utilities.TestRunner.multiFunction'('CCDI', GlobalVariable.G_StatBar_Participants, 'CCDI/ExplorePage/CCDI_Diagnosis_Tbl',
	'CCDI/ExplorePage/CCDI_Diagnosis_TblHdr', 'CCDI/ExplorePage/CCDI_Diagnosis_TblNextBtn', GlobalVariable.G_WebTabnameDiagnosis,
	'TsvDataDiagnosis', GlobalVariable.G_QueryDiagnosisTab)


////clicking the Studies tab - Query needs to be fixed
//WebUI.waitForElementPresent(findTestObject('CCDI/ExplorePage/Studies_ResultsTab'), 5)
//CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/ExplorePage/Studies_ResultsTab')
//CustomKeywords.'utilities.TestRunner.multiFunction'('CCDI', GlobalVariable.G_StatBar_Participants, 'CCDI/ExplorePage/CCDI_Studies_Tbl',
//	'CCDI/ExplorePage/CCDI_Studies_TblHdr', 'CCDI/ExplorePage/CCDI_Studies_TblNextBtn', GlobalVariable.G_WebTabnameStudies,	
//	'TsvDataStudies', GlobalVariable.G_QueryStudiesTab)


//clicking the Samples tab
WebUI.waitForElementPresent(findTestObject('CCDI/ExplorePage/Samples_ResultsTab'), 5)
CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/ExplorePage/Samples_ResultsTab')
CustomKeywords.'utilities.TestRunner.multiFunction'('CCDI', GlobalVariable.G_StatBar_Samples, 'CCDI/ExplorePage/CCDI_Samples_Tbl',
	'CCDI/ExplorePage/CCDI_Samples_TblHdr', 'CCDI/ExplorePage/CCDI_Samples_TblNextBtn', GlobalVariable.G_WebTabnameSamples,	
	'TsvDataSamples', GlobalVariable.G_QuerySamplesTab)
 

////clicking the Files tab - Query needs to be fixed
//WebUI.waitForElementPresent(findTestObject('CCDI/ExplorePage/Files_ResultsTab'), 5)
//CustomKeywords.'utilities.TestRunner.clickTab'('CCDI/ExplorePage/Files_ResultsTab')
//CustomKeywords.'utilities.TestRunner.multiFunction'('CCDI', GlobalVariable.G_StatBar_Files, 'CCDI/ExplorePage/CCDI_Files_Tbl',
//'CCDI/ExplorePage/CCDI_Files_TblHdr', 'CCDI/ExplorePage/CCDI_Files_TblNextBtn', GlobalVariable.G_WebTabnameFiles,
//'TsvDataFiles', GlobalVariable.G_QueryFilesTab)
 
WebUI.closeBrowser()

