package utilities
import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.checkpoint.CheckpointFactory
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testcase.TestCaseFactory
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testdata.TestDataFactory
import com.kms.katalon.core.testobject.ObjectRepository
import com.kms.katalon.core.testobject.TestObject

import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords

import internal.GlobalVariable

import org.openqa.selenium.WebElement
import org.openqa.selenium.WebDriver
import org.openqa.selenium.By

import com.kms.katalon.core.mobile.keyword.internal.MobileDriverFactory
import com.kms.katalon.core.webui.driver.DriverFactory

import com.kms.katalon.core.testobject.RequestObject
import com.kms.katalon.core.testobject.ResponseObject
import com.kms.katalon.core.testobject.ConditionType
import com.kms.katalon.core.testobject.TestObjectProperty

import com.kms.katalon.core.mobile.helper.MobileElementCommonHelper
import com.kms.katalon.core.util.KeywordUtil

import com.kms.katalon.core.webui.exception.WebElementNotFoundException

import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS

import groovy.json.JsonSlurper

import java.io.File
import com.kms.katalon.core.configuration.RunConfiguration
class API_Functions {

	// Function to send API request and return response


	static ResponseObject sendRequestAndCaptureResponse(String apiObjectPath) {
		// Retrieve the project directory dynamically
		String projectDir = RunConfiguration.getProjectDir()

		// Define the output directory relative to the project directory
		String outputDirectory = projectDir + "/OutputFiles/"

		// Get the test case name from Global Variable (set by the listener)
		String testCaseName = GlobalVariable.currentTestCaseName
		if (testCaseName == null || testCaseName.isEmpty()) {
			// Fallback to the execution source name if test case name is not set
			testCaseName = RunConfiguration.getExecutionSourceName()
		}
		println("This is the test case name: " + testCaseName)
		
// Load the API request object
    RequestObject request = findTestObject(apiObjectPath)

    // Set timeouts using global configuration
    request.setHttpHeaderProperties([
        new TestObjectProperty('Connection-Timeout', com.kms.katalon.core.testobject.ConditionType.EQUALS, '10000'),
        new TestObjectProperty('Socket-Timeout', com.kms.katalon.core.testobject.ConditionType.EQUALS, '10000')
    ])
		// Send the API request
		ResponseObject response = WS.sendRequest(findTestObject(apiObjectPath))
		println("Response Status: " + response.getStatusCode())
		println("Response Body: " + response.getResponseBodyContent())
/*
		// Extract the API name to use as part of the file name
		String apiName = apiObjectPath.split('/').last().replaceAll(' ', '_')

		// Combine test case name and API name for the file name
		File responseFile = new File(outputDirectory + testCaseName + "_" + apiName + '.txt')

		// Write the response to the file
		responseFile.write("Response Status: " + response.getStatusCode() + "\n")
		responseFile << "Response Body: " + response.getResponseBodyContent() + "\n"
*/
		return response
	}



	// Function to parse JSON response
	static def parseResponse(ResponseObject response) {
		JsonSlurper jsonSlurper = new JsonSlurper()
		return jsonSlurper.parseText(response.getResponseBodyContent())
	}

	// Function to find the keyed entry in the response
	static def findEntry(def responseData, def key) {
		switch (key) {
			case 'KidsFirst':
				def kidsFirstEntry = responseData.find { it.source == 'KidsFirst' }
				if (kidsFirstEntry == null) {
					KeywordUtil.markFailedAndStop("No entry with 'source: KidsFirst' found in the response.")
				}
				return kidsFirstEntry

			case 'StJude':
				def stJudeEntry = responseData.find { it.source == 'StJude' }
				if (stJudeEntry == null) {
					KeywordUtil.markFailedAndStop("No entry with 'source: StJude' found in the response.")
				}
				return stJudeEntry

			case 'PCDC':
				def pcdcEntry = responseData.find { it.source == 'PCDC' }
				if (pcdcEntry == null) {
					KeywordUtil.markFailedAndStop("No entry with 'source: PCDC' found in the response.")
				}
				return pcdcEntry

			case 'Treehouse':
				def treehouseEntry = responseData.find { it.source == 'Treehouse' }
				if (treehouseEntry == null) {
					KeywordUtil.markFailedAndStop("No entry with 'source: Treehouse' found in the response.")
				}
				return treehouseEntry
			default:
				KeywordUtil.markFailedAndStop("No entry found for the provided key: ${key}")
		}
	}

	// Function to compare API responses
	static void compareAPIResponses(def singleNodeData, def AlextractedEntry, def key) {
		boolean isContained = true

		// Compare 'total' and 'missing' properties
		if (AlextractedEntry.total != singleNodeData.total) {
			isContained = false
			KeywordUtil.markWarning("Mismatch in 'total' property. Expected: ${singleNodeData.total}, Found: ${AlextractedEntry.total}")
		}
		if (AlextractedEntry.missing != singleNodeData.missing) {
			isContained = false
			KeywordUtil.markWarning("Mismatch in 'missing' property. Expected: ${singleNodeData.missing}, Found: ${AlextractedEntry.missing}")
		}

		// Compare 'values' array
		singleNodeData.values.each { singleNodeValue ->
			def matchingValue = AlextractedEntry.values.find { it.value == singleNodeValue.value }

			if (matchingValue == null) {
				isContained = false
				KeywordUtil.markWarning("Value '${singleNodeValue.value}' from single node data not found in ${key} entry.")
			} else if (matchingValue.count != singleNodeValue.count) {
				isContained = false
				KeywordUtil.markWarning("Count mismatch for value '${singleNodeValue.value}'. Expected: ${singleNodeValue.count}, Found: ${matchingValue.count}")
			}
		}

		assert isContained : "The ${key} entry does not match the singleNode data."
	}




	//********************************************************************************
	/**
	 * Send request and verify status code
	 * @param request request object, must be an instance of RequestObject
	 * @param expectedStatusCode
	 * @return a boolean to indicate whether the response status code equals the expected one
	 */
	@Keyword
	def verifyStatusCode(TestObject request, int expectedStatusCode) {
		if (request instanceof RequestObject) {
			RequestObject requestObject = (RequestObject) request
			ResponseObject response = WSBuiltInKeywords.sendRequest(requestObject)
			if (response.getStatusCode() == expectedStatusCode) {
				KeywordUtil.markPassed("Response status codes match")
			} else {
				KeywordUtil.markFailed("Response status code not match. Expected: " +
						expectedStatusCode + " - Actual: " + response.getStatusCode() )
			}
		} else {
			KeywordUtil.markFailed(request.getObjectId() + " is not a RequestObject")
		}
	}

	/**
	 * Add Header basic authorization field,
	 * this field value is Base64 encoded token from user name and password
	 * @param request object, must be an instance of RequestObject
	 * @param username username
	 * @param password password
	 * @return the original request object with basic authorization header field added
	 */
	@Keyword
	def addBasicAuthorizationProperty(TestObject request, String username, String password) {
		if (request instanceof RequestObject) {
			String authorizationValue = username + ":" + password
			authorizationValue = "Basic " + authorizationValue.bytes.encodeBase64().toString()

			// Find available basic authorization field and change its value to the new one, if any
			List<TestObjectProperty> headerProperties = request.getHttpHeaderProperties()
			boolean fieldExist = false
			for (int i = 0; i < headerProperties.size(); i++) {
				TestObjectProperty headerField = headerProperties.get(i)
				if (headerField.getName().equals('Authorization')) {
					KeywordUtil.logInfo("Found existent basic authorization field. Replacing its value.")
					headerField.setValue(authorizationValue)
					fieldExist = true
					break
				}
			}

			if (!fieldExist) {
				TestObjectProperty authorizationProperty = new TestObjectProperty("Authorization",
						ConditionType.EQUALS, authorizationValue, true)
				headerProperties.add(authorizationProperty)
			}
			KeywordUtil.markPassed("Basic authorization field has been added to request header")
		} else {
			KeywordUtil.markFailed(request.getObjectId() + "is not a RequestObject")
		}
		return request
	}
}