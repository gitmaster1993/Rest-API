package com.pst.api.test.automation;
import java.io.IOException;
import java.util.logging.Logger;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.Test;

public class GetRMSchemaAPITest {

	Logger logger = Logger.getLogger("GetRMSchemaAPITest");

	@Test
	public void testGetRMSchemaAPI() throws IOException, InvalidFormatException {

		String rmUpdateFilePath = "src/test/resources/GetRMSchema.xlsx";
		String rmUpdateFileResultedPath = "test-output/api-test-output/GetRMSchemaResult.xlsx";
		String methodType = "get";
		int statusColumnNumber = 5;
		RMSchemaTest.testRMSchema1(rmUpdateFilePath, rmUpdateFileResultedPath, methodType,statusColumnNumber);

	}
}
