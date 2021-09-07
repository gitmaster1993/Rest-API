package com.pst.api.test.automation;

import java.io.IOException;
import java.util.logging.Logger;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.Test;

public class RMUpdateSchemaAPITest {

	Logger logger = Logger.getLogger("RMUpdateSchemaAPITest");

	@Test
	public void testRMUpdateSchema() throws InvalidFormatException, IOException {

		String rmUpdateFilePath = "src/test/resources/RMSchemaUpdate.xlsx";
		String rmUpdateFileResultedPath = "test-output/api-test-output/RMSchemaUpdateResult.xlsx";
		String methodType = "update";
		int statusColumnNumber = 6;
		RMSchemaTest.testRMSchema1(rmUpdateFilePath, rmUpdateFileResultedPath, methodType, statusColumnNumber);

	}

}
