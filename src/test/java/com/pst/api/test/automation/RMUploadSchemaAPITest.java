package com.pst.api.test.automation;

import java.io.IOException;
import java.util.logging.Logger;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.Test;

public class RMUploadSchemaAPITest {

	Logger logger = Logger.getLogger("RMUploadSchemaAPITest");

	@Test
	public void testRMUpdateSchema() throws InvalidFormatException, IOException {

		String rmUpdateFilePath = "src/test/resources/RMSchemaUpload.xlsx";
		String rmUpdateFileResultedPath = "test-output/api-test-output/RMSchemaUploadResult.xlsx";
		String methodType = "upload";
		int statusColumnNumber = 6;
		RMSchemaTest.testRMSchema1(rmUpdateFilePath, rmUpdateFileResultedPath, methodType, statusColumnNumber);

	}

}
