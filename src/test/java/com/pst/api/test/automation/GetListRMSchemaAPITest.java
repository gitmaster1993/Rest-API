package com.pst.api.test.automation;
import java.io.IOException;
import java.util.logging.Logger;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.Test;

public class GetListRMSchemaAPITest {

	Logger logger = Logger.getLogger("GetListRMSchemaAPITest");

	@Test
	public void testGetListRMSchemaAPI() throws IOException, InvalidFormatException {

		String rmUpdateFilePath = "src/test/resources/GetListRMSchema.xlsx";
		String rmUpdateFileResultedPath = "test-output/api-test-output/GetListRMSchemaResult.xlsx";
		String methodType = "get";
		int statusColumnNumber = 5;
		RMSchemaTest.testRMSchema1(rmUpdateFilePath, rmUpdateFileResultedPath, methodType,statusColumnNumber);

	}

}
