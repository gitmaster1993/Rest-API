package com.pst.api.test.automation;
import java.io.IOException;
import java.util.logging.Logger;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.annotations.Test;

public class GetRoomMeasurementAPITest {

	Logger logger = Logger.getLogger("GetRoomMeasurementAPITest");

	@Test
	public void testGetRoomMeasurementAPI() throws IOException, InvalidFormatException {

		String rmUpdateFilePath = "src/test/resources/GetRoomMeasurement.xlsx";
		String rmUpdateFileResultedPath = "test-output/api-test-output/GetRoomMeasurementResult.xlsx";
		String methodType = "get";
		int statusColumnNumber = 5;
		RMSchemaTest.testRMSchema1(rmUpdateFilePath, rmUpdateFileResultedPath, methodType,statusColumnNumber);

	}
}
