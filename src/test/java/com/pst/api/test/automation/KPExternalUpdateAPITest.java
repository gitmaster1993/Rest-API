package com.pst.api.test.automation;

import static io.restassured.RestAssured.given;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import io.restassured.response.Response;

public class KPExternalUpdateAPITest {

	Logger logger = Logger.getLogger("KPExternalUpdateAPITest");

	@SuppressWarnings("resource")
	@Test
	public void testExternalUpdateAPI() throws IOException, InvalidFormatException {

		URL urlString = null;
		TestBase testBase = new TestBase();
		String filePath = "src/test/resources/KPExternalUpdate.xlsx";
		List<String> listOfRows = XLSXReader.fetchAllRowsFromSheet(filePath);
		List<String> listOfResponseKeyValue = XLSXReader.fetchContainsColumnFromSheet(filePath);

		List<String> listOfStatusColumn = XLSXReader.fetchStatusColumnFromSheet(filePath);
		String drawingId = XLSXReader.readDrawingCellValue(1, 4, filePath);
		try {
			urlString = new URL(
					testBase.getPstValidationHost() + "/kitchen/" + drawingId + "?includeCustomerEmail=true");
		} catch (MalformedURLException e) {
			e.printStackTrace();
		}
		int successCount = 0;
		int failureCount = 0;
		int rowNum = 1;
		OPCPackage pkg = OPCPackage.open(new File(filePath));
		XSSFWorkbook wb = new XSSFWorkbook(pkg);
		XSSFSheet sheet = wb.getSheetAt(0);
		CellStyle style = wb.createCellStyle();
		CellStyle styleGreen = wb.createCellStyle();
		CellStyle styleRed = wb.createCellStyle();
		int statusCellNum = 7;
		int passFailCellNum = 8;
		FileOutputStream out = new FileOutputStream(
				new File("test-output/api-test-output/KPExternalUpdateResult.xlsx"));

		for (String requestBody : listOfRows) {
			final Response response = given().header(TestBase.CONTENT_TYPE, TestBase.APPLICATION_JSON)
					.header(TestBase.PSTAPIKEY, testBase.getKitchenApiKeyValue()).body(requestBody).when()
					.put(urlString);
			String respBody = response.getBody().asString();
			XSSFRow row = sheet.getRow(rowNum);
			String respValue = "";
			int statusFromExcel = 0;

			Cell cell = row.createCell(statusCellNum);
			cell.setCellValue(String.valueOf(response.statusCode()));
			TestBase.cellStyleBoarder(style);
			cell.setCellStyle(style);

			if (response.statusCode() == 200) {
				response.then().statusCode(200);
				statusFromExcel = Integer.parseInt(listOfStatusColumn.get(rowNum - 1));
				if (statusFromExcel == response.statusCode()) {
					successCount = successCount + 1;
				} else {
					failureCount = failureCount + 1;
				}
				respValue = listOfResponseKeyValue.get(rowNum - 1);
				rowNum = rowNum + 1;
			} else {
				statusFromExcel = Integer.parseInt(listOfStatusColumn.get(rowNum - 1));
				if (statusFromExcel == response.statusCode()) {
					successCount = successCount + 1;
				} else {
					failureCount = failureCount + 1;
				}
				respValue = listOfResponseKeyValue.get(rowNum - 1);
				rowNum = rowNum + 1;
			}

			TestBase.passOrFailStatus(style, styleGreen, styleRed, response, respBody, row, respValue, statusFromExcel,passFailCellNum);

		}
		final int pass = successCount;
		final int fail = failureCount;
		logger.log(Level.INFO,
				() -> "Kitchen External Update API - Success count :" + pass + ", Failure count :" + fail);
		wb.write(out);
		out.close();
	}

}
