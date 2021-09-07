package com.pst.api.test.automation;

import static io.restassured.RestAssured.given;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import io.restassured.response.Response;
import io.restassured.specification.RequestSpecification;

public class RMSchemaTest {

	static Logger logger = Logger.getLogger("RMSchemaTest");

	private RMSchemaTest() {

	}

	private static void cellStyleBoarder(CellStyle style) {
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
	}

	@SuppressWarnings({ "rawtypes", "resource" })
	public static void testRMSchema1(String rmFilePath, String rmFileResultedPath, String methodType, int columnNumber)
			throws InvalidFormatException, IOException {

		URL urlString = null;
		TestBase testBase = new TestBase();
		List<Map> listOfRows = XLSXReader.fetchAllRowsFromRMSheet(rmFilePath);
		List<String> listOfResponseKeyValue = XLSXReader.fetchContainsColumnFromSheet(rmFilePath);
		List<String> listOfStatusColumn = XLSXReader.fetchStatusColumnFromSheet(rmFilePath);
		int successCount = 0;
		int failureCount = 0;
		int rowNum = 1;
		OPCPackage pkg = OPCPackage.open(new File(rmFilePath));
		XSSFWorkbook wb = new XSSFWorkbook(pkg);
		FileOutputStream out = new FileOutputStream(new File(rmFileResultedPath));
		String fileLocation = null;
		String status = null;
		for (Map map : listOfRows) {
			String formatedUrl = (String) map.get("urlString");

			if (methodType.equals(TestBase.UPLOAD)) {
				fileLocation = (String) map.get("filePath");
			} else if (methodType.equals(TestBase.UPDATE)) {
				status = (String) map.get("status");
			}

			try {
				urlString = new URL(testBase.getRmHost() + formatedUrl);
			} catch (MalformedURLException e) {
				e.printStackTrace();
			}

			XSSFSheet sheet = wb.getSheetAt(0);
			CellStyle style = wb.createCellStyle();
			CellStyle styleGreen = wb.createCellStyle();
			CellStyle styleRed = wb.createCellStyle();
			String requestBody = null;
			try {
				if (methodType.equals(TestBase.GET)) {
					logger.info("No request body");
				} else if (methodType.equals(TestBase.UPLOAD)) {
					requestBody = new String(Files.readAllBytes(Paths.get(fileLocation)));
				} else {
					requestBody = "{\"status\":\"" + status + "\"}";
				}
			} catch (IOException e) {
				requestBody = "";
			}
			Response response = null;

			if (methodType.equalsIgnoreCase(TestBase.GET)) {
				response = commonHeaders(testBase).when().get(urlString);
			} else if (methodType.equalsIgnoreCase(TestBase.UPLOAD)) {
				response = commonHeaders(testBase).body(requestBody).when().post(urlString);
			} else {
				response = commonHeaders(testBase).body(requestBody).when().put(urlString);
			}

			String respBody = response.getBody().asString();

			XSSFRow row = sheet.getRow(rowNum);
			String respValue = "";
			int statusFromExcel = 0;
			if (response.statusCode() == 200) {
				response.then().statusCode(200);
				setStatusInExcelColumn(methodType, columnNumber, style, response, row);
				statusFromExcel = Integer.parseInt(listOfStatusColumn.get(rowNum - 1));
				if (statusFromExcel == response.statusCode()) {
					successCount = successCount + 1;
				} else {
					failureCount = failureCount + 1;
				}
				respValue = listOfResponseKeyValue.get(rowNum - 1);
				rowNum = rowNum + 1;
			} else {
				setStatusInExcelColumn(methodType, columnNumber, style, response, row);
				statusFromExcel = Integer.parseInt(listOfStatusColumn.get(rowNum - 1));
				if (statusFromExcel == response.statusCode()) {
					successCount = successCount + 1;
				} else {
					failureCount = failureCount + 1;
				}
				respValue = listOfResponseKeyValue.get(rowNum - 1);
				rowNum = rowNum + 1;
			}

			TestBase.passOrFailStatus(style, styleGreen, styleRed, response, respBody, row, respValue, statusFromExcel,
					columnNumber + 1);

		}
		final int pass = successCount;
		final int fail = failureCount;
		logger.log(Level.INFO, () -> "RMSchema " + methodType + " External Submit API - Success count :" + pass
				+ ", Failure count :" + fail);
		wb.write(out);
		out.close();

	}

	private static void setStatusInExcelColumn(String methodType, int columnNumber, CellStyle style, Response response,
			XSSFRow row) {
		if (methodType.equalsIgnoreCase(TestBase.GET)) {
			createNewCell(columnNumber, style, response, row);
		}
		if (methodType.equalsIgnoreCase(TestBase.UPLOAD)) {
			createNewCell(columnNumber, style, response, row);
		}
		if (methodType.equalsIgnoreCase(TestBase.UPDATE)) {
			createNewCell(columnNumber, style, response, row);
		}
	}

	private static RequestSpecification commonHeaders(TestBase testBase) {
		return given().header(TestBase.CONTENT_TYPE, TestBase.APPLICATION_JSON)
				.header("Authorization", testBase.getRmBearerToken())
				.header(TestBase.PSTAPIKEY, testBase.getRmApiKeyValue());
	}

	private static void createNewCell(int columnNumber, CellStyle style, Response response, XSSFRow row) {
		Cell cell = row.createCell(columnNumber);
		cell.setCellValue(String.valueOf(response.statusCode()));
		cellStyleBoarder(style);
		cell.setCellStyle(style);
	}

}
