package com.pst.api.test.automation;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.Properties;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFRow;

import io.restassured.response.Response;

public class TestBase {

	Logger logger = Logger.getLogger("TestBase");

	public static final String PSTAPIKEY = "PST_API_KEY";
	public static final String CONTENT_TYPE = "Content-Type";
	public static final String APPLICATION_JSON = "application/json";

	public static final String OP_EX_UPDATE = "OPExternalUpdate";
	public static final String KP_EX_UPDATE = "KPExternalUpdate";
	public static final String OP_EX_SUBMIT = "OPExternalSubmit";
	public static final String KP_EX_SUBMIT = "KPExternalSubmit";
	public static final String GET_RM = "GetRoomMeasurement";
	public static final String GET_RM_SCHEMA = "GetRMSchema";
	public static final String GET_LIST_RM_SCHEMA = "GetListRMSchema";
	public static final String URL_STRING = "urlString";
	public static final String RM_SCHEMA_UPLOAD = "RMSchemaUpload";
	public static final String RM_SCHEMA_UPDATE = "RMSchemaUpdate";
	
	public static final String UPLOAD = "upload";
	public static final String UPDATE = "update";
	public static final String GET = "get";

	private String pstValidationHost;
	private String officeApiKeyValue;
	private String kitchenApiKeyValue;
	private String rmBearerToken;
	private String rmApiKeyValue;
	private String rmHost;

	public String getPstValidationHost() {
		return pstValidationHost;
	}

	public void setPstValidationHost(String pstValidationHost) {
		this.pstValidationHost = pstValidationHost;
	}

	public String getOfficeApiKeyValue() {
		return officeApiKeyValue;
	}

	public void setOfficeApiKeyValue(String officeApiKeyValue) {
		this.officeApiKeyValue = officeApiKeyValue;
	}

	public String getKitchenApiKeyValue() {
		return kitchenApiKeyValue;
	}

	public void setKitchenApiKeyValue(String kitchenApiKeyValue) {
		this.kitchenApiKeyValue = kitchenApiKeyValue;
	}

	public String getRmBearerToken() {
		return rmBearerToken;
	}

	public void setRmBearerToken(String rmBearerToken) {
		this.rmBearerToken = rmBearerToken;
	}

	public String getRmApiKeyValue() {
		return rmApiKeyValue;
	}

	public void setRmApiKeyValue(String rmApiKeyValue) {
		this.rmApiKeyValue = rmApiKeyValue;
	}

	public String getRmHost() {
		return rmHost;
	}

	public void setRmHost(String rmHost) {
		this.rmHost = rmHost;
	}

	TestBase() throws IOException {
		ClassLoader classLoader = this.getClass().getClassLoader();
		File configFile = new File(classLoader.getResource("application.properties").getFile());
		try (FileInputStream inputStream = new FileInputStream(configFile);
				BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream));) {

			Properties prop = new Properties();
			FileReader fileReader = new FileReader(configFile);
			prop.load(fileReader);
			setPstValidationHost(prop.getProperty("PST_VALIDATION_HOST"));
			setOfficeApiKeyValue(prop.getProperty("OFFICE_APIKEY_VALUE"));
			setKitchenApiKeyValue(prop.getProperty("KITCHEN_APIKEY_VALUE"));
			setRmApiKeyValue(prop.getProperty("RM_APIKEY_VALUE"));
			setRmBearerToken(prop.getProperty("BEARER_TOKEN"));
			setRmHost(prop.getProperty("PST_RM_HOST"));

		} catch (FileNotFoundException e) {
			logger.info("Exception occure in TeatBase constructor");
		}
	}
	
	public static void passOrFailStatus(CellStyle style, CellStyle styleGreen, CellStyle styleRed, Response response,
			String respBody, XSSFRow row, String respValue, int statusFromExcel) {
		if (!respValue.isEmpty()) {
			Cell cell = row.createCell(9);
			cellStyleBoarder(style);
			cell.setCellStyle(style);

			if (respBody.contains(respValue) && statusFromExcel == response.statusCode()) {
				styleGreen.setFillBackgroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
				styleGreen.setFillPattern(FillPatternType.FINE_DOTS);
				cell.setCellStyle(styleGreen);
				cellStyleBoarder(styleGreen);
				cell.setCellValue(String.valueOf("PASS"));
			} else {
				styleRed.setFillBackgroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
				styleRed.setFillPattern(FillPatternType.FINE_DOTS);
				cell.setCellStyle(styleRed);
				cellStyleBoarder(styleRed);
				cell.setCellValue(String.valueOf("FAIL"));
			}

		} else {
			Cell cell = row.createCell(9);
			cellStyleBoarder(styleGreen);
			cell.setCellStyle(style);
			if (statusFromExcel == response.statusCode()) {
				styleGreen.setFillBackgroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
				styleGreen.setFillPattern(FillPatternType.FINE_DOTS);
				cell.setCellStyle(styleGreen);
				cell.setCellValue(String.valueOf("PASS"));
			} else {
				styleRed.setFillBackgroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
				styleRed.setFillPattern(FillPatternType.FINE_DOTS);
				cell.setCellStyle(styleRed);
				cell.setCellValue(String.valueOf("FAIL"));
			}
		}
	}
	
	public static void cellStyleBoarder(CellStyle style) {
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
	}

	public static void passOrFailStatus(CellStyle style, CellStyle styleGreen, CellStyle styleRed, Response response,
			String respBody, XSSFRow row, String respValue, int statusFromExcel, int passFailCellNum) {
		if (!respValue.isEmpty()) {
			Cell cell = row.createCell(passFailCellNum);
			cellStyleBoarder(style);
			cell.setCellStyle(style);

			if (respBody.contains(respValue) && statusFromExcel == response.statusCode()) {
				styleGreen.setFillBackgroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
				styleGreen.setFillPattern(FillPatternType.FINE_DOTS);
				cell.setCellStyle(styleGreen);
				cellStyleBoarder(styleGreen);
				cell.setCellValue(String.valueOf("PASS"));
			} else {
				styleRed.setFillBackgroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
				styleRed.setFillPattern(FillPatternType.FINE_DOTS);
				cell.setCellStyle(styleRed);
				cellStyleBoarder(styleRed);
				cell.setCellValue(String.valueOf("FAIL"));
			}

		} else {
			Cell cell = row.createCell(passFailCellNum);
			cellStyleBoarder(styleGreen);
			cell.setCellStyle(style);
			if (statusFromExcel == response.statusCode()) {
				styleGreen.setFillBackgroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
				styleGreen.setFillPattern(FillPatternType.FINE_DOTS);
				cell.setCellStyle(styleGreen);
				cell.setCellValue(String.valueOf("PASS"));
			} else {
				styleRed.setFillBackgroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
				styleRed.setFillPattern(FillPatternType.FINE_DOTS);
				cell.setCellStyle(styleRed);
				cell.setCellValue(String.valueOf("FAIL"));
			}
		}
	}
}