package com.pst.api.test.automation;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSXReader {

	static Logger logger = Logger.getLogger("XLSXReader");

	private XLSXReader() {

	}

	@SuppressWarnings("deprecation")
	public static String readDrawingCellValue(int vRow, int vColumn, String filePath) {
		String value = null;
		Workbook wb = null;
		try (FileInputStream fis = new FileInputStream(filePath)) {
			wb = new XSSFWorkbook(fis);
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(vRow);
			Cell cell = row.getCell(vColumn);

			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				value = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				value = Double.toString(cell.getNumericCellValue());
				break;
			default:
			}
		} catch (IOException e) {
			logger.log(Level.SEVERE, () -> "Exception while readDrawingCellValue from excel. " + e);
		}
		return value; // returns the cell value
	}

	static List<String> fetchContainsColumnFromSheet(String filePath) throws IOException {

		List<String> verifyValue = new ArrayList<>();
		XSSFWorkbook workbook = null;
		try (FileInputStream file = new FileInputStream(new File(filePath))) {
			List<String> list = null;

			workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			String sheetName = sheet.getSheetName();
			Iterator<Row> rowIterator = sheet.iterator();
			if (list == null) {
				list = new ArrayList<>();
			}
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

				if (row.getRowNum() != 0) {
					list = list == null ? new ArrayList<>() : list;
					Iterator<Cell> cellIterator = row.cellIterator();
					iterateCells(list, cellIterator);
					String responseContains = "";
					if (!list.isEmpty()) {
						switch (sheetName) {
						case TestBase.OP_EX_UPDATE:
							responseContains = list.get(4);
							list.clear();
							verifyValue.add(responseContains);
							break;
						case TestBase.KP_EX_UPDATE:
							rmSchemaStatus(verifyValue, list);
							break;
						case TestBase.OP_EX_SUBMIT:
							responseContains = list.get(7);
							list.clear();
							verifyValue.add(responseContains);
							break;
						case TestBase.KP_EX_SUBMIT:
							responseContains = list.get(5);
							list.clear();
							verifyValue.add(responseContains);
							break;
						case TestBase.RM_SCHEMA_UPLOAD:
							getValueFromIndexTwo(verifyValue, list);
							break;
						case TestBase.RM_SCHEMA_UPDATE:
							getValueFromIndexTwo(verifyValue, list);
							break;
						case TestBase.GET_RM:
							getValueFromIndexOne(verifyValue, list);
							break;
						case TestBase.GET_RM_SCHEMA:
							getValueFromIndexOne(verifyValue, list);
							break;
						case TestBase.GET_LIST_RM_SCHEMA:
							getValueFromIndexOne(verifyValue, list);
							break;
						default:
							list.clear();
							verifyValue.add(responseContains);
						}
					}
				}
			}
		} catch (Exception e) {
			logger.log(Level.SEVERE, () -> "Exception while fetchContainsColumnFromSheet from excel. " + e);
		} finally {
			workbook.close();
		}
		return verifyValue;

	}

	private static void getValueFromIndexTwo(List<String> verifyValue, List<String> list) {
		String responseContains;
		responseContains = list.get(2);
		list.clear();
		verifyValue.add(responseContains);
	}

	private static void getValueFromIndexOne(List<String> verifyValue, List<String> list) {
		String responseContains;
		responseContains = list.get(1);
		list.clear();
		verifyValue.add(responseContains);
	}

	static List<String> fetchAllRowsFromSheet(String filePath) throws IOException {
		List<String> jsonBody = new ArrayList<>();
		XSSFWorkbook workbook = null;
		try (FileInputStream file = new FileInputStream(new File(filePath))) {
			List<String> list = null;
			String officeUpdateRequestBody = "{\"status\":\"%1$s\",\"retailUnit\":\"%2$s\",\"coworkerId\":\"%3$s\"}";
			String kitchenUpdateRequestBody = "{\"Retail_Unit\":\"%2$s\",\"status\":\"%1$s\"}";
			String officeSubmitRequestBody = "{\"refCode\":\"%1$s\",\"name\":\"%2$s\",\"createdBy\":\"%3$s\",\"idp\":\"%4$s\",\"retailUnit\":\"%5$s\",\"languageCode\":\"%6$s\",\"comments\":\"%7$s\"}";
			String kitchenSubmitRequestBody = "{\"Project_ID\":\"%1$s\",\"Retail_Unit\":\"%2$s\",\"Language\":\"%3$s\",\"User_Comments\":\"%4$s\"}";

			workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			String sheetName = sheet.getSheetName();
			Iterator<Row> rowIterator = sheet.iterator();
			if (list == null) {
				list = new ArrayList<>();
			}
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

				if (row.getRowNum() != 0) {
					list = list == null ? new ArrayList<>() : list;
					Iterator<Cell> cellIterator = row.cellIterator();
					iterateCells(list, cellIterator);
					String jsonRequstBody = "";
					if (!list.isEmpty()) {
						switch (sheetName) {
						case TestBase.OP_EX_UPDATE:
							jsonRequstBody = String.format(officeUpdateRequestBody, list.get(0), list.get(1),
									list.get(2));
							list.clear();
							jsonBody.add(jsonRequstBody);
							break;
						case TestBase.KP_EX_UPDATE:
							jsonRequstBody = String.format(kitchenUpdateRequestBody, list.get(0), list.get(1));
							list.clear();
							jsonBody.add(jsonRequstBody);
							break;
						case TestBase.OP_EX_SUBMIT:
							jsonRequstBody = String.format(officeSubmitRequestBody, list.get(0), list.get(1),
									list.get(2), list.get(3), list.get(4), list.get(5), list.get(6));
							list.clear();
							jsonBody.add(jsonRequstBody);
							break;
						case TestBase.KP_EX_SUBMIT:
							jsonRequstBody = String.format(kitchenSubmitRequestBody, list.get(0), list.get(1),
									list.get(2), list.get(3));
							list.clear();
							jsonBody.add(jsonRequstBody);
							break;
						default:
							list.clear();
							jsonBody.add(jsonRequstBody);
						}
					}
				}
			}
		} catch (Exception e) {
			logger.log(Level.SEVERE, () -> "Exception while fetchAllRowsFromSheet from excel. " + e);
		} finally {
			workbook.close();
		}
		return jsonBody;
	}

	@SuppressWarnings("deprecation")
	private static void iterateCells(List<String> list, Iterator<Cell> cellIterator) {
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			if (cell.getRowIndex() != 0 && cell.getColumnIndex() >= 2) {
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					int xs = (int) cell.getNumericCellValue();
					String sd = Integer.toString(xs);
					list.add(sd);
					break;
				case Cell.CELL_TYPE_STRING:
					list.add(cell.getStringCellValue());
					break;
				case Cell.CELL_TYPE_BLANK:
					list.add("");
					break;
				default:
				}
			}
		}
	}

	public static List<String> fetchStatusColumnFromSheet(String filePath) {
		List<String> actualStatus = new ArrayList<>();
		List<String> list = null;
		XSSFWorkbook workbook = null;
		try (FileInputStream file = new FileInputStream(new File(filePath))) {

			workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			String sheetName = sheet.getSheetName();
			Iterator<Row> rowIterator = sheet.iterator();

			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

				if (row.getRowNum() != 0) {
					list = list == null ? new ArrayList<>() : list;
					Iterator<Cell> cellIterator = row.cellIterator();

					iterateCells(list, cellIterator);
					String statusCode = "";
					if (!list.isEmpty()) {
						switch (sheetName) {
						case TestBase.OP_EX_UPDATE:
							kPOPStatus(actualStatus, list);
							break;
						case TestBase.KP_EX_UPDATE:
							statusCode = list.get(4);
							list.clear();
							actualStatus.add(statusCode);
							break;
						case TestBase.OP_EX_SUBMIT:
							statusCode = list.get(8);
							list.clear();
							actualStatus.add(statusCode);
							break;
						case TestBase.KP_EX_SUBMIT:
							kPOPStatus(actualStatus, list);
							break;
						case TestBase.RM_SCHEMA_UPLOAD:
							rmSchemaStatus(actualStatus, list);
							break;
						case TestBase.RM_SCHEMA_UPDATE:
							rmSchemaStatus(actualStatus, list);
							break;
						case TestBase.GET_RM:
							getStatusFromIndexTw0(actualStatus, list);
							break;
						case TestBase.GET_RM_SCHEMA:
							getStatusFromIndexTw0(actualStatus, list);
							break;
						case TestBase.GET_LIST_RM_SCHEMA:
							getStatusFromIndexTw0(actualStatus, list);
							break;
						default:
							list.clear();
							actualStatus.add(statusCode);
						}
					}
				}
			}

		} catch (Exception e) {
			logger.log(Level.SEVERE, () -> "Exception while fetch status from excel. " + e);
		} finally {
			try {
				workbook.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return actualStatus;
	}

	private static void rmSchemaStatus(List<String> actualStatus, List<String> list) {
		String statusCode;
		statusCode = list.get(3);
		list.clear();
		actualStatus.add(statusCode);
	}

	private static void getStatusFromIndexTw0(List<String> actualStatus, List<String> list) {
		getValueFromIndexTwo(actualStatus, list);
	}

	private static void kPOPStatus(List<String> actualStatus, List<String> list) {
		String statusCode;
		statusCode = list.get(5);
		list.clear();
		actualStatus.add(statusCode);
	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static List<Map> fetchAllRowsFromRMSheet(String rmFilePath) {
		List<Map> jsonBody = new ArrayList<>();
		List<String> list = null;
		XSSFWorkbook workbook = null;
		try (FileInputStream file = new FileInputStream(new File(rmFilePath))) {
			String rmSchemaFilePath = "src/test/resources/%1$s";
			workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			String sheetName = sheet.getSheetName();
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

				if (row.getRowNum() != 0) {
					list = (list == null) ? new ArrayList<>() : list;
					Iterator<Cell> cellIterator = row.cellIterator();
					iterateCells(list, cellIterator);
					String pathValue = "";
					String urlValue = "";
					Map map = new HashMap();
					if (!list.isEmpty()) {
						switch (sheetName) {
						case "RMSchemaUpload":
							pathValue = String.format(rmSchemaFilePath, list.get(0));
							urlValue = list.get(1);
							map.put("filePath", pathValue);
							map.put(TestBase.URL_STRING, urlValue);
							list.clear();
							jsonBody.add(map);
							break;
						case "RMSchemaUpdate":
							String status = list.get(0);
							urlValue = list.get(1);
							map.put("status", status);
							map.put(TestBase.URL_STRING, urlValue);
							list.clear();
							jsonBody.add(map);
							break;
						case "GetRoomMeasurement":
							getValueFromIndexZero(jsonBody, list, map);
							break;
						case "GetRMSchema":
							getValueFromIndexZero(jsonBody, list, map);
							break;
						case "GetListRMSchema":
							getValueFromIndexZero(jsonBody, list, map);
							break;
						default:
							list.clear();
							jsonBody.add(map);
						}
					}
				}
			}
		} catch (Exception e) {
			logger.log(Level.SEVERE, () -> "Exception while fetchAllRowsFromSheet from excel. " + e);
		}
		return jsonBody;

	}

	@SuppressWarnings({ "rawtypes", "unchecked" })
	private static void getValueFromIndexZero(List<Map> jsonBody, List<String> list, Map map) {
		String urlValue;
		urlValue = list.get(0);
		map.put("urlString", urlValue);
		list.clear();
		jsonBody.add(map);
	}

}