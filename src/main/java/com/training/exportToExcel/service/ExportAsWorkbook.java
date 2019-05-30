package com.training.exportToExcel.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExportAsWorkbook {
	static int lastActiveCell = 2;

	public ArrayList<String> loadAsExcel(String excelFilePath, String fileName)
			throws IOException, EncryptedDocumentException, InvalidFormatException {
		System.out.println(excelFilePath);
		ArrayList<String> answersList = new ArrayList<>();
		try {
			FileInputStream file = new FileInputStream(new File(excelFilePath));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			answersList.add(fileName.replace(".xlsx", ""));
			for (Row row : sheet) {
				Cell cell = row.getCell(2);
				if (cell != null) {
					switch (cell.getCellTypeEnum()) {
					case NUMERIC:
						answersList.add(String.valueOf(cell.getNumericCellValue()));
						break;
					case STRING:
						answersList.add(cell.getStringCellValue());
						break;
					case BLANK:
						answersList.add("");
						break;
					default:
						break;
					}
				} else {
					continue;
				}
			}
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		System.out.println(answersList);
		return answersList;
	}

	public void ExportAsExcel(ArrayList<String> column) throws FileNotFoundException {
		final String filePath = "E:\\developer_workspace\\exportToExcel\\final\\TrainerFeedback.xlsx";
		try {
			FileInputStream file = new FileInputStream(new File(filePath));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			sheet.autoSizeColumn(lastActiveCell);
			int rownum = 0;
			sheet.autoSizeColumn(lastActiveCell);
			for (String content : column) {
				if (sheet.getRow(rownum) != null) {
					Cell cell = sheet.getRow(rownum).createCell(lastActiveCell);
					cell.setCellValue(content);
				} else {
					Cell cell = sheet.createRow(rownum).createCell(lastActiveCell);
					cell.setCellValue(content);
				}
				rownum++;
			}
			lastActiveCell++;
			FileOutputStream fileOutputStream = new FileOutputStream(filePath);
			workbook.write(fileOutputStream);
			fileOutputStream.close();
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
