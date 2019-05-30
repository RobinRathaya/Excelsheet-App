package com.training.exportToExcel;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.training.exportToExcel.service.ExportAsWorkbook;

public class App {
	public static void main(String[] args) {
		final File folder = new File("E:\\developer_workspace\\exportToExcel\\TrainerFeedback");
		File[] listOfFiles = folder.listFiles();
		ExportAsWorkbook exportAsWorkbook = new ExportAsWorkbook();
		for (File file : listOfFiles) {
			if (file.isFile()) {
				try {
					ArrayList<String>column=exportAsWorkbook.loadAsExcel(file.getPath(), file.getName());
					exportAsWorkbook.ExportAsExcel(column);
				} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
}
