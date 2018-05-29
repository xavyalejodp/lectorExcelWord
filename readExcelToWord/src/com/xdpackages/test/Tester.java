package com.xdpackages.test;

import java.io.IOException;

import org.junit.Test;

import com.xdpackages.poi.ExcelReader;
import com.xdpackages.poi.WordWriter;

public class Tester {
	
	@Test
	public void leerArchivoExcel() {
		System.out.println("tetst");
	
		ExcelReader excelReader = new ExcelReader();
		try {
			excelReader.readExcelWriteWord("C:\\Temp\\test.xlsx", "C:\\Temp\\test.docx");
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	
	@Test
	public void escribirArchivoWord() {
		WordWriter wordWriter= new WordWriter("");
		try {
			wordWriter.writeWord();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
