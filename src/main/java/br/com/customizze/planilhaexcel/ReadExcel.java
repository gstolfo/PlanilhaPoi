package br.com.customizze.planilhaexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel extends HttpServlet {
	
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	@Override
	protected void doGet(HttpServletRequest req, HttpServletResponse resp) throws ServletException, IOException {
	
		String filePath = "Planilha_Origem.xlsm";
		
		try {
			// Abrindo o arquivo e recuperando a planilha			
			URL url = getClass().getResource(filePath);
			File f = new File(url.getPath());
			FileInputStream file = new FileInputStream(f);
			
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			
			
			//Dados da Rede
			XSSFSheet sheet = workbook.getSheetAt(2);	
			Cell cell = null;
			
			//Update the value of cell
			cell = sheet.getRow(1).getCell(8);
			cell.setCellValue("Avaliação");
			
			
			
			//write out the XLSM
			//FileOutputStream out = new FileOutputStream(f);
			//workbook.write(out);
			//out.close();

			
			//FileOutputStream os = new FileOutputStream(f); 
			//workbook.write(os); 
			
			//FileOutputStream fileOut = new FileOutputStream(f);

			//write this workbook to an Outputstream.
			//workbook.write(fileOut);
			//fileOut.flush();
			//fileOut.close();
			System.out.println("Writing on XLSX file Finished ...");

			
			file.close();
			workbook.close();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
}
