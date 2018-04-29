package br.com.customizze.planilhaexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

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
			XSSFSheet sheet = workbook.getSheetAt(0);
			
			List products = new ArrayList<>();
			
			Iterator rowIterator = sheet.rowIterator();
			
			/*while (rowIterator.hasNext()) {
				Row row = (Row) rowIterator.next();
				
				// Descantando a primeira linha com o header
				if(row.getRowNum() == 0){
					continue;
				}
		
				Iterator cellIterator = row.cellIterator();
				
				while (cellIterator.hasNext()) {
					
					Cell cell = (Cell) cellIterator.next();

					Product product = new Product();
					
					switch (cell.getColumnIndex()) {
						case 0:
							product.setId(((Double)cell.getNumericCellValue()).longValue());
						break;
						case 1:
							product.setName(cell.getStringCellValue());
						break;
						case 2:
							product.setPrice(cell.getNumericCellValue());
						break;
					}
					products.add(product);
				}
			
			}
			
			for (Product product : products) {
				System.out.println(product.getId() + " – " + product.getName() + " – " + product.getPrice());
			}*/
			
			file.close();
			workbook.close();
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
	}
}
