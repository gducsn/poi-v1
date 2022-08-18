package com.start;

import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileCreation {
	
	private FileCreation() {}
	
	 static void getFile(String PATH){
		 try (XSSFWorkbook workbook = new XSSFWorkbook()) {
				XSSFSheet sheet = workbook.createSheet("Lista");

				
				// crei le righe
				for (int i = 0; i <= 3; i++) {
					sheet.createRow(i);
				}
				
				// prendi la prima riga, selezioni le celle di volta in volta e aggiungi il valore
				sheet.getRow(0).createCell(0).setCellValue("valori: ");
				sheet.getRow(0).createCell(1).setCellValue(12);
				sheet.getRow(0).createCell(2).setCellValue(4);
				sheet.getRow(0).createCell(3).setCellValue(16);
				// prendi la prima riga, selezioni le celle di volta in volta e aggiungi il valore
				sheet.getRow(1).createCell(0).setCellValue("valori: ");
				sheet.getRow(1).createCell(1).setCellValue(6);
				sheet.getRow(1).createCell(2).setCellValue(7);
				sheet.getRow(1).createCell(3).setCellValue(13);
				
				
				// questa sarebbe la terza riga (in base 0)
				sheet.getRow(2).createCell(0).setCellValue("risultati: ");
				
				// per usare le forumule:
				// selezioni la riga e crei la cella
				// recuperi la cella e inserisi la formula
				// devi convalidare la formula
				// deve inserire il valore con 'evaluate'
				sheet.getRow(2).createCell(1);
				sheet.getRow(2).getCell(1).setCellFormula("SUM(B1:B2)");
				XSSFFormulaEvaluator a = workbook.getCreationHelper().createFormulaEvaluator();
				a.evaluateFormulaCell(sheet.getRow(2).getCell(1));
				///
				
				///
				sheet.getRow(2).createCell(2);
				sheet.getRow(2).getCell(2).setCellFormula("SUM(C1:C2)");
				XSSFFormulaEvaluator b = workbook.getCreationHelper().createFormulaEvaluator();
				b.evaluateFormulaCell(sheet.getRow(2).getCell(2));
				///
				
				///
				sheet.getRow(2).createCell(3);
				sheet.getRow(2).getCell(3).setCellFormula("SUM(D1:D2)");
				XSSFFormulaEvaluator c = workbook.getCreationHelper().createFormulaEvaluator();
				c.evaluateFormulaCell(sheet.getRow(2).getCell(3));
				///
				
				
				// crei il file con il path
				FileOutputStream output = new FileOutputStream(PATH);
				// chiudi il foglio 
				workbook.write(output);
				// chiudi il file
				output.close();
				
				System.out.println("Completed...");

			} catch (Exception e) {
				e.getStackTrace();
			}
		 
		 
	 }

}
