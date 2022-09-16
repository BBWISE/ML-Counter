import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelCounter {

	public static void main(String[] args) {
		
	//	String fileName = "absence";
		String fileName = "presence";
		File arrfile = new File("C:\\Users\\B.B. WiSE\\Desktop\\BC\\"+fileName+".xlsx");
		
			
		FileInputStream excelFIS = null;
		BufferedInputStream excelBIS = null;
		XSSFWorkbook excelJTableImport = null;
		
		try {
			excelFIS = new FileInputStream(arrfile);
			excelBIS = new BufferedInputStream(excelFIS);
		
			excelJTableImport = new XSSFWorkbook(excelBIS);
			XSSFSheet excelSheet = excelJTableImport.getSheetAt(0);
			
			int number = 0;
			for(int row=0;row<=excelSheet.getLastRowNum();row++) {
				
				XSSFRow excelRow = excelSheet.getRow(row);
				
				
				XSSFCell excelSN = excelRow.getCell(8);
				
				String ss = excelSN.toString();
				
				double nTake1 = Double.parseDouble(ss);
				
		//		if(nTake1 >=121 && nTake1 <=130 ) {
		//		if(nTake1 <=70 ) {
				if(nTake1 >=1100 ) {
		//		if(nTake1 ==1 ) {
		//		if(nTake1 < 100 ) {
		//		if(nTake1 >=1000 && nTake1 <1100 ) {
					
					System.out.println(row+" === "+nTake1);
					number +=1;
				}
				
			}
			System.out.println(number);
		}
		catch(Exception e) {
			System.out.println(e);
		}
		
	}

}
