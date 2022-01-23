import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class FormulaCPGF {

	@SuppressWarnings({ "rawtypes", "incomplete-switch", "unused" })
	public static void main(String[] args) throws IOException {

		String path="C:\\Users\\sueli\\eclipse-workspace\\Projeto-CPGF\\ValorTotal_CGPF.xls";
		
		FileInputStream fis=new FileInputStream(path);
		
		HSSFWorkbook workbook=new HSSFWorkbook(fis);
		
		
		HSSFSheet sheet = workbook.getSheetAt(0);
		Iterator<Sheet> rowIterator = workbook.iterator(); // Cria objeto iterador
		Row row3 = sheet.getRow(1);
		
		
	
		Row header = sheet.createRow(0);
		header.createCell(0).setCellValue("Total");
		
		Row dataRow = sheet.createRow(2);
		dataRow.createCell(2).setCellFormula("SUM(A2:A8175)");
		  System.out.println(dataRow);
		 
		
		fis.close();
		
		FileOutputStream fos=new FileOutputStream(path);
		workbook.write(fos);
		
		workbook.close();
		fos.close();

                           
Iterator iterator=sheet.iterator();
		
		while(iterator.hasNext())
		{
			HSSFRow row1=(HSSFRow) iterator.next();
			
			Iterator cellIterator=row1.cellIterator();
			
			while(cellIterator.hasNext())
			{
				HSSFCell cell=(HSSFCell) cellIterator.next();
				
				switch(cell.getCellType())
				{
				case STRING: System.out.print(cell.getStringCellValue()); break;
				case NUMERIC: System.out.print(cell.getNumericCellValue());break;
				
				}
				System.out.print(" |  ");
			}
			System.out.println();
		}
                            
                    }
	}
            

            
    
    
    

    
    
	


