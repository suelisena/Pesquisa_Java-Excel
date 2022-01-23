import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;



public class SheetFormulaCPGF {

@SuppressWarnings({ "incomplete-switch", "unused", "null", "resource" })
public static void main(String[] args) throws IOException {
		
		FileInputStream file=new FileInputStream("C:\\Users\\sueli\\eclipse-workspace\\Projeto-CPGF\\ValorTotal_CGPF.xls");
		
	
		HSSFWorkbook workbook=new HSSFWorkbook(file);
		
		DataFormatter formatter = new DataFormatter();
	  
		Sheet sheet1 = workbook.getSheetAt(0);
		Row row3 = sheet1.getRow(0);
		
		
	
		
		for (Row row : sheet1) {
		    for (Cell cell : row) {
		        
				
				CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
		        System.out.print(cellRef.formatAsString());
		        System.out.print(" - ");
		      
		        String text = formatter.formatCellValue(cell);
		        System.out.println(text);
		       
		        switch (cell.getCellType()) {
		            case STRING:
		                System.out.println(cell.getRichStringCellValue().getString()  + "\t\t");
		                break;
		       
		            case NUMERIC: 
		            	System.out.println(cell.getNumericCellValue()  + "\t\t");
		            break;
		            
		            case FORMULA:
		                System.out.println(cell.getCellFormula() + "\t\t");
		                break;
		          
		        }
		        
		  
		      
		        System.out.print(" |  ");
			}
			System.out.println();
		    }
		    
		}
}
