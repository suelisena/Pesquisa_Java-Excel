import java.io.FileInputStream;
import java.io.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;
import java.util.Iterator;

public class GerenciadorCPGF { 
	
     

		
		public static void main(String[] args) throws IOException{
               
        	FileInputStream input_document = new FileInputStream(new File("C:\\Users\\sueli\\eclipse-workspace\\Projeto-CPGF\\202110_CPGF.xls")); // Ler documento XLSX - formato Office 2007, 2010    
                
        	HSSFWorkbook my_xls_workbook = new HSSFWorkbook(input_document); // Ler a pasta de trabalho do Excel em um objeto de instância 
               
                HSSFSheet my_worksheet = my_xls_workbook.getSheetAt(0); // Isso lerá a planilha para nós em outro objeto 
                
                Iterator<Row> rowIterator = my_worksheet.iterator(); // Cria objeto iterador
                @SuppressWarnings("unused")
				Row row1 = my_worksheet.getRow(0);

                while(rowIterator.hasNext()) {
                        Row row = rowIterator.next(); // Ler linhas do documento Excel       
                        Iterator<Cell> cellIterator = row.cellIterator();// Lê cada coluna para cada linha que é READ 
                       
                        
                        while(cellIterator.hasNext()) {
                                        Cell cell1 = cellIterator.next(); // Busca 
                                        switch(cell1.getCellType()) {  // Identifica o tipo CELL 
                                        case NUMERIC:
                                                System.out.print(cell1.getNumericCellValue() + "\t\t"); // Identifica o tipo CELL
                                                break;
                                        case STRING:
                                                System.out.print(cell1.getStringCellValue() + "\t\t"); // imprime o valor numérico
                                                break;
                                        }
                                        
                                        
                                }
                        

                        
                System.out.println(""); // Para iterar até a próxima linha
                
                
           
                
                }
                input_document.close(); // Fechar o arquivo XLS aberto para impressão 
        }
}