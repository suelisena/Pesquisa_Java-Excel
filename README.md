# Pesquisa_Java-Excel





A – Com suas palavras explique o que é lavagem de dinheiro
Lavagem de dinheiro é a prática financeira de esconder a origem ilícita de alguns bens e dinheiro recebido, para dificultar o rastreamento da origem. Emitir notas fiscais de serviços que não foram prestados, notas falsa ou comprar bens em nome de outras pessoas (laranjas) para e serem declarados como patrimônio licito.


B – O que é Cartão de Pagamento do Governo Federal (CPGF), e qual a sua finalidade.
O Cartão de Pagamento do Governo Federal (CPGF) é um meio de pagamento utilizado pelo governo que funciona de forma similar ao cartão de crédito que utilizamos em nossas vidas, porém dentro de limites e regras específicas. O governo utiliza o CPGF para pagamentos de despesas próprias, que possam ser enquadradas como suprimento de fundos.


C – Quem pode utilizar o CPGF?
O CPGF é utilizado pelo Administração Pública, Servidores Público.

D – Qual a URL onde é possível fazer o download dos arquivos do CPGF?
URL:
https://www.portaltransparencia.gov.br/download-de-dados

E – Qual a URL da paǵina com a descrição dos campos (ou dicionário de dados) da CPGF?
https://www.portaldatransparencia.gov.br/pagina-interna/603393-dicionario-de-dados-cpgf

F – É possível identificar o nome e o documento do portador do CPGF, em todas as
movimentações ou há movimentações onde não é possível identificar o portador?
Em algumas movimentações não é possível idenficar o portador, a casos em que aparece como “Sigiloso”.


G – É possível identificar o Órgão do portador do CPGF?
Sim, é possível localizar o Órgão do portador.

H - Qual o nome do Órgão cujo código é 20402?
Código 20402 -  Agência Espacial Brasileira

I - É possível identificar o Nome e Documento (CNPJ) dos favorecidos pela utilização do CPGF?
Em alguns casos é possível a identificação do nome e do documento.


J – É possível identificar a data e o valor das movimentações financeiras do CPGF, em todas as movimentações? Ou há movimentações onde não é possível identificar as datas e ou valores?
Em alguns dados não é possível a identificação, aparece como “Informações protegidas por sigilo“.
K (código) – Qual a soma total das movimentações utilizando o CPGF?
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
            


L (código) – Qual a soma das movimentações sigilosas ?
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

M (código) – Qual o Órgão que mais realizou movimentações sigilosas no período e qual o valor (somado)?

N (código) – Qual o nome do portador que mais realizou saques no período? Qual a soma dos saques realizada por ele? Qual o nome do Órgão desse portador?


O (código) – Qual o nome do favorecido que mais recebeu compras realizadas utilizando o CPGF?

P - Descreva qual a abordagem utilizada para desenvolver o código para os ítens de K a O.
Foi utilizado a linguagem de programação Java, juntamente com o Java Apache Poi, onde foi inseridas as bibliotecas em jar, para utilizar o HSSFWorkbook para a leitura de todo o documento da planilha 202110_CPGF, inserido no código pelo FileInputStream, utilizando o Interator e o while para percorrer e apresentar os dados no system.out.println. O setCellFormula foi utilizado para inserir a formula e somar o valor total.
