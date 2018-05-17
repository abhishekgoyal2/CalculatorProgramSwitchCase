package Operandsclass;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
public class ExcelReader {

static String operand;
static int valA;
static int valB;
public int result;
		// TODO Auto-generated method stub
		public static XSSFWorkbook wb;
		public static XSSFSheet wbsheet;
		public static XSSFRow row;
		public static XSSFCell cell;

		public static FileInputStream fis;
		public static FileOutputStream fileout;
		public String path;
		
	

		
/*		public static void main(String[] args) throws IOException {
			// TODO Auto-generated method stub
			
	              
		fis =new FileInputStream("D:\\automationXpath\\Cal.xlsx");

		      
		operand=CalOperatorSaveResult(operand);
	System.out.println();
		
	fis.close();	
	fileout.close();
		
}*/
	
		//function to call file which calls and reads excel and update result based on operand
	public	String 	getfile() throws IOException{
		
				fis =new FileInputStream("D:\\automationXpath\\Cal.xlsx");
			operand=CalOperatorSaveResult(operand);
			fis.close();	
			fileout.close();
			return "Sucees";
		}
		
		
		//function to call specific operand 
		 public static  String CalOperatorSaveResult(String operand) throws IOException
			{
			//workbook and sheet read
				wb=new XSSFWorkbook(fis);
				
				 wbsheet=wb.getSheetAt(0);
			 
				 //get row number 
				 int row=wbsheet.getLastRowNum();
				 System.out.println("number of Rows--  " + " = " + row);
				 //get col number 

			int col=wbsheet.getRow(0).getLastCellNum();	
			
			System.out.println("number of Colmuns--  " + " = " + col);
			
			//for loop to read value ,cal function based on operand call from operand class and update result in excel
			
			for(int i=1;i<=row;i++)
			{
				XSSFCell opcell=wbsheet.getRow(i).getCell(0);
				XSSFCell Acell=wbsheet.getRow(i).getCell(1);//Cells under ColA
				XSSFCell Acel2=wbsheet.getRow(i).getCell(2);//Cells under ColB
				
				 operand =opcell.getStringCellValue();
				
				 valA=(int) Acell.getNumericCellValue();
				 valB=(int) Acel2.getNumericCellValue();
				 
				
				 Row rowcal = wbsheet.getRow(i);
				 
				 //create cell 
				 Cell cell = rowcal.getCell(3); 
				 
				 if (cell == null)
				     cell = rowcal.createCell(3);
				 cell.setCellType(Cell.CELL_TYPE_NUMERIC);

				 
				
				System.out.println("Operand " + operand);
				//call  plus operand class function based on operator in excel
				if(operand.equalsIgnoreCase("+"))
				{
					System.out.println("Call plus operand");
					Add add = new Add();
				
					 int c = add.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
				
				cell.setCellValue(c);
				 int result =(int) cell.getNumericCellValue();
				
				System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +result);
				 fis.close();
				 //update ,write values in result column
				 fileOutStream1();


		            }
			
				//call  minus operand class function based on operator in excel

				else if(operand.equalsIgnoreCase("-"))
				{
					System.out.println("Call minus operand");
					
					Minus min = new Minus();
					
					 int c = min.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
					 cell.setCellValue(c);
				
					 int result =(int) cell.getNumericCellValue();

						System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +result);
						 fis.close();
						 //update ,write values in result column
						 fileOutStream1();
}
				//call  multiplication operand class function based on operator in excel

				else if(operand.equalsIgnoreCase("*"))
				{
					System.out.println("Call multiplication operand");
					
					
					Multiplication mul = new Multiplication();
					
					 int c = mul.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
					cell.setCellValue(c);
					
					 int result =(int) cell.getNumericCellValue();

						System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +result);
						 fis.close();
						 //update ,write values in result column
						 fileOutStream1();

  }
				//call  division operand class function based on operator in excel

				else if(operand.equalsIgnoreCase("/"))
				{
					System.out.println("Call division operand");
					
					Division div = new Division();
					
					 int c = div.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
					cell.setCellValue(c);
					
					 int result =(int) cell.getNumericCellValue();

						System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +result);
						 fis.close();
						 //update ,write values in result column
						 fileOutStream1();
}
				//call  modulo operand class function based on operator in excel

				else if(operand.equalsIgnoreCase("%"))
				{
					System.out.println("Call  Modulo operand");

					
					Modulo mod = new Modulo();
					
					 int c = mod.calculate(operand,valA,valB);
					System.out.println("Final value : " +c);
					cell.setCellValue(c);
					
					 int result =(int) cell.getNumericCellValue();
						System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +result);

					 fis.close();
					 fileOutStream1();

}
				//Exceptional handling if the operand column is not having appropriate values.

				else {
					try {
					
						System.out.println("Not a valid operand ");
					}
					catch(Exception e)
					{
						e.printStackTrace();
					}
				}
				

			}
			

			return operand;
			
			
			}
		
		 public static  void fileOutStream1() throws IOException
		 {
			 fileout = new FileOutputStream("D:\\automationXpath\\Cal.xlsx");
			 System.out.print("calling filepath");
	            wb.write(fileout);
fileout.close();
		 }
		
	
	}