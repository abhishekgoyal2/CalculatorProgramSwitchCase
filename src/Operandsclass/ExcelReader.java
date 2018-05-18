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
public static int result;
public static int c;
		// TODO Auto-generated method stub
		public static XSSFWorkbook wb;
		public static XSSFSheet wbsheet;
		public static XSSFRow row;
		public static XSSFCell cell;
		public static XSSFCell opcell;
		public static FileInputStream fis;
		public static FileOutputStream fileout;
		public static String path;
		
		
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
			return "Success";
		}
	
	//function to Write value  to excel
	public static  void fileOutStream1() throws IOException
	 {
		 fileout = new FileOutputStream("D:\\automationXpath\\Cal.xlsx");
//		 System.out.print("calling filepath");
           wb.write(fileout);
fileout.close();
	 }
		
		//function to call specific operand 
		 @SuppressWarnings({ "deprecation", "null" })
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
				 opcell=wbsheet.getRow(i).getCell(0);
				XSSFCell Acell=wbsheet.getRow(i).getCell(1);//Cells under ColA
				XSSFCell Acel2=wbsheet.getRow(i).getCell(2);//Cells under ColB
			
					
				
//			Checck Opcell type and call exception
				switch(opcell.getCellType())
				{
					case Cell.CELL_TYPE_STRING:
						System.out.println(opcell.getStringCellValue());
						
						//exception handle if a cell is null 
					
						operand =opcell.getStringCellValue();
						
						//exception handle if a cell is null 
						 try{
									valA=(int) Acell.getNumericCellValue();
									
								}
								catch(Exception e)
								{
									System.out.println(e.getMessage());
									System.out.println("no cell exist as null pointer exception occured");
									e.printStackTrace();
								}
						 
						//exception handle if a cell is null 
						 try{
								
							valB=(int) Acel2.getNumericCellValue();
									
								}
								catch(Exception e)
								{
									System.out.println(e.getMessage());
									System.out.println("no cell exist as null pointer exception occured");
									e.printStackTrace();
								}
						 
						
						 
						 Row rowcal = wbsheet.getRow(i);
						 
						 //create cell 
						 Cell cell = rowcal.getCell(3); 
						 
						 if (cell==null)
						     cell = rowcal.createCell(3);
						 cell.setCellType(Cell.CELL_TYPE_NUMERIC);
						 System.out.println("value of String in Operand Column is" +" = " + operand);
						 
					//Checking if this valid operand and based on it calling valid operator  inside switch case 
				if (operand.equalsIgnoreCase("+")||operand.equalsIgnoreCase("-")||operand.equalsIgnoreCase("*")||operand.equalsIgnoreCase("/")||operand.equalsIgnoreCase("%"))
							{
								//passing operand using switch case
							switch(operand)
							{
						//call  plus operand class function based on operator in excel
								case "+":
									System.out.println("Call plus operand");
									Add add = new Add();
								
									 c = add.calculate(operand,valA,valB);
									System.out.println("Final value : " +c);
								
								cell.setCellValue(c);
								  result =(int) cell.getNumericCellValue();
								
								System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +result);
								 fis.close();
								 //update ,write values in result column
								 fileOutStream1();
								break;
								//call  minus operand class function based on operator in excel	
							case "-":
								System.out.println("Call minus operand");
								
								Minus min = new Minus();
								
								  c = min.calculate(operand,valA,valB);
								System.out.println("Final value : " +c);
								 cell.setCellValue(c);
							
								  result =(int) cell.getNumericCellValue();

									System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +result);
									 fis.close();
									 //update ,write values in result column
									 fileOutStream1();
									break;
									//call  multiplication operand class function based on operator in excel	
								case "*":
									System.out.println("Call multiplication operand");
									
									
									Multiplication mul = new Multiplication();
									
									  c = mul.calculate(operand,valA,valB);
									System.out.println("Final value : " +c);
									cell.setCellValue(c);
									
									  result =(int) cell.getNumericCellValue();
				
										System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +result);
										 fis.close();
										 //update ,write values in result column
										 fileOutStream1();
									break;
									//call  division operand class function based on operator in excel
									
								case "/":
									
									Division div = new Division();
									
									//catch run time Arithmetic exception
									
									try{
									  c = div.calculate(operand,valA,valB);
									}
									catch (ArithmeticException e) {
										// TODO: handle exception
										
										System.out.println("Division by zero Error ocured");
									}
									  
									  
									System.out.println("Final value : " +c);
									cell.setCellValue(c);
									
									  result =(int) cell.getNumericCellValue();
				
										System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +result);
									
										 fis.close();
										 //update ,write values in result column
										 fileOutStream1();
										 
									break;
									
									//call  modulo operand class function based on operator in excel
								case "%":
									System.out.println("Call  Modulo operand");
									
									Modulo mod = new Modulo();

									c = mod.calculate(operand,valA,valB);
									System.out.println("Final value : " +c);
									cell.setCellValue(c);
									
									  result =(int) cell.getNumericCellValue();
										System.out.println("Value of cell after calcultion : " + " of .."+ operand + "..  " +result);
				
									 fis.close();
									 fileOutStream1();
									 break;
							}
						
				}
				
				//Checking not a valid operand print in console for not a valid operand and skip calcualtor function
						 else{
							 
 System.out.println("this is not a valid operand skip out the Calculator function");
							
				}
			
						break;
					case Cell.CELL_TYPE_BOOLEAN:
							System.out.println(opcell.getBooleanCellValue());
							try {
						opcell.getBooleanCellValue();
							}
							catch(Exception e)
							{
				System.out.println("not String type value");
								System.out.println(e.getMessage());
								e.printStackTrace();
							}
								 break;
					case Cell.CELL_TYPE_BLANK:
						System.out.println("blank");
						try {
						System.out.println("not a valid operand");
								}
								catch(Exception e)
								{
//									e.printStackTrace();
								}
						break;
					case Cell.CELL_TYPE_ERROR:
						System.out.println("error");
						try {
							
								}
								catch(Exception e)
								{
									e.printStackTrace();
								}
						break;
					case Cell.CELL_TYPE_FORMULA:
						System.out.println(opcell.getCellFormula());
						try {
							
								}
								catch(Exception e)
								{
									e.printStackTrace();
								}
						break;
					case Cell.CELL_TYPE_NUMERIC:
						System.out.println(opcell.getNumericCellValue());
						try {
							
								}
								catch(Exception e)
								{
									e.printStackTrace();
								}
						break;
						
				}
				

				}
			
	return operand;			
			
}
		
}