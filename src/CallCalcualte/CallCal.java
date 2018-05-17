package CallCalcualte;

import java.io.IOException;

import org.apache.log4j.Logger;

import Operandsclass.ExcelReader;

public class CallCal {
	public static String operand  ;

public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
	//intilaizing logger object
		Logger log =Logger.getLogger("sdevpinbyLoger");
		
		ExcelReader excelReader = new  ExcelReader();
		//debugs logs in file 
	log.debug("calling constructor of Class ExcelReader");
	//calling get file method of Excel Reader class to read and update.
	String responce =	excelReader.getfile();
	
	//debug logs in seperate file
	log.debug("calling get file to get operand and update the result for operand");
	System.out.println("sucesfully updated"+responce);
	
	log.info("this is calling excel read and update");

		
		}

}
