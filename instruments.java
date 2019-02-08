//javac -classpath .:/home/exp.exactpro.com/oleg.legkov/Corrector/lib/poi-3.9-20121203.jar instruments.java
//java -classpath .:/home/exp.exactpro.com/oleg.legkov/Corrector/lib/poi-3.9-20121203.jar instruments

import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.w3c.dom.Attr;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

public class instruments{
	private static FileWriter log;
	
	public static void main(String[] args) throws Exception{
		log = new FileWriter("createXMLDB.log");
		log.write("\ncreateXMLDB Func\n");
		
		FileInputStream fisInformation = new FileInputStream("/home/exp.exactpro.com/oleg.legkov/Corrector/instruments.xls");
		HSSFWorkbook documentInformation = new HSSFWorkbook(fisInformation);
		HSSFSheet sheet = documentInformation.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();
		
		//System.out.println(lastRow);
		
		LinkedHashMap<String, String[]> map = new LinkedHashMap<String, String[]>();
		
		for(int i = 1; i < lastRow + 1; i++){
			Row row = sheet.getRow(i);
			String[] arr = new String[29];

			if(row.getCell(0) == null) continue;

			switch(row.getCell(0).getStringCellValue()){
				case("Active"): arr[0] = "0"; break;
				case("Instrument Suspended"): arr[0] = "1"; break;
				case("Inactive"): arr[0] = "2"; break;
				case("Halt"): arr[0] = "3"; break;
				case("SI Quote Prohibited"): arr[0] = "4"; break;
			}
			
			
			
			arr[1] = (Integer.toString((int)row.getCell(1).getNumericCellValue()));
			arr[2] = row.getCell(2).getStringCellValue();
			arr[3] = row.getCell(3).getStringCellValue();
			arr[4] = row.getCell(7).getStringCellValue();			
			arr[5] = row.getCell(9).getStringCellValue();
			
			if(row.getCell(17) == null) arr[6] = "#"; 
			else {
				switch(row.getCell(17).getStringCellValue()){
					case(""): arr[6] = "#"; break;
					case("EUA"): arr[6] = "0"; break;
					case("CER"): arr[6] = "1"; break;
					case("ERU"): arr[6] = "2"; break;
					case("EUAA"): arr[6] = "3"; break;
					case("OTHR"): arr[6] = "4"; break;
				}
			}
			
			switch(row.getCell(18).getStringCellValue()){
				case("MDMS"): arr[7] = "1"; break;
				case("UnaVista"): arr[7] = "2"; break;
			}
			
			
			if(row.getCell(20) == null) arr[8] = "#"; else arr[8] = row.getCell(20).getStringCellValue();
			
			//if(row.getCell(21) == null) arr[9] = "#"; else arr[9] = row.getCell(21).getStringCellValue();
			if(row.getCell(21) == null) arr[9] = "#"; //StrikePrice
			else {
				if(row.getCell(21).getCellType() == Cell.CELL_TYPE_STRING){
					switch(row.getCell(21).getStringCellValue()){
						case(""): arr[9] = "#"; break;
						default: arr[9] = row.getCell(21).getStringCellValue(); break;
					}
				}
				if(row.getCell(21).getCellType() == Cell.CELL_TYPE_NUMERIC) arr[9] = String.valueOf(row.getCell(21).getNumericCellValue());
			}
			
			
			//if(row.getCell(22) == null) arr[10] = "#"; else arr[10] = row.getCell(22).getStringCellValue();
			if(row.getCell(22) == null) arr[10] = "#"; //Yield
			else {
				if(row.getCell(22).getCellType() == Cell.CELL_TYPE_STRING){
					switch(row.getCell(22).getStringCellValue()){
					case(""): arr[10] = "#"; break;
					default: arr[10] = row.getCell(22).getStringCellValue(); break;
					}
				}
				if(row.getCell(22).getCellType() == Cell.CELL_TYPE_NUMERIC) arr[10] = String.valueOf(row.getCell(22).getNumericCellValue());
			}
			
			
			//if(row.getCell(23) == null) arr[11] = "#"; else arr[11] = (Integer.toString((int)row.getCell(23).getNumericCellValue()));
			if(row.getCell(23) == null) arr[11] = "#"; 
			else {
				switch(row.getCell(23).getStringCellValue()){
				case(""): arr[11] = "#"; break;
				default: arr[11] = row.getCell(23).getStringCellValue(); break;
				}
			}
			
			
			//if(row.getCell(24) == null) arr[12] = "#"; else arr[12] = row.getCell(24).getStringCellValue();
			if(row.getCell(24) == null) arr[12] = "#"; 
			else {
				switch(row.getCell(24).getStringCellValue()){
				case(""): arr[12] = "#"; break;
				default: arr[12] = row.getCell(24).getStringCellValue(); break;
				}
			}
			
			
//			try{
			//if(row.getCell(26) == null) arr[13] = "#"; else arr[13] = row.getCell(26).getStringCellValue();
			if(row.getCell(26) == null) arr[13] = "#"; //ParValue
			else {
				if(row.getCell(26).getCellType() == Cell.CELL_TYPE_NUMERIC)
					arr[13] = Double.toString(row.getCell(26).getNumericCellValue());
				else{
					switch(row.getCell(26).getStringCellValue()){
						case(""): arr[13] = "#"; break;
						default: arr[13] = row.getCell(26).getStringCellValue(); break;
					}
				}
			}
//			}catch(Exception e){}
//			finally{ arr[13] = "#"; }
			
			
			//if(row.getCell(27) == null) arr[14] = "#"; else arr[14] = row.getCell(27).getStringCellValue();
			if(row.getCell(27) == null) arr[14] = "#"; 
			else {
				switch(row.getCell(27).getStringCellValue()){
				case(""): arr[14] = "#"; break;
				default: arr[14] = row.getCell(27).getStringCellValue(); break;
				}
			}
			
			
			
						
			if(row.getCell(28) == null) arr[15] = "#"; 
			else{
				if(row.getCell(28).getCellType() == Cell.CELL_TYPE_STRING) arr[15] = row.getCell(28).getStringCellValue();
				else arr[15] = (Integer.toString((int)row.getCell(28).getNumericCellValue()));
			}
						
			arr[16] = row.getCell(29).getStringCellValue();
			
			//if(row.getCell(30) == null) arr[17] = "#"; else arr[17] = row.getCell(30).getStringCellValue();
			if(row.getCell(30) == null) arr[17] = "#"; 
			else {
				switch(row.getCell(30).getStringCellValue()){
				case(""): arr[17] = "#"; break;
				default: arr[17] = row.getCell(30).getStringCellValue(); break;
				}
			}
			
			
			//if(row.getCell(31) == null) arr[18] = "#"; else arr[18] = row.getCell(31).getStringCellValue();
			if(row.getCell(31) == null) arr[18] = "#"; 
			else {
				switch(row.getCell(31).getStringCellValue()){
				case(""): arr[18] = "#"; break;
				default: arr[18] = row.getCell(31).getStringCellValue(); break;
				}
			}
			
						
			arr[19] = row.getCell(32).getStringCellValue();

			
			if(row.getCell(34).getCellType() == Cell.CELL_TYPE_STRING) arr[20] = row.getCell(34).getStringCellValue();
			if(row.getCell(34).getCellType() == Cell.CELL_TYPE_NUMERIC) arr[20] = String.valueOf(row.getCell(34).getNumericCellValue());
			


			arr[21] = (Integer.toString((int)row.getCell(35).getNumericCellValue()));
			
			
			arr[22] = (Integer.toString((int)row.getCell(38).getNumericCellValue()));
			
			//if(row.getCell(40) == null) arr[23] = "#"; else arr[23] = row.getCell(40).getStringCellValue();
			if(row.getCell(40) == null) arr[23] = "#"; 
			else {
				switch(row.getCell(40).getStringCellValue()){
				case(""): arr[23] = "#"; break;
				default: arr[23] = row.getCell(40).getStringCellValue(); break;
				}
			}
			
			
			//if(row.getCell(41) == null) arr[24] = "#"; else arr[24] = (Integer.toString((int)row.getCell(41).getNumericCellValue()));
			if(row.getCell(41) == null) arr[24] = "#"; //SecurityType
			else {
				switch(row.getCell(41).getStringCellValue()){
					case(""): arr[24] = "#"; break;
					default: arr[24] = row.getCell(41).getStringCellValue(); break;
				}
			}
			
			
//			if(row.getCell(52) == null) arr[25] = "#"; 
//			else { 
//				arr[25] = row.getCell(52).getStringCellValue();
//				arr[25] = arr[25].split("T")[0];
//				arr[25] = arr[25].replaceAll("-", "");
//				//System.out.println("EXPIRATION DATE: " + arr[25]);
//			}
			
			if(row.getCell(52) == null) arr[25] = "#"; 
			else {
				switch(row.getCell(52).getStringCellValue()){
					case(""): arr[25] = "#"; break;
					default: 
						arr[25] = row.getCell(52).getStringCellValue();
						arr[25] = arr[25].split("T")[0];
						arr[25] = arr[25].replaceAll("-", "");
						//arr[25] = row.getCell(52).getStringCellValue(); break;
				}
			}
			
			
			
			if(row.getCell(62) != null){
				switch(row.getCell(62).getStringCellValue()){
					case("No"): arr[26] = "0"; break;
					case("Yes"): arr[26] = "1"; break;
					case(""): arr[26] = "#"; break;
				}
			}  else arr[26] = "#";

			
			if((int)row.getCell(1).getNumericCellValue() % 2 == 0) arr[27] = "1"; else arr[27] = "2";
			//if(row.getCell(10).getStringCellValue().matches("Shares")) arr[28] = "1"; else arr[28] = "0";
			switch(row.getCell(10).getStringCellValue()){
				case("Shares"): 
				case("Depository Receipts"): 
				case("ETF:s"): 
				case("Other equity-like financial instrument"): 
				case("Certificates"): arr[28] = "1"; break; 
				default: arr[28] = "0"; break;
			}
			
			map.put(arr[1], arr);
		}
		
		//Print
		String[] arr = map.get("120000031");
		String prefix = "_1";

		
		System.out.println("\t\t\tInstrument Status\t\t\tInstrumentStatus" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[0]);
		System.out.println("\t\t\tInstrument\t\t\tInstrument" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[1]);
		System.out.println("\t\t\tInstrument ISIN\t\t\tInstrumentISIN" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[2]);
		System.out.println("\t\t\tShort Name\t\t\tShortName" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[3]);
		System.out.println("\t\t\tUnderlying\t\t\tUnderlying" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[4]);
		System.out.println("\t\t\tCFI Code\t\t\tCFICode" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[5]);
		
		if(arr[6].matches("#")) System.out.println("\t\t\tEmission Allowance Type\t\t\tEmissionAllowanceType" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[6]);
		else System.out.println("\t\t\tEmission Allowance Type\t\t\tEmissionAllowanceType" + prefix + "\t\tSetStatic\t\t\t\tInteger\t" + arr[6]);
		System.out.println("\t\t\tInstrument Source\t\t\tInstrumentSource" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[7]);
		System.out.println("\t\t\tInstrument Currency\t\t\tInstrumentCurrency" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[8]);
		System.out.println("\t\t\tStrike Price\t\t\tStrikePrice" + prefix + "\t\tSetStatic\t\t\t\tBigDecimal\t" + arr[9]);
		System.out.println("\t\t\tYield\t\t\tYield" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[10]);
		
		
		if(!arr[11].matches("#")) System.out.println("\t\t\tClosing Price\t\t\tClosingPrice" + prefix + "\t\tSetStatic\t\t\t\tBigDecimal\t" + arr[11]);
		else System.out.println("\t\t\tClosing Price\t\t\tClosingPrice" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[11]);
		
		System.out.println("\t\t\tReference Price\t\t\tReferencePrice" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[12]);
		
		if(arr[13].matches("#")) System.out.println("\t\t\tPar Value\t\t\tParValue" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[13]);
		else System.out.println("\t\t\tPar Value\t\t\tParValue" + prefix + "\t\tSetStatic\t\t\t\tBigDecimal\t" + arr[13]);
		
		System.out.println("\t\t\tPar Value Currency\t\t\tParValueCurrency" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[14]);
		System.out.println("\t\t\tDenominated Par Value\t\t\tDenominatedParValue" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[15]);
		System.out.println("\t\t\tCountry Of Issue\t\t\tCountryOfIssue" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[16]);
		System.out.println("\t\t\tMarket ID\t\t\tMarketID" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[17]);
		System.out.println("\t\t\tMarket Source\t\t\tMarketSource" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[18]);
		System.out.println("\t\t\tMarket Code\t\t\tMarketCode" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[19]);
		System.out.println("\t\t\tSymbol\t\t\tSymbol" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[20]);
		System.out.println("\t\t\tADT\t\t\tADT" + prefix + "\t\tSetStatic\t\t\t\tBigDecimal\t" + arr[21]);
		System.out.println("\t\t\tNMS\t\t\tNMS" + prefix + "\t\tSetStatic\t\t\t\tBigDecimal\t" + arr[22]);
		System.out.println("\t\t\tSegment\t\t\tSegment" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[23]);
		System.out.println("\t\t\tSecurity Type\t\t\tSecurityType" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[24]);
		System.out.println("\t\t\tExpiration Date\t\t\tExpirationDate" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[25]);
		System.out.println("\t\t\tLSEG Clearing Type\t\t\tLSEGClearingType" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[26]);
		System.out.println("\t\t\tOrigin\t\t\tOrigin" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[27]);
		System.out.println("\t\t\tEquity\t\t\tEquity" + prefix + "\t\tSetStatic\t\t\t\tString\t" + arr[28]);
		
		//System.out.println("Map: " + map);
	}
}
