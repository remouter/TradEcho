//javac -cp .:/home/exp.exactpro.com/oleg.legkov/Desktop/Corrector_3.0/lib/poi-3.9-20121203.jar Bugs.java
//java -cp .:/home/exp.exactpro.com/oleg.legkov/Desktop/Corrector_3.0/lib/poi-3.9-20121203.jar Bugs

//javac -cp .;C:\tmp\Corrector_3.0\lib\poi-3.9-20121203.jar Bugs.java

import java.io.*;
import java.util.*;

import javax.xml.parsers.*;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.w3c.dom.*;

public class Bugs{
	private static final String HOME = "/home/exp.exactpro.com/oleg.legkov/Desktop/Corrector_3.0/";
	//private static final String HOME = "C:\\tmp\\Corrector_3.0\\";
	private static FileWriter log;
	
	public static void main(String[] args) throws Exception{
		try{
			log = new FileWriter("createXMLDB_3.0.log");
		}
		catch(FileNotFoundException e){
			e.printStackTrace();
		}
		catch(IOException e){
			e.printStackTrace();
		}
	
		Bugs b = new Bugs();
		System.out.println("createXMLDB from information_3.0.xml");
		b.createDB("MessagesDB_3.0.xml", "Information_3.0.xls");
		log.write("\ncreateXMLDB Func\n");
		System.out.println("Done");
		
		log.write("\nupdate bugs\n");
		System.out.println("Update Bugs from Bugs_3.0.xml");
		b.update("MessagesDB_3.0.xml", "Bugs_3.0.xls");
		System.out.println("Done");
	}

	public void createDB(String dbFILE, String FILE) throws Exception{	
		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document docXML = dBuilder.newDocument();
		
		FileInputStream fisInformation = new FileInputStream(HOME + FILE);
		HSSFWorkbook documentInformation = new HSSFWorkbook(fisInformation);
		HSSFSheet sheet = documentInformation.getSheet("messages2");
		int lastRow = sheet.getLastRowNum();
		int headerRow = 0;
		int messageNumber = 0;

		String[] tempArr = { "si_e", "si_c", "su_e", "su_c", "ssd_e", "ssd_c", "smd_e", "smd_c", "psd_e", "psd_c",
			"pmd_e", "pmd_c", "ai_e", "ai_c", "au_e", "au_c", "asdb_e", "asdb_c","amdb_e", "amdb_c", "asda_e",
			"asda_c", "amda_e", "amda_c" };

		LinkedHashSet<String> defaultSet = new LinkedHashSet<String>(Arrays.asList(tempArr));

		Element rootElement = docXML.createElement("apamessages");
		docXML.appendChild(rootElement);

		for(int i = 0; i < lastRow; i++){
			if(sheet.getRow(i).getCell(0).getStringCellValue().matches("") || sheet.getRow(i).getCell(0).getCellType() == Cell.CELL_TYPE_BLANK) {
				headerRow = 0;
				continue;
			}

			if(sheet.getRow(i).getCell(0).getStringCellValue().matches("TCR-Ack") ||
				sheet.getRow(i).getCell(0).getStringCellValue().matches("TCR-S") ||
				sheet.getRow(i).getCell(0).getStringCellValue().matches("TTC")){
					headerRow = i;
					continue;
			}

			String str = sheet.getRow(i).getCell(0).getStringCellValue();
			if(str.matches("Submit/Publication") || str.matches("Cancellation") || str.matches("Pre-Release") ||
				str.matches("Amend") || str.matches("TCR-C")) continue;

			Element message = docXML.createElement("message");
			rootElement.appendChild(message);
			Attr nameAttr = docXML.createAttribute("name");
			nameAttr.setValue(str);
			message.setAttributeNode(nameAttr);
			Attr numberAttr = docXML.createAttribute("number");
			numberAttr.setValue(Integer.toString(messageNumber++));
			message.setAttributeNode(numberAttr);

			int lastCell = sheet.getRow(i).getLastCellNum();
			for(int j = 1; j < lastCell; j++){
				String key = sheet.getRow(headerRow).getCell(j).getStringCellValue();
				String defaultValue = "";
				if(key.matches("SecurityIDSource")) key = "Instrument";
				if(key.matches("SecurityID") || key.matches("CountryOfIssue") || key.matches("")) continue;

				//SET DEFAULT
				if(key.matches("AssistedReportAPA") && defaultSet.contains(sheet.getRow(i).getCell(0).getStringCellValue())) defaultValue = "2";
				if(key.matches("QtyType") && defaultSet.contains(sheet.getRow(i).getCell(0).getStringCellValue())) defaultValue = "0";
				if(key.matches("PriceType") && defaultSet.contains(sheet.getRow(i).getCell(0).getStringCellValue())) defaultValue = "2";
				if(key.matches("OnExchangeInstr") && defaultSet.contains(sheet.getRow(i).getCell(0).getStringCellValue())) defaultValue = "0";
				if(key.matches("ClearingIntention") && defaultSet.contains(sheet.getRow(i).getCell(0).getStringCellValue())) defaultValue = "0";
				if(key.matches("PxQtyReviewed") && defaultSet.contains(sheet.getRow(i).getCell(0).getStringCellValue())) defaultValue = "N";
				if(key.matches("TrdType") && defaultSet.contains(sheet.getRow(i).getCell(0).getStringCellValue())) defaultValue = "0";
				if(key.matches("ExecMethod") && defaultSet.contains(sheet.getRow(i).getCell(0).getStringCellValue())) defaultValue = "0";
				if(key.matches("AlgorithmicTradeIndicator") && defaultSet.contains(sheet.getRow(i).getCell(0).getStringCellValue())) defaultValue = "0";
				if(key.matches("TradePublishIndicator") && defaultSet.contains(sheet.getRow(i).getCell(0).getStringCellValue())) defaultValue = "2";
				if(key.matches("PreviouslyReported") && defaultSet.contains(sheet.getRow(i).getCell(0).getStringCellValue())) defaultValue = "N";
				if(key.matches("ApplySupplementaryDeferral") && defaultSet.contains(sheet.getRow(i).getCell(0).getStringCellValue())) defaultValue = "Y";
				if(key.matches("TargetAPA") && defaultSet.contains(sheet.getRow(i).getCell(0).getStringCellValue())) defaultValue = "ECHO";

				String value = "";
				Cell tempCell =  sheet.getRow(i).getCell(j);
				int tempCellType = tempCell.getCellType();
				switch(tempCellType){
					case(Cell.CELL_TYPE_STRING): value = tempCell.getStringCellValue(); break;
					case(Cell.CELL_TYPE_NUMERIC):
						Double d = tempCell.getNumericCellValue();
						if(d % 1 == 0){
							Integer tmp = (int)tempCell.getNumericCellValue();
							value = Integer.toString(tmp);
						}
						else value = Double.toString(tempCell.getNumericCellValue());
						break;
				}

				Element nodeElement = docXML.createElement("tag");
				message.appendChild(nodeElement);
				Attr tagAttr = docXML.createAttribute("name");
				tagAttr.setValue(key);
				nodeElement.setAttributeNode(tagAttr);

				Element originalVal = docXML.createElement("original");
				originalVal.appendChild(docXML.createTextNode(value));
				nodeElement.appendChild(originalVal);
				Element defaultVal = docXML.createElement("default");
				defaultVal.appendChild(docXML.createTextNode(defaultValue));
				nodeElement.appendChild(defaultVal);
				Element bugVal = docXML.createElement("bug");
				bugVal.appendChild(docXML.createTextNode(""));
				nodeElement.appendChild(bugVal);

				Element crossBug = docXML.createElement("crossBug");
				crossBug.appendChild(docXML.createTextNode(""));
				nodeElement.appendChild(crossBug);

				Element versionVal = docXML.createElement("version");
				versionVal.appendChild(docXML.createTextNode("3.0.0"));
				nodeElement.appendChild(versionVal);
			}
		}

		TransformerFactory transformerFactory = TransformerFactory.newInstance();
		Transformer transformer = transformerFactory.newTransformer();
		transformer.setOutputProperty(OutputKeys.INDENT, "yes");
		transformer.setOutputProperty(OutputKeys.METHOD, "xml");
		docXML.getDocumentElement().normalize();
		DOMSource source = new DOMSource(docXML);

		StreamResult result = new StreamResult(new File(HOME + dbFILE));
		transformer.transform(source, result);
	}

	public void update(String dbFILE, String FILE) throws Exception{
		File inputFile = new File(HOME + dbFILE);

		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document docXML = dBuilder.parse(inputFile);
		docXML.getDocumentElement().normalize();

		FileInputStream bugsStream = new FileInputStream(HOME + FILE);
		HSSFWorkbook doc = new HSSFWorkbook(bugsStream);
		HSSFSheet sheet = doc.getSheetAt(0);

		int lastRow = sheet.getLastRowNum();

		for(int i = 0; i <= lastRow; i++){
			Row row = sheet.getRow(i);
			//System.out.println(i);
			String message = row.getCell(0).getStringCellValue();
			String tag = row.getCell(1).getStringCellValue();
			String bugValue = row.getCell(2).getStringCellValue();
			String bugDesc;

			Cell bugDescCell = row.getCell(3);
			if(bugDescCell == null)
				bugDesc = "NOT FOUND";
			else 
				bugDesc = bugDescCell.getStringCellValue();

			int messageNum = -1;
			int tagNumber = -1;

			switch(message){
				case("si_a"): messageNum = 0; break;
				case("su_a"): messageNum = 1; break;
				case("ssd_a"): messageNum = 2; break;
				case("smd_a"): messageNum = 3; break;
				case("ci_a"): messageNum = 4; break;
				case("cu_a"): messageNum = 5; break;
				case("csdb_a"): messageNum = 6; break;
				case("cmdb_a"): messageNum = 7; break;
				case("csda_a"): messageNum = 8; break;
				case("cmda_a"): messageNum = 9; break;
				case("pri_a"): messageNum = 10; break;
				case("pru_a"): messageNum = 11; break;
				case("prsda_a"): messageNum = 12; break;
				case("prmda_a"): messageNum = 13; break;
				case("prsd_a"): messageNum = 14; break;
				case("prmd_a"): messageNum = 15; break;
				case("ai_a"): messageNum = 16; break;
				case("au_a"): messageNum = 17; break;
				case("asdb_a"): messageNum = 18; break;
				case("amdb_a"): messageNum = 19; break;
				case("asda_a"): messageNum = 20; break;
				case("amda_a"): messageNum = 21; break;
				case("si_e"): messageNum = 22; break;
				case("si_c"): messageNum = 23; break;
				case("su_e"): messageNum = 24; break;
				case("su_c"): messageNum = 25; break;
				case("ssd_e"): messageNum = 26; break;
				case("ssd_c"): messageNum = 27; break;
				case("smd_e"): messageNum = 28; break;
				case("smd_c"): messageNum = 29; break;
				case("psd_e"): messageNum = 30; break;
				case("psd_c"): messageNum = 31; break;
				case("pmd_e"): messageNum = 32; break;
				case("pmd_c"): messageNum = 33; break;
				case("ci_e"): messageNum = 34; break;
				case("ci_c"): messageNum = 35; break;
				case("cu_e"): messageNum = 36; break;
				case("cu_c"): messageNum = 37; break;
				case("csdb_e"): messageNum = 38; break;
				case("csdb_c"): messageNum = 39; break;
				case("cmdb_e"): messageNum = 40; break;
				case("cmdb_c"): messageNum = 41; break;
				case("csda_e"): messageNum = 42; break;
				case("csda_c"): messageNum = 43; break;
				case("cmda_e"): messageNum = 44; break;
				case("cmda_c"): messageNum = 45; break;
				case("prsd_e"): messageNum = 46; break;
				case("prsd_c"): messageNum = 47; break;
				case("prmd_e"): messageNum = 48; break;
				case("prmd_c"): messageNum = 49; break;
				case("ai_e"): messageNum = 50; break;
				case("ai_c"): messageNum = 51; break;
				case("au_e"): messageNum = 52; break;
				case("au_c"): messageNum = 53; break;
				case("asdb_e"): messageNum = 54; break;
				case("asdb_c"): messageNum = 55; break;
				case("amdb_e"): messageNum = 56; break;
				case("amdb_c"): messageNum = 57; break;
				case("asda_e"): messageNum = 58; break;
				case("asda_c"): messageNum = 59; break;
				case("amda_e"): messageNum = 60; break;
				case("amda_c"): messageNum = 61; break;
				case("si_dss_a"): messageNum = 62; break;
				case("si_dss_e"): messageNum = 63; break;
				case("si_dss_c"): messageNum = 64; break;
				case("si_gtp"): messageNum = 65; break;
				case("su_dss_a"): messageNum = 66; break;
				case("su_dss_e"): messageNum = 67; break;
				case("su_dss_c"): messageNum = 68; break;
				case("ssd_dss_a"): messageNum = 69; break;
				case("ssd_dss_e"): messageNum = 70; break;
				case("ssd_dss_c"): messageNum = 71; break;
				case("psd_dss_e"): messageNum = 72; break;
				case("psd_dss_c"): messageNum = 73; break;
				case("psd_gtp"): messageNum = 74; break;
				case("smd_dss_a"): messageNum = 75; break;
				case("smd_dss_e"): messageNum = 76; break;
				case("smd_dss_c"): messageNum = 77; break;
				case("pmd_dss_e"): messageNum = 78; break;
				case("pmd_dss_c"): messageNum = 79; break;
				case("pmd_gtp"): messageNum = 80; break;
				case("ci_dss_a"): messageNum = 81; break;
				case("ci_dss_e"): messageNum = 82; break;
				case("ci_dss_c"): messageNum = 83; break;
				case("ci_gtp"): messageNum = 84; break;
				case("cu_dss_a"): messageNum = 85; break;
				case("cu_dss_e"): messageNum = 86; break;
				case("cu_dss_c"): messageNum = 87; break;
				case("csdb_dss_a"): messageNum = 88; break;
				case("csdb_dss_e"): messageNum = 89; break;
				case("csdb_dss_c"): messageNum = 90; break;
				case("cmdb_dss_a"): messageNum = 91; break;
				case("cmdb_dss_e"): messageNum = 92; break;
				case("cmdb_dss_c"): messageNum = 93; break;
				case("csda_dss_a"): messageNum = 94; break;
				case("csda_dss_e"): messageNum = 95; break;
				case("csda_dss_c"): messageNum = 96; break;
				case("csda_gtp"): messageNum = 97; break;
				case("cmda_dss_a"): messageNum = 98; break;
				case("cmda_dss_e"): messageNum = 99; break;
				case("cmda_dss_c"): messageNum = 100; break;
				case("cmda_gtp"): messageNum = 101; break;
				case("pri_dss_a"): messageNum = 102; break;
				case("pru_dss_a"): messageNum = 103; break;
				case("prsda_dss_a"): messageNum = 104; break;
				case("prmda_dss_a"): messageNum = 105; break;
				case("prsd_dss_a"): messageNum = 106; break;
				case("prsd_dss_e"): messageNum = 107; break;
				case("prsd_dss_c"): messageNum = 108; break;
				case("prsd_gtp"): messageNum = 109; break;
				case("prmd_dss_a"): messageNum = 110; break;
				case("prmd_dss_e"): messageNum = 111; break;
				case("prmd_dss_c"): messageNum = 112; break;
				case("prmd_gtp"): messageNum = 113; break;
				case("ai_dss_a"): messageNum = 114; break;
				case("ai_dss_e"): messageNum = 115; break;
				case("ai_dss_c"): messageNum = 116; break;
				case("ai_gtp"): messageNum = 117; break;
				case("au_dss_a"): messageNum = 118; break;
				case("au_dss_e"): messageNum = 119; break;
				case("au_dss_c"): messageNum = 120; break;
				case("asdb_dss_a"): messageNum = 121; break;
				case("asdb_dss_e"): messageNum = 122; break;
				case("asdb_dss_c"): messageNum = 123; break;
				case("amdb_dss_a"): messageNum = 124; break;
				case("amdb_dss_e"): messageNum = 125; break;
				case("amdb_dss_c"): messageNum = 126; break;
				case("asda_dss_a"): messageNum = 127; break;
				case("asda_dss_e"): messageNum = 128; break;
				case("asda_dss_c"): messageNum = 129; break;
				case("asda_gtp"): messageNum = 130; break;
				case("amda_dss_a"): messageNum = 131; break;
				case("amda_dss_e"): messageNum = 132; break;
				case("amda_dss_c"): messageNum = 133; break;
				case("amda_gtp"): messageNum = 134; break;
			}

			//System.out.println(message);

			if(messageNum > 61){
				switch(tag){
					case("BOATSequenceNumber"): tagNumber = 0; break;
					case("RoutingSeq"): tagNumber = 1; break;
					case("Origin"): tagNumber = 2; break;
					case("SeqNo"): tagNumber = 3; break;
					case("ContractMultiplier"): tagNumber = 4; break;
					case("ExecutedPrice"): tagNumber = 5; break;
					case("ExecutedPriceDivisor"): tagNumber = 6; break;
					case("PriceNotation"): tagNumber = 7; break;
					case("ExecutedSize"): tagNumber = 8; break;
					case("ExecutedSizeDivisor"): tagNumber = 9; break;
					case("NotationOfTheQuantityInMeasurementUnit"): tagNumber = 10; break;
					case("QuantityInMeasurementUnit"): tagNumber = 11; break;
					case("QuantityInMeasurementUnitDivisor"): tagNumber = 12; break;
					case("NotionalAmount"): tagNumber = 13; break;
					case("NotionalAmountDivisor"): tagNumber = 14; break;
					case("NotionalCurrency"): tagNumber = 15; break;
					case("GrossTradeAmt"): tagNumber = 16; break;
					case("ReferencePrice"): tagNumber = 17; break;
					case("ReferencePriceDivisor"): tagNumber = 18; break;
					case("StrikePrice"): tagNumber = 19; break;
					case("Yield"): tagNumber = 20; break;
					case("EventType"): tagNumber = 21; break;
					case("FixTradeReportType"): tagNumber = 22; break;
					case("VenueType"): tagNumber = 23; break;
					case("MultilegReportingType"): tagNumber = 24; break;
					case("NoUnderlyingsTCR"): tagNumber = 25; break;
					case("RejectCode"): tagNumber = 26; break;
					case("RejectReasonText"): tagNumber = 27; break;
					case("Capacity"): tagNumber = 28; break;
					case("OriginalCapacity"): tagNumber = 29; break;
					case("ContraCapacity"): tagNumber = 30; break;
					case("IsAggressor"): tagNumber = 31; break;
					case("LSEGClearingAccountType"): tagNumber = 32; break;
					case("LSEGClearingType"): tagNumber = 33; break;
					case("ClearingInstructions"): tagNumber = 34; break;
					case("TransactionToBeCleared"): tagNumber = 35; break;
					case("MatchStatus"): tagNumber = 36; break;
					case("OffBookStatus"): tagNumber = 37; break;
					case("OptionType"): tagNumber = 38; break;
					case("PublicationPending"): tagNumber = 39; break;
					case("PublishIndicator"): tagNumber = 40; break;
					case("SecurityType"): tagNumber = 41; break;
					case("Side"): tagNumber = 42; break;
					case("TradeReportTransType"): tagNumber = 43; break;
					case("TradeReportingModel"): tagNumber = 44; break;
					case("AgreedTime"): tagNumber = 45; break;
					case("BookDefinitionID"): tagNumber = 46; break;
					case("CFICode"): tagNumber = 47; break;
					case("LSEGContraClearingMember"): tagNumber = 48; break;
					case("ClientID"): tagNumber = 49; break;
					case("ClientOrderID"): tagNumber = 50; break;
					case("ContraClientID"): tagNumber = 51; break;
					case("ContraFirmPartyID"): tagNumber = 52; break;
					case("ContraOwnerPartyID"): tagNumber = 53; break;
					case("ContraTraderPartyID"): tagNumber = 54; break;
					case("Currency"): tagNumber = 55; break;
					case("EnteringFirmPartyID"): tagNumber = 56; break;
					case("ExecutingTraderPartyID"): tagNumber = 57; break;
					case("EventLinkID"): tagNumber = 58; break;
					case("ExecType"): tagNumber = 59; break;
					case("ExecutingFirm"): tagNumber = 60; break;
					case("ExpirationDate"): tagNumber = 61; break;
					case("FirmTradeID"): tagNumber = 62; break;
					case("InstrumentID"): tagNumber = 63; break;
					case("InstrumentSource"): tagNumber = 64; break;
					case("InstrumentCurrency"): tagNumber = 65; break;
					case("IntendedPublishTime"): tagNumber = 66; break;
					case("LSEGClearingMember"): tagNumber = 67; break;
					case("OrgTrdMatchID"): tagNumber = 68; break;
					case("OwnerID"): tagNumber = 69; break;
					case("Segment"): tagNumber = 70; break;
					case("ReportedTime"): tagNumber = 71; break;
					case("SettlementDate"): tagNumber = 72; break;
					case("ISIN"): tagNumber = 73; break;
					case("Symbol"): tagNumber = 74; break;
					case("TradeMatchID"): tagNumber = 75; break;
					case("TradeReportID"): tagNumber = 76; break;
					case("TradeReportRefID"): tagNumber = 77; break;
					case("Underlying"): tagNumber = 78; break;
					case("VenueOfExecution"): tagNumber = 79; break;
					case("VenueOfPublication"): tagNumber = 80; break;
					case("MarketSource"): tagNumber = 81; break;
					case("TradeReportLinkID"): tagNumber = 82; break;
					case("LateTradeIndicator"): tagNumber = 83; break;
					case("LateCancellation"): tagNumber = 84; break;
					case("EmissionAllowanceType"): tagNumber = 85; break;
					case("PriceQuantityReviewed"): tagNumber = 86; break;
					case("InstrumentStatus"): tagNumber = 87; break;
					case("MarketID"): tagNumber = 88; break;
					case("ShortName"): tagNumber = 89; break;
					case("ADT"): tagNumber = 90; break;
					case("ClosingPrice"): tagNumber = 91; break;
					case("ParValue"): tagNumber = 92; break;
					case("ParValueCurrency"): tagNumber = 93; break;
					case("NormalMarketSize"): tagNumber = 94; break;
					case("DelegatedReport"): tagNumber = 95; break;
					case("PackageTradeID"): tagNumber = 96; break;
					case("OnBookTradingMode"): tagNumber = 97; break;
					case("PendingPrice"): tagNumber = 98; break;
					case("SecondaryPublication"): tagNumber = 99; break;
					case("Flags"): tagNumber = 100; break;
					case("TradeDetails"): tagNumber = 101; break;
					case("ReservedBitMask"): tagNumber = 102; break;
					case("ReservedString1"): tagNumber = 103; break;
					case("ReservedString2"): tagNumber = 104; break;
					case("TransactTime"): tagNumber = 105; break;
					case("Equity"): tagNumber = 106; break;
					case("UnitQuantity"): tagNumber = 107; break;
					case("UnitQuantityDivisor"): tagNumber = 108; break;
					case("InstrumentUniverse"): tagNumber = 109; break;
				}
			}

			if(messageNum < 62){
				switch(tag){
					case("TradeID"): tagNumber = 0; break;
					case("TradeReportID"): tagNumber = 1; break;
					case("FirmTradeID"): tagNumber = 2; break;
					case("OrigTradeID"): tagNumber = 3; break;
					case("AggPublicationID"): tagNumber = 4; break;
					case("AssistedReportAPA"): tagNumber = 5; break;
					case("Instrument"): tagNumber = 6; break;
					case("Currency"): tagNumber = 7; break;
					case("ExecType"): tagNumber = 8; break;
					case("LastQty"): tagNumber = 9; break;
					case("QtyType"): tagNumber = 10; break;
					case("ContractMultiplier"): tagNumber = 11; break;
					case("LastPx"): tagNumber = 12; break;
					case("PriceType"): tagNumber = 13; break;
					case("NotionalAmount"): tagNumber = 14; break;
					case("NotionalCurrency"): tagNumber = 15; break;
					case("UnitOfMeasure"): tagNumber = 16; break;
					case("TimeUnit"): tagNumber = 17; break;
					case("UnitOfMeasureQty"): tagNumber = 18; break;
					case("EmissionAllowanceType"): tagNumber = 19; break;
					case("Issuer"): tagNumber = 20; break;
					case("TransactTime"): tagNumber = 21; break;
					case("SettlDate"): tagNumber = 22; break;
					case("OnExchangeInstr"): tagNumber = 23; break;
					case("ClearingIntention"): tagNumber = 24; break;
					case("Text"): tagNumber = 25; break;
					case("PxQtyReviewed"): tagNumber = 26; break;
					case("TradeReportSystem"): tagNumber = 27; break;
					case("DelayToTime"): tagNumber = 28; break;
					case("RptTime"): tagNumber = 29; break;
					case("PackageID"): tagNumber = 30; break;
					case("TradeNumber"): tagNumber = 31; break;
					case("TotNumTradeReports"): tagNumber = 32; break;
					case("VenueType"): tagNumber = 33; break;
					case("MatchType"): tagNumber = 34; break;
					case("TrdType"): tagNumber = 35; break;
					case("OrderCategory"): tagNumber = 36; break;
					case("TrdSubType"): tagNumber = 37; break;
					case("TradeReportTransType"): tagNumber = 38; break;
					case("SecondaryTrdType"): tagNumber = 39; break;
					case("ExecMethod"): tagNumber = 40; break;
					case("AlgorithmicTradeIndicator"): tagNumber = 41; break;
					case("TradePublishIndicator"): tagNumber = 42; break;
					case("RegulatoryReportType"): tagNumber = 43; break;
					case("PreviouslyReported"): tagNumber = 44; break;
					case("NoTradePriceConditions"): tagNumber = 45; break;
					case("NoTrdRegPublications"): tagNumber = 46; break;
					case("NoSides"): tagNumber = 47; break;
				}
			}

			NodeList nList1 = docXML.getElementsByTagName("message");
			Element eElement1 = (Element) nList1.item(messageNum);
			NodeList nList2 = eElement1.getElementsByTagName("tag");
			Element eElement2 = (Element) nList2.item(tagNumber);

			if(bugDescCell == null){
				eElement2.getElementsByTagName("bug").item(0).setTextContent(bugValue);
				eElement2.getElementsByTagName("crossBug").item(0).setTextContent("");
			}
			else{
				eElement2.getElementsByTagName("bug").item(0).setTextContent("");
				eElement2.getElementsByTagName("crossBug").item(0).setTextContent(bugValue);
			}
		}

		TransformerFactory transformerFactory = TransformerFactory.newInstance();
		Transformer transformer = transformerFactory.newTransformer();
		transformer.setOutputProperty(OutputKeys.INDENT, "yes");
		transformer.setOutputProperty(OutputKeys.METHOD, "xml");
		docXML.getDocumentElement().normalize();
		DOMSource source = new DOMSource(docXML);
		StreamResult result = new StreamResult(new File(HOME + dbFILE));
		transformer.transform(source, result);
	}
}
