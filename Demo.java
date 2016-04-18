package decken;

import java.io.*;
import java.net.*;

public class Demo {
	public static void main(String[] argv) throws Exception{
		excelParser exlPar = new excelParser();
		fetchFile fetFile = new fetchFile();
		
		/*fetFile.downloadStockPriceFile();
		fetFile.deleteNullFile();
		fetFile.transferCSVtoXLS();
		exlPar.parseStockPrice();
		
		fetFile.getPagePERPBR();
		fetFile.deleteNullPERPBRFile();
		fetFile.getPERPBR();*/
		
		exlPar.combinePricePERPBR();
	}
}
