package decken;

import java.net.HttpURLConnection;
import java.net.URL;                                                                                                                                                                                                                                                                                                                                                                                                                                                              
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.GZIPInputStream;
import java.io.*;                                                                                                                                                                                                                                                                                                                                                                                                                                                                 

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import com.gargoylesoftware.htmlunit.BrowserVersion;                                                                                                                                                                                                                                                                                                                                                                                                                              
import com.gargoylesoftware.htmlunit.WebClient;                                                                                                                                                                                                                                                                                                                                                                                                                                   
import com.gargoylesoftware.htmlunit.html.*;                                                                                                                                                                                                                                                                                                                                                                                                                                      
import com.gargoylesoftware.htmlunit.*; 

public class fetchFile {
	String createDirPath = "D:/Java/stockPrice";
	// Referer ºô§}
	String[] tempUrl = {                                                                                                                                                                                                                                                                                                                                                                                                                                                            
		"http://www.google.com.tw/",                                                                                                                                                                                                                                                                                                                                                                                                                                              
		"http://tw.search.yahoo.com/search?p=hinet&fr=yfp&ei=utf-8&v=0",                                                                                                                                                                                                                                                                                                                                                                                                                                                       
		"http://news.google.com.tw/news?ned=tw&hl=zh-TW",
		"http://www.bookzone.com.tw/",
		"http://dxmonline.com/html/",
		"http://www.javaworld.com.tw/jute/",
		"http://udn.com/NEWS/main.html",
		"http://news.google.com.tw/nwshp?hl=zh-TW&tab=wn",
		"http://www.graphviz.org/",                                                                                                                                                                                                                                                                                                                                                                                                                                               
		"http://noder.tw/"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
	};
	
	int stockCount = 0;
	String[][] stockList = new String[5000][2];
	
	public void downloadStockPriceFile() throws Exception{
		String mon, day;
		
		WebClient webClient = new WebClient(BrowserVersion.INTERNET_EXPLORER_6);                                                                                                                                                                                                                                                                                                                                                                                                    
		webClient.setJavaScriptEnabled(true);
		webClient.addRequestHeader("Host", "http://etds.ncl.edu.tw");
		webClient.addRequestHeader("Referer", tempUrl[(int)(Math.random()*100)%tempUrl.length]);
		webClient.addRequestHeader("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8");
		webClient.addRequestHeader("Accept-Language", "zh-tw,en-us;q=0.7,en;q=0.3");
		webClient.addRequestHeader("Accept-Charset", "Big5,utf-8;q=0.7,*;q=0.7");
		webClient.addRequestHeader("Connection", "keep-alive");
		webClient.addRequestHeader("Accept-Encoding", "gzip,deflate");
		
		for(int i=1; i<=3; i++) {
			for(int j=1; j<=31; j++) {
				if(i/10 != 0)
					mon = Integer.toString(i);
				else
					mon = "0" + Integer.toString(i);
				
				if(j/10 != 0)
					day = Integer.toString(j);
				else
					day = "0" + Integer.toString(j);

				try {
					String stockUrl = "http://www.twse.com.tw/ch/trading/exchange/MI_INDEX/MI_INDEX3_print.php?genpage=genpage/Report2012" + mon + "/A1122012" + mon + day + "ALL_1.php&type=csv";
					URL url = new URL(stockUrl);
					
					//webClient = new WebClient(BrowserVersion.INTERNET_EXPLORER_6);
					BufferedInputStream bufferedInputStream =
						new BufferedInputStream(new DataInputStream(webClient.getPage(url).getWebResponse().getContentAsStream()));					
					

					BufferedOutputStream bufferedOutputStream = 
						new BufferedOutputStream(new FileOutputStream(createDirPath + "/A1122012" + mon + day + "ALL_1.csv"));
					byte[] data = new byte[1];                                                                                                                                                                                                                                                                                                                                                                                                                                              
					while(bufferedInputStream.read(data) != -1) {                                                                                                                                                                                                                                                                                                                                                                                                                           
						bufferedOutputStream.write(data);                                                                                                                                                                                                                                                                                                                                                                                                                                     
					}                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
					bufferedOutputStream.flush();                                                                                                                                                                                                                                                                                                                                                                                                                                           
					bufferedInputStream.close();
					bufferedOutputStream.close();
					
					//Thread.currentThread().sleep(((long)(Math.random()*10)+1)*100); 
				}
				catch(Exception e){                                                                                                                                                                                                                                                                                                                                                                                                                                                         
					e.printStackTrace();                                                                                                                                                                                                                                                                                                                                                                                                                                                      
				}   
			}
		}
	}
	
	public void deleteNullFile() throws Exception{
		String basePath = "D:/Java/stockPrice/";
		File dirPath = new File(basePath);
		String[] files = dirPath.list();
		
		for (int index = 0; index < files.length; index++) {
			File file = new File(basePath + files[index]);
			if(file.length() < 1000)
				file.delete();
		}
	}
	
	private String getFileName(String csvFileName){
		return csvFileName.substring(0, csvFileName.length()-4)+".xls";
	}
	
	private List<String> format (String[] cells){
		List<String> cellList = new ArrayList<String>();
		StringBuffer buffer = new StringBuffer();
		for(String cell : cells){
			boolean isHead = cell.startsWith("\"");
			boolean isTail = cell.endsWith("\"");
			if(isHead){
				if(isTail){
					cellList.add(cell.substring(1, cell.length()-1));
				}else{
					buffer.append( cell.substring(1)).append(',');
				}
			}else{
				if(isTail){
					buffer.append(cell.substring(0, cell.length()-1));
					cellList.add(buffer. toString());
					buffer.delete(0, buffer.length());
				}else if(buffer.length()>0){
					buffer.append(cell).append(',');
				}else{
					cellList.add(cell);
				}
			}
		}
		return cellList; 
	}
	
	public void transferCSVtoXLS() throws Exception{
		String basePath = "D:/Java/stockPrice/";
		File dirPath = new File(basePath);
		String[] cvsFiles = dirPath.list();
		
		for (int index = 0; index < cvsFiles.length; index++) {
			FileReader fileReader = new FileReader(basePath + cvsFiles[index]);   
			BufferedReader bufferedReader = new BufferedReader(fileReader);
			
			String xlsFileName = getFileName(cvsFiles[index].toString());
			File xls = new File("D:/Java/stockPrice/" + xlsFileName);
			WritableWorkbook workbook = Workbook.createWorkbook(xls);
			WritableSheet sheet = workbook.createSheet("XLS", 0);
			int rowNumber = 0;
			String row = bufferedReader.readLine();
			
			while(row != null) {    
				String[] cells = row.split(",");    
				List<String> cellList = format(cells);
				
				for(int i=0,size=cellList.size();i<size;i++){   
					sheet.addCell(new Label(i,rowNumber,cellList.get(i)));
				}
				
				rowNumber++;
				row = bufferedReader.readLine();
			}
			
			workbook.write();
			workbook.close();
			
			File cvsFile = new File("D:/Java/stockPrice/" + cvsFiles[index]);
			cvsFile.delete();
		}
		//return xls;
	}
	
	public void getStockList() throws Exception{
		String basePath = "D:/Java/stockPrice";
		File dirPath = new File(basePath);
		
		String[] files = dirPath.list();
		
		try{
			int index = files.length-1;
			InputStream is = new FileInputStream(basePath + "/" + files[index]);
			jxl.Workbook rwb = Workbook.getWorkbook(is);
			Sheet rs = rwb.getSheet(0);
			
			for (int row = 1, over = 0; row < rs.getRows() && over == 0; row++) {
				Cell c  = rs.getCell(0, row);
				String cContent = c.getContents();
				
				if(cContent.equals("1101")) {
					c  = rs.getCell(1, row);
					if(c.getContents().equals("¥xªd")) {
						stockList[stockCount][0] = cContent;
						stockList[stockCount++][1] = c.getContents();
						
						for (++row; row < rs.getRows() && over == 0; row++) {
							c  = rs.getCell(0, row);
							cContent = c.getContents();
							
							if(cContent.length() > 4)
								continue;
							
							stockList[stockCount][0] = cContent;
						
							try {
								Integer.parseInt(cContent);
							}
							catch(Exception e){
								over = 1;
								e.printStackTrace();
							}
							
							c  = rs.getCell(1, row);
							cContent = c.getContents();
							stockList[stockCount++][1] = cContent;
						}
					}
				}
			}//for (int row = 1, over = 0; row < rs.getRows() && over == 0; row++) {
		}
		catch(Exception e){
			e.printStackTrace();
		}
	}
	
	public void getPERPBR() throws Exception{
		//BufferedReader br = null;
		//OutputStream outputPersonStream = new FileOutputStream("D:/Java/PERPBR.txt");
		//Writer out = new OutputStreamWriter(outputPersonStream, "utf-8");
		String basePath = "D:/Java/stockPERPBR/";
		File dirPath = new File(basePath);
		String[] files = dirPath.list();
		
		int bufReader;
		int dateCount = 2;
		String in = "";
		
		getStockList();
		
		WritableWorkbook book = Workbook.createWorkbook(new File("D:/Java/parseStockPERPBR.xls"));
		WritableSheet sheet = book.createSheet("stockPERPBR" ,0);
				
		for (int i = 0; i < files.length; i++, dateCount++) {
			FileInputStream inputStream = new FileInputStream(basePath + files[i]); 
			BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(inputStream,"utf-8"));
			
			String fileName = basePath + "0" + files[i].substring(0, files[i].length()-4) + ".txt"; 
			
			OutputStream outputPersonStream2 = new FileOutputStream(fileName);
			Writer out2 = new OutputStreamWriter(outputPersonStream2, "utf-8");
			
			try {
				while((bufReader = bufferedReader.read()) != -1) {
					if(bufReader == '\r' || bufReader == '\n')
						;
					else if(bufReader == '<') {
						while((bufReader = bufferedReader.read()) != '>')
							;
						
						out2.write(" ");
					}
					else
						out2.write(bufReader);
				}
			} catch (Exception e) {
				bufferedReader.close();
				out2.close();
				e.printStackTrace();
			}
			bufferedReader.close();
			out2.close();
			
			inputStream = new FileInputStream(fileName); 
			bufferedReader = new BufferedReader(new InputStreamReader(inputStream,"utf-8"));
			
			outputPersonStream2 = new FileOutputStream(basePath + "20" + files[i].substring(0, files[i].length()-4) + ".txt");
			String date = files[i].substring(0, files[i].length()-4);
			System.out.print(date + '\t');
			out2 = new OutputStreamWriter(outputPersonStream2, "utf-8");
			
			String[][] stockPERPBR = new String[3000][2];	//[0]PER, [1]PBR
			String inTemp = "";
			
			int xlsCount = 1;
			Label label = null;
			
			System.out.println(dateCount);
			label = new Label(dateCount, 0, date);
			sheet.addCell(label);
			
			try {
				in = bufferedReader.readLine();
				in = in.replaceAll(" +", " ");
				
				in = in.substring(in.indexOf(stockList[0][1]), in.indexOf(stockList[stockCount-1][1])+100);
				
				for (int j = 0; j < stockCount; j++, xlsCount+=2) {
					try {
						inTemp = in.substring(in.indexOf(stockList[j][1])+stockList[j][1].length()+1, in.indexOf(stockList[j][1])+stockList[j][1].length()+1+50);					
					} catch (Exception e) {
						e.printStackTrace();
						stockPERPBR[j][0] = "none";
						stockPERPBR[j][1] = "none";
						continue;
					}

					String[] inTempSplit = inTemp.split(" ");
					stockPERPBR[j][0] = inTempSplit[0];	//PER
					stockPERPBR[j][1] = inTempSplit[2];	//PBR
					
					if(i==0) {
						label = new Label(0, xlsCount, stockList[j][0]);
						sheet.addCell(label);
						label = new Label(1, xlsCount, stockList[j][1]);
						sheet.addCell(label);
					}
					//System.out.println(i);
					label = new Label(dateCount, xlsCount, stockPERPBR[j][0]);
					sheet.addCell(label);
					label = new Label(dateCount, xlsCount+1, stockPERPBR[j][1]);
					sheet.addCell(label);
					
					out2.write(stockList[j][0] + " " + stockList[j][1] + " " + stockPERPBR[j][0] + " " + stockPERPBR[j][1] + "\r\n");
				}
			} catch (Exception e) {
				bufferedReader.close();
				out2.close();
				e.printStackTrace();
			}
			bufferedReader.close();
			out2.close();
						
			File file = new File(basePath + "0" + files[i].substring(0, files[i].length()-4) + ".txt");
			file.delete();
			file = new File(basePath + files[i].substring(0, files[i].length()-4) + ".txt");
			file.delete();
		}	

		book.write();
		book.close();
		//FileInputStream inputStream = new FileInputStream("D:/Java/PERPBR.txt"); 
		//BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(inputStream,"utf-8"));
		//OutputStream outputPersonStream2 = new FileOutputStream("D:/Java/temp.txt");
		//Writer out2 = new OutputStreamWriter(outputPersonStream2, "utf-8");
		//int bufReader;
		//String in = "";

		/*try {
			while((bufReader = bufferedReader.read()) != -1) {
				if(bufReader == '\r' || bufReader == '\n')
					;
				else if(bufReader == '<') {
					while((bufReader = bufferedReader.read()) != '>')
						;
					
					out2.write(" ");
				}
				else
					out2.write(bufReader);
			}
		} catch (Exception e) {
			bufferedReader.close();
			out2.close();
			e.printStackTrace();
		}
		bufferedReader.close();
		out2.close();*/
		
		/*inputStream = new FileInputStream("D:/Java/temp.txt"); 
		bufferedReader = new BufferedReader(new InputStreamReader(inputStream,"utf-8"));
		String[][] stockPERPBR = new String[3000][2];	//[0]PER, [1]PBR
		String inTemp = "";
		
		try {
			in = bufferedReader.readLine();
			in = in.replaceAll(" +", " ");
			
			in = in.substring(in.indexOf(stockList[0][1]), in.indexOf(stockList[stockCount-1][1])+100);
			
			for (int i = 0; i < stockCount; i++) {
				try {
					inTemp = in.substring(in.indexOf(stockList[i][1])+stockList[i][1].length()+1, in.indexOf(stockList[i][1])+stockList[i][1].length()+1+50);					
				} catch (Exception e) {
					e.printStackTrace();
					stockPERPBR[i][0] = "none";
					stockPERPBR[i][1] = "none";
					continue;
				}

				String[] inTempSplit = inTemp.split(" ");
				stockPERPBR[i][0] = inTempSplit[0];	//PER
				stockPERPBR[i][1] = inTempSplit[2];	//PBR
				System.out.println(stockList[i][1] + " " + stockPERPBR[i][0] + " " + stockPERPBR[i][1]);
			}
			//System.out.println(in);
		} catch (Exception e) {
			e.printStackTrace();
		}*/
		
		/*try {
			while((in = bufferedReader.readLine())!= null) {
				//in = bufferedReader.readLine();
				if(in.indexOf("¥xªd") >= 0) {
					inSplit = in.split(" +");
					
					for (int i = 0; i < inSplit.length; i++) {
						System.out.println(inSplit[i] + "_");	
					}
				}
			}			
		} catch (Exception e) {
			e.printStackTrace();
		}*/
		//int stockNumIdx, stockStrIdx;
		//boolean leftBracket = false, rightBracket = false;
	}
	
	public void deleteNullPERPBRFile() throws Exception{
		String basePath = "D:/Java/stockPERPBR/";
		File dirPath = new File(basePath);
		String[] files = dirPath.list();
		
		try {
			for (int i = 0; i < files.length; i++) {
				File file = new File(basePath + files[i]);
				if(file.length() < 100000)
					file.delete();
				//System.out.print(file.length() + "\r\n");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}		
			
		/*
		String basePath = "D:/Java/stockPrice/";
		File dirPath = new File(basePath);
		String[] files = dirPath.list();
		
		for (int index = 0; index < files.length; index++) {
			File file = new File(basePath + files[index]);
			if(file.length() < 1000)
				file.delete();
		}
		 */
	}
		
	public void getPagePERPBR() throws Exception{
		String mon, day;
		BufferedReader br = null;
		OutputStream outputPersonStream;
		Writer out = null;
		
		try{
			for (int i = 1; i <= 3; i++) {
				for (int j = 1; j <= 31; j++) {
					if(i/10 != 0)
						mon = Integer.toString(i);
					else
						mon = "0" + Integer.toString(i);
					
					if(j/10 != 0)
						day = Integer.toString(j);
					else
						day = "0" + Integer.toString(j);
					
					outputPersonStream = new FileOutputStream("D:/Java/stockPERPBR/12" + mon + day + ".txt");
					out = new OutputStreamWriter(outputPersonStream, "utf-8");
					
					//http://www.twse.com.tw/ch/trading/exchange/BWIBBU/BWIBBU_d.php?input_date=101/04/02
				    URL url = new URL( "http://www.twse.com.tw/ch/trading/exchange/BWIBBU/BWIBBU_d.php?input_date=101/" + mon + "/" + day);
				    HttpURLConnection connection = (HttpURLConnection)url.openConnection();
				    connection.setDoInput( true );
				    connection.setRequestMethod( "GET" );
				    connection.setRequestProperty( "Host", url.getHost() );
				    connection.setRequestProperty( "User-Agent", "Mozilla/5.0 (Windows; U; Windows NT 5.1) Gecko/20100316 Firefox/3.6.2" );
				    connection.setRequestProperty( "Accept", "*/*" );
				    connection.setRequestProperty( "Accept-Encoding", "gzip, deflate" );
				    connection.setRequestProperty( "Connection", "close" );
				    
				    if( "gzip".equalsIgnoreCase( connection.getContentEncoding() ) )
				        br = new BufferedReader( new InputStreamReader( new GZIPInputStream( connection.getInputStream() ), "Big5" ) );
				    else
				        br = new BufferedReader( new InputStreamReader( connection.getInputStream(), "Big5" ) );
				    String inputLine = "";
				    while( ( inputLine = br.readLine() ) != null ){
				    	out.write(inputLine);
				    }
				    br.close();
				    out.close();
				    br = null;
				}
			}
		}
		catch( Exception e ){
		    try{
		        if( br != null ){
		            br.close();
		        }
		        if(out != null) {
		        	out.close();
		        }
		    }
		    catch( IOException ioe ){}
		}
	}
}