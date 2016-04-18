package decken;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;

import jxl.*;
import jxl.write.*;

public class excelParser {
	public void parseStockPrice() throws Exception {
		String basePath = "D:/Java/stockPrice";
		File dirPath = new File(basePath);
		
		String[] files = dirPath.list();
		/*for (int index = files.length-1; index >=0; index--) {
			System.out.println(files[index]);
		}*/
		
		//String fileName = "D:/Java/A11220120402ALL_1.xls";
       	
		WritableWorkbook book = Workbook.createWorkbook(new File("D:/Java/parseStockInfo.xls"));
		WritableSheet sheet = book.createSheet("stockInfo" ,0);
		/*
		Label label = new Label(0,0,"test");
		sheet.addCell(label); 
		jxl.write.Number number = new jxl.write.Number(1,0,789.123);   
		sheet.addCell(number); 
		book.write();
		book.close();
		*/
   		//OutputStream outputStream = new FileOutputStream("D:/Java/parse.txt");
		//Writer out = new OutputStreamWriter(outputStream, "utf-8");

		int index, stockCount = 0, tableRow = 2, over = 0, dateCount = 2, k;
		int allIndex;
		String date = "", cContent;
		InputStream is;
		Sheet rs = null;
		Label label = null;
		Cell c = null;
		String[] stocks = new String[5000];
		
		//System.out.println(date);
		//String date = fileName.indexOf("ALL");
		try{
			//out.write("\t\t");
			index = files.length-1;
			is = new FileInputStream(basePath + "/" + files[index]);
			jxl.Workbook rwb = Workbook.getWorkbook(is);
			rs = rwb.getSheet(0);
			
			for (int row = 1; row < rs.getRows() && over == 0; row++) {
				c  = rs.getCell(0, row);
				cContent = c.getContents();
				
				if(cContent.equals("1101")) {
					c  = rs.getCell(1, row);
					if(c.getContents().equals("台泥")) {
						stocks[stockCount++] = cContent;
						
						label = new Label(0, 1, "1101");
						sheet.addCell(label);
						label = new Label(1, 1, "台泥");
						sheet.addCell(label);
						
						for (++row; row < rs.getRows() && over == 0; row++) {
							c  = rs.getCell(0, row);
							cContent = c.getContents();
							
							if(cContent.length() > 4)
								continue;
							
							stocks[stockCount++] = cContent;
							
							label = new Label(0, tableRow, cContent);
							sheet.addCell(label);
							
							try {
								Integer.parseInt(cContent);
							}
							catch(Exception e){
								over = 1;
								e.printStackTrace();
							}
							
							c  = rs.getCell(1, row);
							cContent = c.getContents();
							label = new Label(1, tableRow++, cContent);
							sheet.addCell(label);
						}
					}
				}//END_for (int row = 1; row < rs.getRows() && over == 0; row++) {
			}
			
			for (index = 0, over = 0, tableRow = 2; index < files.length; index++, tableRow = 2, over = 0) {
				//System.out.println(files[index]);
				
				is = new FileInputStream(basePath + "/" +files[index]);
				rwb = Workbook.getWorkbook(is);
				rs = rwb.getSheet(0);
				
				allIndex = files[index].indexOf("ALL");
				date = files[index].substring(allIndex - 6, allIndex);
				
				label = new jxl.write.Label(dateCount++, 0, date);
				sheet.addCell(label);
				//jxl.write.Number number = new jxl.write.Number(1,0,789.123);   
				//sheet.addCell(number);
				//book.write();				
				
				for (int row = 1; row < rs.getRows() && over == 0; row++) {
				//for (int row = 1; row < 100 && over == 0; row++) {
					//for (int column = 0; column < rs.getColumns() && over == 0; column++) {
						c  = rs.getCell(0, row);
						cContent = c.getContents();
						
						if(cContent.equals("1101")) {
							//System.out.println(cContent);
							c  = rs.getCell(1, row);
							//System.out.println(c.getContents());
							if(c.getContents().equals("台泥")) {
								//System.out.println(cContent + '\t' + c.getContents() + '\t');
								//out.write(cContent + '\t' + "台泥" + '\t');
								//label = new Label(0, 1, "1101");
								//sheet.addCell(label);
								//book.write();
								//label = new Label(1, 1, "台泥");
								//sheet.addCell(label);
								//book.write();
								//jxl.write.Number number = new jxl.write.Number(1,0,789.123);   
								//sheet.addCell(number);
								//for (++column; column < rs.getColumns(); column++) {
								c  = rs.getCell(8, row);
								cContent = c.getContents();
								label = new Label(dateCount-1, 1, cContent);
								sheet.addCell(label);
									//book.write();
									//out.write(cContent + "\r\n");
									

								//}
								
								//out.write("\r\n");
								
								for (++row; row < rs.getRows() && over == 0; row++) {
								//for (++row; row < 6500 && over == 0; row++) {
									//for (column = 0; column < rs.getColumns(); column++) {
										//System.out.print(row + "\t");
										c  = rs.getCell(0, row);
										cContent = c.getContents();
										//System.out.print(cContent + "\t");
										
										if(cContent.length() > 4)
											continue;
										
										for(k=0; k < stockCount; k++) {
											if(stocks[k].equals(cContent))
												break;
										}
										
										if(k == stockCount)
											continue;
										//label = new Label(0, k+1, cContent);
										//sheet.addCell(label);
										//book.write();
										
										try {
											Integer.parseInt(cContent);
										}
										catch(Exception e){
											over = 1;
											e.printStackTrace();
										}
										
										/*c  = rs.getCell(1, row);
										cContent = c.getContents();
										System.out.print(cContent + "\t");
										label = new Label(1, tableRow, cContent);
										sheet.addCell(label);*/
										//book.write();
										
										c  = rs.getCell(8, row);
										cContent = c.getContents();
										label = new Label(dateCount-1, k+1, cContent);
										sheet.addCell(label);
										//book.write();
									//}
									//out.write("\r\n");
								}
								//over = 1;
							}
							else
								continue;
						}
						else
							continue;
					//}
					//out.write("\r\n");
				}
			}//END_for (index = 0, over = 0, tableRow = 2; index < files.length; index++, tableRow = 2, over = 0) {			

			book.write();
			book.close();
			//out.write("\t\t" + date + "\r\n");
			//out.close();
		}
		catch(Exception e){
			e.printStackTrace();
		}
	}
	public void combinePricePERPBR() throws Exception {
		//String basePath = "D:/Java/stockPrice";
		InputStream is, is2;
		Sheet rs, rs2;
		jxl.Workbook rwb, rwb2;
		Cell c = null;
		String cContent;
		Label label;
		
		OutputStream outputStream, outputStream2;
		Writer out1, out2;
		
		WritableWorkbook book = Workbook.createWorkbook(new File("D:/Java/stockPricePERPBR.xls"));
		WritableSheet sheet = book.createSheet("stockInfo" ,0);
		
		is = new FileInputStream("D:/Java/parseStockInfo.xls");
		rwb = Workbook.getWorkbook(is);
		rs = rwb.getSheet(0);
		
		is2 = new FileInputStream("D:/Java/parseStockPERPBR.xls");
		rwb2 = Workbook.getWorkbook(is2);
		rs2 = rwb2.getSheet(0);
		
		for (int i = 1; i < rs.getColumns(); i++) {
			c  = rs.getCell(i, 0);
			cContent = c.getContents();
			
			label = new Label(i, 0, cContent);
			sheet.addCell(label);
		}
		
		for (int i = 1, j = 1; i < rs.getRows(); i++) {
			c  = rs.getCell(0, i);
			cContent = c.getContents();
			
			label = new Label(0, j, cContent);
			sheet.addCell(label);
			
			c  = rs.getCell(1, i);
			cContent = c.getContents();
			
			label = new Label(1, j, cContent);
			sheet.addCell(label);
			
			j = i*3+1;
		}
		
		jxl.write.Number number;
		
		for (int i = 1, j = 1; i < rs.getRows(); i++, j+=2) {
			for (int k = 2; k < rs.getColumns(); k++) {
				try {
					c  = rs.getCell(k, i);
					cContent = c.getContents();
					
					//System.out.println(.valueOf(cContent).intValue());
					
					//
					number = new jxl.write.Number(k, 1+(i-1)*3, Double.parseDouble(cContent));   
					sheet.addCell(number); 
					//
					
					//label = new Label(k, 1+(i-1)*3, cContent);
					//sheet.addCell(label);
				} catch (Exception e) {
					e.printStackTrace();
				}
					
				try {
					c  = rs2.getCell(k, j);
					cContent = c.getContents();
					
					//
					number = new jxl.write.Number(k, 2+(i-1)*3, Double.parseDouble(cContent));   
					sheet.addCell(number); 
					//
					
					//label = new Label(k, 2+(i-1)*3, cContent);
					//sheet.addCell(label);
				} catch (Exception e) {
					e.printStackTrace();
				}

				try {
					c  = rs2.getCell(k, j+1);
					cContent = c.getContents();
					
					//
					number = new jxl.write.Number(k, 3+(i-1)*3, Double.parseDouble(cContent));   
					sheet.addCell(number); 
					//
					
					//label = new Label(k, 3+(i-1)*3, cContent);
					//sheet.addCell(label);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}
		rwb.close();
		rwb2.close();
		is.close();
		is2.close();
		
		book.write();
		book.close();
		
		String stockName;
		
		is = new FileInputStream("D:/Java/stockPricePERPBR2.xls");
		rwb = Workbook.getWorkbook(is);
		rs = rwb.getSheet(0);
		int j = 1;

		c = rs.getCell(0, 1);
		stockName = c.getContents();
		
		while(stockName != "") {
			try {
				OutputStream outputXLSStream = new FileOutputStream("D:/Java/Test/stockPricePERPBR.txt");
				Writer out = new OutputStreamWriter(outputXLSStream, "utf-8");
							
				int k = 1;
				
				//System.out.println(stockName);
				
				//for (int i = 1; i <= 3; i++) {
				for (int i = 2; i <= rs.getColumns(); i++, k++) {
					for (int p = 1; p <= 3; p++) {
						try {
							c = rs.getCell(i, j+p-1);
							cContent = c.getContents();
							
							if(p==1)
								out.write(Integer.toString(k) + '\t');
							
							if(p!=3)
								out.write(cContent + '\t');
							else
								out.write(cContent);
						} catch (Exception e) {
							e.printStackTrace();
						}
					}
					out.write("\r\n");
				}
				out.close();
				//rwb.close();
				//is.close();
				
				/*String temp = "plot \"stockPricePERPBR.txt\" using 1:2 title \"stockPrice\" with linespoint, \"stockPricePERPBR.txt\" using 1:3 axes x1y2 title \"stockPER\" with linespoint, \"stockPricePERPBR.txt\" using 1:4 title \"stockPBR\" with linespoint";
				
				String temp2 = "plot [x=0:2] [0:20] exp(x**2)";
				
				String temp3 = "set terminal png";
				
				String temp4 = "set output 'filename.png'";
				
				String temp5 = "set output";*/
				
				//String cLine = "\"d:\\Program Files\\gnuplot\\bin\\gnuplot.exe\" d:\\Program Files\\gnuplot\\bin\\sin-plot.gnuplot\"";
				//String cLine = "\"d:/Program Files/gnuplot/bin/gnuplot.exe\"" + " \"d:/Program Files/gnuplot/bin/sin-plot.gnuplot\"";
				
				//ProcessBuilder proc = new ProcessBuilder( "\"d:/Program Files/gnuplot/bin/gnuplot.exe\"" , " \"d:/Program Files/gnuplot/bin/sin-plot.gnuplot\"");                         
				
				//proc.start();
				
				outputStream = new FileOutputStream("D:/Java/Test/pricePER.gnuplot");
				out1 = new OutputStreamWriter(outputStream, "utf-8");
				
				outputStream2 = new FileOutputStream("D:/Java/Test/pricePBR.gnuplot");
				out2 = new OutputStreamWriter(outputStream2, "utf-8");
				
				//stockName = "1101";
				
				out1.write("set title \"" + stockName + "\"\r\n");
				out1.write("set y2tics border\r\n");
				out1.write("set terminal png\r\n");
				out1.write("set output \"" + stockName + "_pricePER.png\"\r\n");
				out1.write("plot \"stockPricePERPBR.txt\" using 1:2 title \"stockPrice\" with linespoint, \"stockPricePERPBR.txt\" using 1:3 axes x1y2 title \"stockPER\" with linespoint\r\n");
				
				out2.write("set title \"" + stockName + "\"\r\n");
				out2.write("set y2tics border\r\n");
				out2.write("set terminal png\r\n");
				out2.write("set output \"" + stockName + "_pricePBR.png\"\r\n");
				out2.write("plot \"stockPricePERPBR.txt\" using 1:2 title \"stockPrice\" with linespoint, \"stockPricePERPBR.txt\" using 1:4 axes x1y2 title \"stockPBR\" with linespoint\r\n");
				
				out1.close();
				outputStream.close();
				out2.close();
				outputStream2.close();
				
				String[] s = {"\"d:\\Program Files\\gnuplot\\bin\\wgnuplot.exe\"", "pricePER.gnuplot"};
				String[] s2 = {"\"d:\\Program Files\\gnuplot\\bin\\wgnuplot.exe\"", "pricePBR.gnuplot"};
				
			    //Process process = Runtime.getRuntime().exec(s);
			    //String[] s 
		
				
				try {
					//System.out.println("======");
					Process p = Runtime.getRuntime().exec(s);
					//p = Runtime.getRuntime().exec("cmd /c "+cLine+temp3);
					//p = Runtime.getRuntime().exec("cmd /c "+cLine+temp4);
					//p = Runtime.getRuntime().exec("cmd /c "+cLine+temp5);
					p.waitFor();
					p.destroy();			
					
					Process p2 = Runtime.getRuntime().exec(s2);
					p2.waitFor();
					p2.destroy();
				} catch (Exception e) {
					e.printStackTrace();
				}
				
				/*File f = new File("D:/Java/Test/stockPricePERPBR.txt");
				f.delete();
				f = new File("D:/Java/Test/pricePER.gnuplot");
				f.delete();
				f = new File("D:/Java/Test/pricePBR.gnuplot");
				f.delete();*/
				
			} catch (Exception e) {
				e.printStackTrace();
			}
			

			j += 3;
			c = rs.getCell(0, j);
			stockName = c.getContents();
		}
		
		rwb.close();
		is.close();
		
		/*
			String cLine = "\"c:\\Program Files\\Graphviz2.21\\bin\\neato.exe\" -Tpng -o" +
				" e:\\School\\SchoolProject\\graphviz\\JavaOutput.png e:\\School\\SchoolProject\\graphviz\\out.txt";
			Process p = Runtime.getRuntime().exec("cmd /c "+cLine);		
			p.waitFor();
			p.destroy();
		 */
		 
		
		/*c  = rs.getCell(0, row);
		cContent = c.getContents();
		
		label = new Label(1, tableRow++, cContent);
		sheet.addCell(label);*/
	}
}
