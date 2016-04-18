public final class XlsProcessor {  
	private static final XlsProcessor processor = new XlsProcessor();
	
	private XlsProcessor(){}
	
	public static XlsProcessor getInstance() {
		return processor ;  
	}
	
	public File convert(File cvsFile) throws IOException, RowsExceededException, WriteException{   
		FileReader fileReader = new FileReader(cvsFile);   
		BufferedReader bufferedReader = new BufferedReader(fileReader);
		String xlsFileName = getFileName(cvsFile.getName());
		File xls = new File(xlsFileName);
		WritableWorkbook workbook = Workbook.createWorkbook(xls);
		WritableSheet sheet = workbook.createSheet("CSV", 0);
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
		return xls;  
	}
	
	public boolean compare( File a, File b) throws BiffException, IOException {
		Workbook aa = Workbook.getWorkbook(a);
		Workbook bb = Workbook.getWorkbook(b);
		Sheet aaSheet = aa.getSheet(0);
		Sheet bbSheet = bb.getSheet(0) ; 
		int aaSheetRows = aaSheet.getRows();
		int bbSheetRows = bbSheet.getRows();
		int aaSheetCols = aaSheet.getColumns();
		int bbSheetCols = bbSheet.getColumns(); 
		boolean result = true;
		if(aaSheetRows == bbSheetRows && aaSheetCols == bbSheetCols){
			for(int i=0;i<aaSheetRows;i++){
				for(int j=0;j<aaSheetCols;j++){
					String aaCell = aaSheet.getCell(j, i).getContents();
					String bbCell = bbSheet.getCell(j, i).getContents();
					if(!aaCell.equals(bbCell)){ 
						System.out.println("Row:"+(i+1)+",Column:"+( j+1)+",one is "+aaCell+",the other is "+bbCell);
						result = false;
						break;
					}
				}
				if(result){
					continue;
				}else{
					break;
				}
			}   
		}   
		System.out.println ("Compare result : "+result);
		return result;
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
					cellList. add(cell);
				}
			}
		}
		return cellList; 
	}
}