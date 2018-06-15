package docs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.model.InternalWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class Excel {

	
	String excelFolderPath, xmlFolderPath ,newFile ,existingFilePath , existingDynamicFilePath ;
	
	FileOutputStream out;
	
	Workbook excelWorkBook,excelWorkBookD;
	
	Sheet excelWSheet,excelWSheetDynamic,excelWSheetDefault;
	
	Row row;
	
	Cell cell;
	
	FileInputStream ExcelFile, excelFileDynamic;
	
	File excelFolder,xmlFolder;
	
	String existence = "existing", existence1="dynamic";
	
	// constructor when used existing excel file
		public Excel(String FilePath) throws Exception {
			
			
			ExcelFile = getExcelFile(FilePath);
			
			excelWorkBook = Init(ExcelFile,FilePath);
			
			excelWSheetDefault = excelWorkBook.getSheetAt(0);
				
				this.existingFilePath = FilePath;
				
			
		}

		
		//initialize workbook objects
		
		public Workbook Init(FileInputStream ExcelFile,String FilePath) throws IOException {
			
			if((FilenameUtils.getExtension(FilePath)).equalsIgnoreCase("xlsx")) {
				
				
				return new XSSFWorkbook(ExcelFile);
			
			}
			else if(FilenameUtils.getExtension(FilePath).equalsIgnoreCase("xls")){
				
				return new HSSFWorkbook(ExcelFile);
				
			}
			return null;
			
		
			
		}
		
		//initialize workbook objects of new excel file
		
				public Workbook Init(String FilePath) throws IOException {
					
					if((FilenameUtils.getExtension(FilePath)).equalsIgnoreCase("xlsx")) {
						
						
						return new XSSFWorkbook();
					
					}
					else if(FilenameUtils.getExtension(FilePath).equalsIgnoreCase("xls")){
						
						return new HSSFWorkbook();
						
					}
					return null;
					
				
					
				}
				
		
		// default constructor if you want to create new workbook and do operations on it 
		public Excel() {
			// TODO Auto-generated constructor stub
		}


		// returns Input stream of particular excel
		public FileInputStream getExcelFile(String path) throws Exception{
			return new FileInputStream(new File(path));
		}
		
		
		// creates output stream of excel according file existence and vice versa
		public FileOutputStream output(String exist) throws FileNotFoundException {
			
			if(exist.equalsIgnoreCase("existing"))
			return new FileOutputStream(existingFilePath);
			else if(exist.equalsIgnoreCase("newFile"))
				return new FileOutputStream(newFile);
			else if(exist.equalsIgnoreCase("dynamic"))
				return new FileOutputStream(existingDynamicFilePath);
			return null;
		}
		
		
		// function to set default sheet by sheet Name which uses excelWbook as reference
		public void setSheetDefault(String sheetName) {
			
			excelWSheetDefault = getSheetObject(sheetName);
		}
		
		
		// function to set default sheet by sheet Number which uses excelWbook as reference
		public void setSheetDefault(int sheetNum) {
			
			excelWSheetDefault = getSheetObject(sheetNum);
		}
		
		
		// function to set default sheet by workbook,sheet Name which uses excelWbookD as reference
		public void setSheetDefault(String workBookname, String sheetName) throws Exception {
			
			excelWSheetDefault = getSheetObject(workBookname,sheetName);

		}
		
		// function to set default sheet by workbook, sheet Name which uses excelWbookD as reference
		public void setSheetDefault(String workBookname, int sheetNum) throws Exception {
			
			excelWSheetDefault = getSheetObject(workBookname,sheetNum);
		}
		
		
		//function to create new work book
		public void createBook(String name) throws Exception {
			

			excelFolder= new File(new File("").getAbsolutePath()+File.separator+"ExcelOps");  // Folder to store excel files created dynamically
			
			xmlFolder= new File(new File("").getAbsolutePath()+File.separator+"Xml");    // Folder to store xml files created dynamically
			
			// creating folders according to their existence
			if(excelFolder.exists() ){
				
				System.out.println("file already exists");
			}
			else {
				
				excelFolder.mkdir();
			}
			
			if(xmlFolder.exists() ){
				
				System.out.println("file already exists");
			}
			else {
				
				xmlFolder.mkdir();
			}
			
			// initializing path variables
			
		    this.excelFolderPath  	= excelFolder.getPath();
		    this.xmlFolderPath	= xmlFolder.getPath();
			this.newFile  = excelFolderPath+File.separator+name;
			
		    excelWorkBookD = Init(name);
			 excelWorkBookD.createSheet();
			 updateFile("newFile");
		    
		}
		
		
		// function to create Sheet by passing sheet name uses excelWbook as reference
		public Sheet createSheet( String name) {
			try {
	
			excelWSheet = excelWorkBook.createSheet(name);
				
			updateFile(existence);
			
				return excelWSheet;
		
				
		} catch(Exception e) {
				return null;
		}
		}
		
		// function to create Sheet by passing workbook,sheet name uses excelWbookD as reference
		public Sheet createSheet( String workBookName,String name) {
			try {

				initDWbook(workBookName);
				
							
				excelWSheetDynamic = excelWorkBookD.createSheet(name);
		        
		         updateFile(existence1);
				
				return excelWSheetDynamic;
		
		} catch(Exception e) {
				return null;
		}
		}

		
		
		// function to get default Sheet 
		public Sheet getSheetObject() {
			
				return excelWSheetDefault;
			
		}
		
	
		// function to get Sheet by passing workbook,sheet number 
		public Sheet getSheetObject(String workBookname,int sheetNum) throws Exception {
			
				initDWbook(workBookname);
				
				System.out.println(excelWorkBookD.getSheetAt(sheetNum).getSheetName());
				return excelWorkBookD.getSheetAt(sheetNum);
			
		}
		
		// function to get Sheet by passing workbook,sheet name
		public Sheet getSheetObject(String workBookname,String sheetName) throws Exception {
			
				initDWbook(workBookname);
				//System.out.println(excelWorkBookD.getSheet(sheetName).getSheetName());	
				return excelWorkBookD.getSheet(sheetName);
			
		}
				
		// function to get Sheet by passing sheet number
		public Sheet getSheetObject(int sheetNum) {
			
				return excelWorkBook.getSheetAt(sheetNum);
			
		}
		
		// function to get Sheet by passing sheet name
		public Sheet getSheetObject(String sheetName) {
			
				return excelWorkBook.getSheet(sheetName);
			
		}

		
		/*
		// function to create rows in sheet by passing sheet 
		public void createRows(int Rows,int cols)  {
			
			try {
		
			
			
			
			for(int roow=0;roow<Rows;++roow){
				
				
				for(int columnNum=0;columnNum<cols;columnNum++) {
					
				createRowD().createCell(columnNum);
		           
				}
				
				//System.out.println(roow.getPhysicalNumberOfCells());
			}
			
			updateFile("existing");
			
			}
			catch(Exception e) {
				
				e.printStackTrace();
			}
			
		}
		
		
		// function to create rows in sheet by passing sheet 
				public void createRows(int Rows,int cols,String sheetName)  {
					
					try {
				
					
					
					
					for(int roow=0;roow<Rows;++roow){
						
						
						for(int columnNum=0;columnNum<cols;columnNum++) {
							
						createRowD(sheetName).createCell(columnNum);
				           
						}
						
						//System.out.println(roow.getPhysicalNumberOfCells());
					}
					
					updateFile("existing");
					
					}
					catch(Exception e) {
						
						e.printStackTrace();
					}
					
				}
				
				
				// function to create rows in sheet by passing sheet 
				public void createRows(int Rows,int cols,int sheetNum)  {
					
					try {
				
					
					
					
					for(int roow=0;roow<Rows;++roow){
						
						
						for(int columnNum=0;columnNum<cols;columnNum++) {
							
						createRowD(sheetNum).createCell(columnNum);
				           
						}
						
						//System.out.println(roow.getPhysicalNumberOfCells());
					}
					
					updateFile("existing");
					
					}
					catch(Exception e) {
						
						e.printStackTrace();
					}
					
				}
				
		
				// function to create rows in sheet by passing sheet 
				public void createRows(int Rows,int cols,String workBookName, int sheetNum)  {
					
					try {
				
					
					
					
					for(int roow=0;roow<Rows;++roow){
						
						
						for(int columnNum=0;columnNum<cols;columnNum++) {
							
						createRowD(workBookName,sheetNum).createCell(columnNum);
				           
						}
						
						//System.out.println(roow.getPhysicalNumberOfCells());
					}
					
					updateFile("existing");
					
					}
					catch(Exception e) {
						
						e.printStackTrace();
					}
					
				}
				
				

				// function to create rows in sheet by passing sheet 
				public void createRows(int Rows,int cols,String workBookName, String sheetName)  {
					
					try {
				
					
					
					
					for(int roow=0;roow<Rows;++roow){
						
						
						for(int columnNum=0;columnNum<cols;columnNum++) {
							
						createRowD(workBookName,sheetName).createCell(columnNum);
				           
						}
						
						//System.out.println(roow.getPhysicalNumberOfCells());
					}
					
					updateFile("existing");
					
					}
					catch(Exception e) {
						
						e.printStackTrace();
					}
					
				}*/
		
		
		public void delRow(int rowNum) {
			
			
		}
		
		
		// function to create row in sheet default 
				public Row createRowD()  {
					
					int rowNum = 1;
					
					try {
						
					 if(excelWSheetDefault.getLastRowNum()!= 0)
							 rowNum +=excelWSheetDefault.getLastRowNum();
					
					}
					catch(Exception e) {
						
						e.printStackTrace();
					}
					
					return excelWSheetDefault.createRow(rowNum);
					
				}
		

				// function to create row in sheet by passing sheet num
				public Row createRowD(int sheetNum)  {
					int rowNum = 1;
					
					try {
				
									 
						 if(getSheetObject(sheetNum).getLastRowNum()!= 0)
							 rowNum +=getSheetObject(sheetNum).getLastRowNum();
						 
					
					}
					catch(Exception e) {
						
						e.printStackTrace();
					}
					
					return getSheetObject(sheetNum).createRow(rowNum);
					
				}
		
				
				// function to create row in sheet by passing sheet name
				public Row createRowD(String sheetName)  {
					int rowNum = 1;
					
					try {
				
							
						 if(getSheetObject(sheetName).getLastRowNum()!= 0) {
							 rowNum +=getSheetObject(sheetName).getLastRowNum();
							 System.out.println(getSheetObject(sheetName).getLastRowNum());
						 }
						
					}
					catch(Exception e) {
						
						e.printStackTrace();
					}
					
					return getSheetObject(sheetName).createRow(rowNum);
					
				}
				
				
				// function to create row in sheet by passing workbook, sheet name 
				public Row createRowD(String workBookName,String sheetName) throws Exception  {
					int rowNum = 1;
					
					try {
				
						///System.out.println(getSheetObject(workBookName,sheetName).getSheetName());
							
						 if(getSheetObject(workBookName,sheetName).getLastRowNum()!= 0) {
							 rowNum +=getSheetObject(workBookName,sheetName).getLastRowNum();
						 
						
						System.out.println(getSheetObject(workBookName,sheetName).getLastRowNum());
						 }
					 
						
					}
					catch(Exception e) {
						
						e.printStackTrace();
					}
					
					//updateFile(existence1);
					return getSheetObject(workBookName, sheetName).createRow(rowNum);
					
				}
		
				
				// function to create row in sheet by passing work book, sheet number 
				public Row createRowD(String workBookName,int sheetNum) throws Exception  {
					int rowNum = 1;
					
					try {
				
							 if(getSheetObject(workBookName,sheetNum).getLastRowNum()!= 0)
							 rowNum +=getSheetObject(workBookName,sheetNum).getLastRowNum();
						
					}
					catch(Exception e) {
						
						e.printStackTrace();
					}
					getSheetObject(workBookName, sheetNum).createRow(rowNum);
					updateFile(existence1);
					return getSheetObject(workBookName, sheetNum).createRow(rowNum);
					
				}
		
				
		
		
				
		// function to fill rows in sheet by passing data Object of default sheet
		public void insertData(Object data[][])  {
			
			
			
				int columnNum =0;
				for(Object[] roow:data){
					
					columnNum=0;
					row= createRowD();
					for(Object field:roow) {
						
			            if (field instanceof String) {
			                row.createCell(columnNum++).setCellValue((String) field);
			            } else if (field instanceof Integer) {
			            	 row.createCell(columnNum++).setCellValue((Integer) field);
			            }
					}
					updateFile(existence);
				}
				
			
		}

		
		
		
		// function to fill rows in sheet by passing data Object,sheet name  
		public void insertData(Object data[][], String sheetName )  {
			
			try {
				
		int columnNum =0;
			for(Object[] roow:data){
				
				columnNum=0;
				row= createRowD(sheetName);
				for(Object field:roow) {
					
		            if (field instanceof String) {
		                row.createCell(columnNum++).setCellValue((String) field);
		            } else if (field instanceof Integer) {
		            	 row.createCell(columnNum++).setCellValue((Integer) field);
		            }
				}
				updateFile(existence);
			}
			
			
			}
			catch(Exception e) {
				
				e.printStackTrace();
			}
			
		}
		
		// function to fill rows in sheet by passing data Object,sheet number
		public void insertData(Object data[][], int sheetNum )  {
			
			try {
				
				int columnNum =0;
					for(Object[] roow:data){
						
						columnNum=0;
						row= createRowD(sheetNum);
						for(Object field:roow) {
							
				            if (field instanceof String) {
				                row.createCell(columnNum++).setCellValue((String) field);
				            } else if (field instanceof Integer) {
				            	 row.createCell(columnNum++).setCellValue((Integer) field);
				            }
						}
						updateFile(existence1);
						
					}
					
			
			}
			catch(Exception e) {
				
				e.printStackTrace();
			}
			
		}
		
		
		// function to fill rows in sheet by passing data Object,workbook,sheet number
		public void insertData(Object data[][],String workBookName, int sheetNum )  {
			
			try {
				
				int columnNum =0;
					for(Object[] roow:data){
						
						columnNum=0;
						row= createRowD(workBookName,sheetNum);
						for(Object field:roow) {
							
				            if (field instanceof String) {
				                row.createCell(columnNum++).setCellValue((String) field);
				            } else if (field instanceof Integer) {
				            	 row.createCell(columnNum++).setCellValue((Integer) field);
				            }
						}
						updateFile(existence1);
					}
					
		
			}
			catch(Exception e) {
				
				e.printStackTrace();
			}
			
		}
		
		
		// function to fill rows in sheet by passing data Object,workbook,sheet name 
		public void insertData(Object data[][],String workBookName, String sheetName )  {
			
			try {
				
				int columnNum =0;
					for(Object[] roow:data){
						
						columnNum=0;
						row= createRowD(workBookName,sheetName);
						for(Object field:roow) {
							
				            if (field instanceof String) {
				                row.createCell(columnNum++).setCellValue((String) field);
				            } else if (field instanceof Integer) {
				            	 row.createCell(columnNum++).setCellValue((Integer) field);
				            }
						}
						updateFile("dynamic");
				
					}
					
					
			}
			catch(Exception e) {
				
				e.printStackTrace();
			}
			
		}
		
		
		/* functions to fill rows with map 
		
		case 1: with only map as argument , which fills default sheet
		case 2: with map and sheet number as argument, fills data of excelWbook
		case 3: with map and sheet name .......................................
		case 4: with map , workbook, sheet number , fills data of excelWbookD
		case 5: with map, workbook, sheet name , fills data of excelWbookD
		
		
		
		*/
		public void fillRows(  Map<Integer,LinkedList<String>> map) {
			
			
			int column=0;
			for(Entry<Integer, LinkedList<String>> b:map.entrySet()) {
				
				
				LinkedList<String> rows = b.getValue();
				
				// Loop for columns
				
				row= createRowD();
				
				for(String val:rows) {
					
					row.createCell(column++).setCellValue(val);
						
					}
				
				updateFile(existence);
				
					
					
				}
			
			
			
		}
		
		
		public void fillRows(  Map<Integer,LinkedList<String>> map, int sheetNum) {
			
			

			int column=0;
			for(Entry<Integer, LinkedList<String>> b:map.entrySet()) {
				
				
				LinkedList<String> rows = b.getValue();
				
				// Loop for columns
				row= createRowD(sheetNum);
				for(String val:rows) {
					
					row.createCell(column++).setCellValue(val);
						
					}
				
				updateFile(existence);
					
				}
			
			
			
		}
		
		
		public void fillRows(  Map<Integer,LinkedList<String>> map, String sheetName) {

			int column=0;
			for(Entry<Integer, LinkedList<String>> b:map.entrySet()) {
				
				
				LinkedList<String> rows = b.getValue();
				
				// Loop for columns
				row= createRowD(sheetName);
				for(String val:rows) {
					
					row.createCell(column++).setCellValue(val);
						
					}
				
				updateFile(existence);
				}
			
			
			
		}
		
		
		public void fillRows( String workBookName, Map<Integer,LinkedList<String>> map,String sheetName) throws Exception {
			

			int column=0;
			for(Entry<Integer, LinkedList<String>> b:map.entrySet()) {
				
				
				LinkedList<String> rows = b.getValue();
				
				// Loop for columns
				row= createRowD(workBookName,sheetName);
				
				for(String val:rows) {
				
					row.createCell(column++).setCellValue(val);
						
					}
				updateFile(existence);

					
				}
			
		
		}
		
	public void fillRows( String workBookName, Map<Integer,LinkedList<String>> map,int sheetNum) throws Exception {
			

		int column=0;
		for(Entry<Integer, LinkedList<String>> b:map.entrySet()) {
			
			
			LinkedList<String> rows = b.getValue();
			
			// Loop for columns
			row= createRowD(workBookName,sheetNum);
			
			for(String val:rows) {
				
				row.createCell(column++).setCellValue(val);
					
				}
			updateFile(existence);

			}
		
			
		}
		
		
	
	/*generic functions to fill rows with List,Set
	
	case 1: with only Collection as argument , which fills default sheet
	case 2: with Collection and sheet number as argument, fills data of excelWbook
	case 3: with Collection and sheet name .......................................
	case 4: with Collection , workbook, sheet number , fills data of excelWbookD
	case 5: with Collection, workbook, sheet name , fills data of excelWbookD
	
	
	
	*/
	
	
		
	public <T> void fillRows( Collection<Collection<T>> c) {
		int columnNum=0;
		
	
		for(Collection<T> roow:c){
			
			columnNum = 0;
			
			row= createRowD();
			for(T field:roow) {
				
	            if (field instanceof String) {
	                row.createCell(columnNum++).setCellValue((String) field);
	            } else if (field instanceof Integer) {
	            	 row.createCell(columnNum++).setCellValue((Integer) field);
	            }
			}
			updateFile(existence);
			
		}
		
		
		
	}
		
		public <T> void fillRows( Collection<Collection<T>> c, String sheetName) {
			int columnNum=0;
			
		
			for(Collection<T> roow:c){
				
				columnNum = 0;
				
				row= createRowD(sheetName);
				for(T field:roow) {
					
		            if (field instanceof String) {
		                row.createCell(columnNum++).setCellValue((String) field);
		            } else if (field instanceof Integer) {
		            	 row.createCell(columnNum++).setCellValue((Integer) field);
		            }
				}
				updateFile(existence);
				
			}
			
			
			
		}
		
		public <T> void fillRows( Collection<Collection<T>> c, int sheetNum) {
			int  columnNum=0;
			
			
			for(Collection<T> roow:c){
				columnNum=0;
				
				row = createRowD(sheetNum);
				for(T field:roow) {
					
					
					 if (field instanceof String) {
			                row.createCell(columnNum++).setCellValue((String) field);
			            } else if (field instanceof Integer) {
			            	 row.createCell(columnNum++).setCellValue((Integer) field);
			            }
				}
				
				updateFile(existence);
			}
			
		
			
		}
		
		
		public <T> void fillRows( Collection<Collection<T>> c,String workBookName, String sheetName) throws Exception {
			int columnNum=0;
			
		
			for(Collection<T> roow:c){
				
				columnNum = 0;
				
				row= createRowD(workBookName,sheetName);
				for(T field:roow) {
					
		            if (field instanceof String) {
		                row.createCell(columnNum++).setCellValue((String) field);
		            } else if (field instanceof Integer) {
		            	 row.createCell(columnNum++).setCellValue((Integer) field);
		            }
				}
				updateFile(existence1);
				
			}
			
			
			
		}
		

		public <T> void fillRows( Collection<Collection<T>> c,String workBookName,int sheetNum) throws Exception {
			int columnNum=0;
			
		
			for(Collection<T> roow:c){
				
				columnNum = 0;
				
				row= createRowD(workBookName,sheetNum);
				for(T field:roow) {
					
		            if (field instanceof String) {
		                row.createCell(columnNum++).setCellValue((String) field);
		            } else if (field instanceof Integer) {
		            	 row.createCell(columnNum++).setCellValue((Integer) field);
		            }
				}
				updateFile(existence1);
				
			}
			
			
			
		}
		
		// function to initialize Header in sheet by passing data Object,sheet 
		public void headerRow(Object[] data) {
			
		        row = excelWSheetDefault.getRow(0);
		        int columns = 0;
		        for(Object field:data) {
		        	
		        	 cell= row.getCell(columns++);
		        	
		        //	System.out.println(cell.getCellType());
		        	
		        	if (field instanceof String) {
		                cell.setCellValue((String) field);
		            } else if (field instanceof Integer) {
		                cell.setCellValue((Integer) field);
		            }
		        	
		        }
		        
		        
		        updateFile(existence);
		        
		     }

		
		
		// function to initialize Header in sheet by passing data Object,sheet 
		public void headerRow(Object[] data,String sheetName) {
			
		       excelWSheet = getSheetObject(sheetName);
		       
		       row = excelWSheet.getRow(0);
		        int columns = 0;
		        for(Object field:data) {
		        	
		        	 cell= row.getCell(columns++);
		        	
		        //	System.out.println(cell.getCellType());
		        	
		        	if (field instanceof String) {
		                cell.setCellValue((String) field);
		            } else if (field instanceof Integer) {
		                cell.setCellValue((Integer) field);
		            }
		        	
		        }
		        
		        
		        updateFile(existence);
		        
		     }

		
		

		// function to initialize Header in sheet by passing data Object,sheet number
		public void headerRow(Object[] data,int sheetNum) {
			
		       excelWSheet = getSheetObject(sheetNum);
		       
		       row = excelWSheet.getRow(0);
		        int columns = 0;
		        for(Object field:data) {
		        	
		        	 cell= row.getCell(columns++);
		        	
		        //	System.out.println(cell.getCellType());
		        	
		        	if (field instanceof String) {
		                cell.setCellValue((String) field);
		            } else if (field instanceof Integer) {
		                cell.setCellValue((Integer) field);
		            }
		        	
		        }
		        
		     }


		
		// function to fill rows in sheet by passing xml file,sheet   to be implemented
		public void fillRows(String file, Sheet sheet )  {
				
				try {
					
					if((FilenameUtils.getExtension(file)).equalsIgnoreCase("xml")) {
		
						
						File inputFile = new File(xmlFolderPath+File.separator+file);
				         DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
				         DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
				         Document doc = dBuilder.parse(inputFile);
				         doc.getDocumentElement().normalize();
				         System.out.println("Root element :" + doc.getDocumentElement().getNodeName());
				         NodeList nList = doc.getElementsByTagName("student");
				         
				         
					//	out  = output();
				
						int columns =0;
						
						ArrayList<String> list = getRow(sheet,0);
						
						System.out.println(list.toString());
						int k=0;
						for (int temp = 0; temp < nList.getLength(); temp++) {
				            Node nNode = nList.item(temp);
				            System.out.println("\nCurrent Element :" + nNode.getNodeName());
				            Element eElement = (Element) nNode;
				            Row row = sheet.getRow(temp+1);
				            
				            for(String var:list) {
				            	System.out.println(var);
				            	Cell cell= row.getCell(columns);
				            	System.out.println(eElement.getElementsByTagName(var));
				            	
				            	if(k++==0)
				            	cell.setCellValue(eElement.getAttribute(var));
				            	else
				            	cell.setCellValue(eElement.getElementsByTagName(var).item(0).getTextContent());
				            	
				            }
				            
				           k=0;
				            
				         }
				
				excelWorkBook.write(out);
				out.close();
				
					}
				}
				catch(IOException e) {
					
					e.printStackTrace();
				}
				catch(Exception e) {
					
					e.printStackTrace();
				}
				
			}
		

		/*generic functions to get row count
		
		case 1: with No argument , from default sheet
		case 2: with  sheet number as argument, from excelWbook
		case 3: with  sheet name .......................................
		case 4: with  workbook, sheet number , from excelWbookD
		case 5: with  workbook, sheet name , from excelWbookD
		
		
		
		*/
		
		 // function to get row count of default sheet
		public int getRowCount() throws Exception {
			
				int totRowCount 					= 0;
				
				try {
					
					totRowCount 			= excelWSheetDefault.getLastRowNum();
					excelWorkBook.close();

			
				} catch(Exception e) {
					throw new Exception("Unable to get row count");
				}
				
					return totRowCount;
				}
		
		
		 // function to get row count by passing sheet number
		   public int getRowCount(int sheetNo) throws Exception {
			
				int totRowCount 					= 0;
				
				try {
					
							
					totRowCount 			= getSheetObject(sheetNo).getLastRowNum();
					excelWorkBook.close();
		
					
				} catch(Exception e) {
					throw new Exception("Unable to get row count");
				}
				
					return totRowCount;
			}
		
		 // function to get row count by passing sheet number
			public int getRowCount(String sheetName) throws Exception {
				
					int totRowCount 					= 0;
					
					try {
						
								
						totRowCount 			= getSheetObject(sheetName).getLastRowNum();
						excelWorkBook.close();
			
						
					} catch(Exception e) {
						throw new Exception("Unable to get row count");
					}
					
						return totRowCount;
					}
		
			
			 // function to get row count by passing sheet number
			public int getRowCount(String workBookName,String sheetName) throws Exception {
				
					int totRowCount 					= 0;
					
					try {
						
						initDWbook(workBookName);
						totRowCount 			= getSheetObject(workBookName,sheetName).getLastRowNum();
						excelWorkBookD.close();
			
						
					} catch(Exception e) {
						throw new Exception("Unable to get row count");
					}
					
						return totRowCount;
					}
		
			
			 // function to get row count by passing sheet number
			public int getRowCount(String workBookName,int sheetNum) throws Exception {
				
					int totRowCount 					= 0;
					
					try {
						
						initDWbook(workBookName);
								
						totRowCount 			=  getSheetObject(workBookName,sheetNum).getLastRowNum();
						excelWorkBookD.close();
			
						
					} catch(Exception e) {
						throw new Exception("Unable to get row count");
					}
					
						return totRowCount;
					}
		
			
		// function to get row number of Sheet value by passing sheet and value
		@SuppressWarnings("deprecation")
		public int getRowNumber(String searchValue) {
				
				try {
					
					for (Row row : excelWSheetDefault) 
						for (Cell cell : row) 
							if (cell.getCellType() == Cell.CELL_TYPE_STRING) 
								if (cell.getRichStringCellValue().getString().trim().equalsIgnoreCase(searchValue)) 
									return row.getRowNum();
					
				} catch(Exception e) {
					e.printStackTrace();
				}
				
				return 0;
			}
		
		
		// function to get row number of Sheet value by passing sheet and value
		@SuppressWarnings("deprecation")
		public int getRowNumber(String sheetName,String searchValue) {
				
				try {
					
					excelWSheet= excelWorkBook.getSheet(sheetName);
					
					for (Row row : excelWSheet) 
						for (Cell cell : row) 
							if (cell.getCellType() == Cell.CELL_TYPE_STRING) 
								if (cell.getRichStringCellValue().getString().trim().equalsIgnoreCase(searchValue)) 
									return row.getRowNum();
					
				} catch(Exception e) {
					e.printStackTrace();
				}
				
				return 0;
			}
		

		
		// function to get row number of Sheet value by passing sheet and value
		@SuppressWarnings("deprecation")
		public int getRowNumber(String workBookName,String sheetName,String searchValue) {
				
				try {
					initDWbook(workBookName);
					excelWSheetDynamic = excelWorkBookD.getSheet(sheetName);
					
					for (Row row : excelWSheetDynamic) 
						for (Cell cell : row) 
							if (cell.getCellType() == Cell.CELL_TYPE_STRING) 
								if (cell.getRichStringCellValue().getString().trim().equalsIgnoreCase(searchValue)) 
									return row.getRowNum();
					
				} catch(Exception e) {
					e.printStackTrace();
				}
				
				return 0;
			}
		

		
		
		// function to set value in particular cell of Sheet by passing workbook,sheet name,row number, column number, value
		 public void setCellData(String workBookName, String sheetName, int rownum, int columnum, String value)  {
			
			 
			        try {
			        	
			        	getCellObject(workBookName, sheetName, rownum, columnum).setCellValue(value);
						
					} catch (Exception e) {
					
						e.printStackTrace();
					}  
			       
			        updateFile(existence1);
					
				
				
			}
				 

			// function to set value in particular cell of Sheet by passing workbook,sheet number,row number, column number, value
		 public void setCellData(String workBookName, int sheetNum, int rownum, int columnum, String value)  {
			
			 
			        try {
			        	
			        	getCellObject(workBookName, sheetNum, rownum, columnum).setCellValue(value);
						
					} catch (Exception e) {
					
						e.printStackTrace();
					}  
			       
			        updateFile(existence1);
					
				
				
			}
				 

				
				 
		// function to set value in particular cell of Sheet by passing sheet name ,row number, column number, value
		 public void setCellData(String sheetName, int rownum, int columnum, String value)  {
			
				try {
				
					
					getCellObject(sheetName,rownum, columnum).setCellValue(value);
			        
			        updateFile(existence);
					
				} 
				catch (Exception e) {
					e.printStackTrace();
				}
				
				
			}
		 
		 
		
		// function to set value in particular cell of Sheet by passing sheet number,row number, column number, value
		 public void setCellData(int sheetNum, int rownum, int columnum, String value)  {
			
				try {
				
				
					getCellObject(sheetNum, rownum, columnum).setCellValue(value);
			      
			        updateFile(existence);
				} 
				catch (Exception e) {
					e.printStackTrace();
				}
				
				
			}
		 
		 
		 

			// function to set value in particular cell of Sheet by passing row number, column number, value for default sheet
		 public void setCellData( int rownum, int columnum, String value)  {
			
				try {
				
					
					getCellObject( rownum, columnum).setCellValue(value);
			    
			        updateFile(existence);
					
				} 
				catch (Exception e) {
					e.printStackTrace();
				}
				
				
			}
	 


		// function to get value in particular cell of Sheet by passing workbook, sheet number,row number, column number
		public String getCellData(String workBookname,int sheetNum,int rownum, int columnum) {
				
				String value  = "";
				try {
				
			        value  =  getCellObject(workBookname,sheetNum, rownum, columnum).getStringCellValue();
			      
			     
			        return value;
				}
				catch(IOException e) {
					
					e.printStackTrace();
				}
				catch (Exception e) {
					e.printStackTrace();
				}
				
				return value;
		
		}
		
		
		// function to get value in particular cell of Sheet by passing workbook, sheet name ,row number, column number
		public String getCellData(String workBookname,String sheetName,int rownum, int columnum) {
				
				String value  = "";
				try {
									    
			        value  =  getCellObject(workBookname, sheetName, rownum, columnum).getStringCellValue();
			      
			        return value;
				}
				catch(IOException e) {
					
					e.printStackTrace();
				}
				catch (Exception e) {
					e.printStackTrace();
				}
				
				return value;
		
		}

		
		
		
		// function to get value in particular cell of Sheet by passing sheetName,row number, column number for excelWbook
		public String getCellData(String sheetName,int rownum, int columnum) {
				
				String value  = "";
				try {
				
			        value  =  getCellObject(sheetName, rownum, columnum).getStringCellValue();
			    	
			        return value;
				}
				catch (Exception e) {
					e.printStackTrace();
				}
				
				return value;
		
		}
		
		// function to get value in particular cell of Sheet by passing sheet number,row number, column number for excelWbook
		public String getCellData(int sheetNum,int rownum, int columnum) {
				
				String value  = "";
				try {
				
			        value  =  getCellObject(sheetNum, rownum, columnum).getStringCellValue();
			    	
			        return value;
				}
				catch (Exception e) {
					e.printStackTrace();
				}
				
				return value;
		
		}

		
		// function to get value in particular cell of Sheet by passing row number, column number for default sheet
		public String getCellData(int rownum, int columnum) {
				
				String value  = "";
				try {
				
					
			        value  = getCellObject( rownum, columnum).getStringCellValue();
			    	
			        return value;
				}
				catch (Exception e) {
					e.printStackTrace();
				}
				
				return value;
		
		}
	
		// function to get number of sheets present in  excelworkbook
		public int getNumberOfSheets() {
		
				return 	excelWorkBook.getNumberOfSheets();
		}
			
	
		// function to get number of sheets present in particular workbook
		public int getNumberOfSheets(String WorkBookName) throws Exception {
		
				initDWbook(WorkBookName);
				
				return excelWorkBookD.getNumberOfSheets();
		}
	
		
		
		public String getCellDataString(Cell cell) {
			
			String val = "";
			
			if(cell.getCellType()==0) {
				val 						= Double.toString(cell.getNumericCellValue());
	    	} else if(cell.getCellType()==1) {
	    		val 						= cell.getStringCellValue();
	    	} else if (cell.getCellType()==3) {
	    		val 						= "";
	    	}
			
			
			return val;
			
			
		}
		
		// function to get complete row in sheet by passing row number for default sheet 
		public ArrayList<String> getRow(int rownum){
		
				ArrayList<String> list = new ArrayList<String>();
			
				int numofColumns = excelWSheetDefault.getRow(0).getLastCellNum();
				
				
				
				for(int colCounter = 0 ; colCounter < numofColumns; colCounter++) {
					
					 list.add(getCellDataString(getCellObject(rownum,colCounter)));
				}
			
			return list;
		}
	
			// function to get complete row in sheet by passing sheetname , row number for excelWbook
			public ArrayList<String> getRow(String sheetName, int rownum){
			
					ArrayList<String> list = new ArrayList<String>();
				
					int numofColumns = getSheetObject(sheetName).getRow(0).getLastCellNum();
				
					for(int colCounter = 0 ; colCounter < numofColumns; colCounter++) {
						
						 list.add(getCellDataString(getCellObject(sheetName, rownum, colCounter )));
					}
				
				return list;
			}
			
			
			
			// function to get complete row in sheet by passing sheet number , row number for excel workbook  
			public ArrayList<String> getRow(int sheetNum, int rownum){
			
					ArrayList<String> list = new ArrayList<String>();
					 
					int numofColumns = getSheetObject(sheetNum).getRow(0).getLastCellNum();
					
					for(int colCounter = 0 ; colCounter < numofColumns; colCounter++) {
						
						 list.add(getCellDataString(getCellObject(sheetNum, rownum, colCounter )));
					}
				
				return list;
			}
				
		// function to get complete row in sheet by passing sheet, row number 
			public ArrayList<String> getRow(Sheet sheet, int rownum){
		
				ArrayList<String> list = new ArrayList<String>();
				
				
				int numofColumns = sheet.getRow(0).getLastCellNum();
			
				for(int colCounter = 0 ; colCounter < numofColumns; colCounter++) {
				
					 list.add(getCellDataString(getCellObject(sheet, rownum, colCounter)));
				}
			
			return list;
		}
	
	
		
		// function to get sheet data in hash map of default sheet
				public LinkedHashMap<String, ArrayList<String>> getSheetData( ) {
				
						LinkedHashMap<String, ArrayList<String>> map 	= new LinkedHashMap<String, ArrayList<String>>();
						
							
							try {
							
								// Get total number of used columns
								
								int numOfColumns 						= excelWSheetDefault.getRow(0).getLastCellNum();
						
								// Loop for rows starting from 1 as 0 is header of row
								
								for(int counter = 1; counter <= excelWSheetDefault.getLastRowNum(); counter ++) {
									
									ArrayList<String> mapData 			= new ArrayList<String>();
								
								
									
									// Loop for columns
									
									for(int colCounter = 1 ; colCounter < numOfColumns; colCounter++) {
										
											mapData.add(getCellDataString(getCellObject(counter,colCounter)));
										}
										
										map.put(getCellObject(counter,0).getStringCellValue(), mapData);
									}
									
								} catch(Exception e) {
									System.out.println(e.getMessage());
								}
							
							
							return map;
					
				}
		
		// function to get sheet data in hash map by passing sheet number for excelWbook
		public LinkedHashMap<String, ArrayList<String>> getSheetData( int sheetNum) {
		
				LinkedHashMap<String, ArrayList<String>> map 	= new LinkedHashMap<String, ArrayList<String>>();
				
					
					try {
						
						// Create object for sheet
						
						 excelWSheet 					= getSheetObject(sheetNum);
					
						// Get total number of used columns
						
						int numOfColumns 						= excelWSheet.getRow(0).getLastCellNum();
				
						// Loop for rows starting from 1 as 0 is header of row
						
						for(int counter = 1; counter <= excelWSheet.getLastRowNum(); counter ++) {
							
							ArrayList<String> mapData 			= new ArrayList<String>();
							
							
							// Loop for columns
							
							for(int colCounter = 1 ; colCounter < numOfColumns; colCounter++) {
								
								
									mapData.add(getCellDataString( getCellObject(sheetNum, counter, colCounter)));
								}
								
								map.put( getCellObject(sheetNum, counter, 0).getStringCellValue(), mapData);
							}
							
						} catch(Exception e) {
							System.out.println(e.getMessage());
						}
					
					
					return map;
			
		}
		
		
		// function to get sheet data in hash map by passing sheet name for excelWbook
		public LinkedHashMap<String, ArrayList<String>> getSheetData( String sheetName) {
		
				LinkedHashMap<String, ArrayList<String>> map 	= new LinkedHashMap<String, ArrayList<String>>();
				
					
					try {
						
						// Create object for sheet
						
						 excelWSheet 					= getSheetObject(sheetName);
						
						
						int numOfColumns 						= excelWSheet.getRow(0).getLastCellNum();
				
						// Loop for rows starting from 1 as 0 is header of row
						
						for(int counter = 1; counter <= excelWSheet.getLastRowNum(); counter ++) {
							
							ArrayList<String> mapData 			= new ArrayList<String>();
						
							String val 							= "";
							
							// Loop for columns
							
							for(int colCounter = 1 ; colCounter < numOfColumns; colCounter++) {
								
								cell 							= getCellObject(sheetName, counter, colCounter);
							
									mapData.add(getCellDataString(getCellObject(sheetName, counter, colCounter)));
								}
								
								map.put(getCellObject(sheetName, counter, 0).getStringCellValue(), mapData);
							}
							
						} catch(Exception e) {
							System.out.println(e.getMessage());
						}
					
					
					return map;
			
		}
		
		

		// function to get sheet data in hash map by passing workbook,sheet number for excelWbookD
		public LinkedHashMap<String, ArrayList<String>> getSheetData(String WorkBookName, String sheetName) {
		
				LinkedHashMap<String, ArrayList<String>> map 	= new LinkedHashMap<String, ArrayList<String>>();
				
					
					try {
						
						// Create object for sheet
						

						initDWbook(WorkBookName);
						
						excelWSheetDynamic					= getSheetObject(WorkBookName,sheetName);
						
						int numOfColumns 						= excelWSheetDynamic.getRow(0).getLastCellNum();
				
						// Loop for rows starting from 1 as 0 is header of row
						
						for(int counter = 1; counter <= excelWSheetDynamic.getLastRowNum(); counter ++) {
							
							ArrayList<String> mapData 			= new ArrayList<String>();
							
							
							// Loop for columns
							
							for(int colCounter = 1 ; colCounter < numOfColumns; colCounter++) {
								
									mapData.add(getCellDataString(getCellObject(WorkBookName,sheetName,counter,  colCounter )));
								}
								
								map.put(getCellObject(WorkBookName,sheetName,counter, 0).getStringCellValue(), mapData);
							}
							
						} catch(Exception e) {
							System.out.println(e.getMessage());
						}
					
					
					return map;
			
		}
	
		// function to get sheet data in hash map by passing workbook,sheet number for excelWbookD
				public LinkedHashMap<String, ArrayList<String>> getSheetData(String WorkBookName, int sheetNum) {
				
						LinkedHashMap<String, ArrayList<String>> map 	= new LinkedHashMap<String, ArrayList<String>>();
						
							
							try {
								
								// Create object for sheet
								

								initDWbook(WorkBookName);
								
								excelWSheetDynamic					= getSheetObject(WorkBookName,sheetNum);
								
								int numOfColumns 						= excelWSheetDynamic.getRow(0).getLastCellNum();
						
								// Loop for rows starting from 1 as 0 is header of row
								
								for(int counter = 1; counter <= excelWSheetDynamic.getLastRowNum(); counter ++) {
									
									ArrayList<String> mapData 			= new ArrayList<String>();
									
									
									// Loop for columns
									
									for(int colCounter = 1 ; colCounter < numOfColumns; colCounter++) {
										
											mapData.add(getCellDataString(getCellObject(WorkBookName,sheetNum,counter,  colCounter )));
										}
										
										map.put(getCellObject(WorkBookName,sheetNum,counter, 0).getStringCellValue(), mapData);
									}
									
								} catch(Exception e) {
									System.out.println(e.getMessage());
								}
							
							
							return map;
					
				}
			
				
		
		
		
		
		////////
		
		
		/// functions with various possibilities to get cell objects

		 public Cell getCellObject(String workBookName, String sheetName, int rownum, int columnum) throws Exception {
			 
				row = getRowObject(workBookName,sheetName,rownum);
				 
				 
				return row.getCell(columnum);
			 }
		 
		 public Cell getCellObject(String workBookName, int sheetNum, int rownum, int columnum) throws Exception {
			 
				row = getRowObject(workBookName,sheetNum,rownum);
				 
				 
				return row.getCell(columnum);
			 }
		 
		 
		 public Cell getCellObject(String sheetName, int rownum, int columnum) {
			 
			row = getRowObject(sheetName,rownum);
			 
			 
			return row.getCell(columnum);
		 }
		 
		 
		 public Cell getCellObject(int sheetNum, int rownum, int columnum) {
			 
				row = getRowObject(sheetNum,rownum);
				 
				 
				return row.getCell(columnum);
			 }
			 
		 
		 public Cell getCellObject(Sheet sheet, int rownum, int columnum) {
			 
				row = getRowObject(sheet,rownum);
				 
				 
				return row.getCell(columnum);
			 }
			 
		 
		 public Cell getCellObject(int rownum, int columnum) {
			 
				row = getRowObject(rownum);
				 
				 
				return row.getCell(columnum);
			 }
			 
		 
		 
		 
		 
		 
		 //////////////// functions with various possibilities to get cell objects
		 public Row getRowObject(String workBookName, String sheetName, int rownum) throws Exception {
			 
			 excelWSheetDynamic = getSheetObject(workBookName,sheetName);
			 
			 return excelWSheetDynamic.getRow(rownum);
		 }
		 
		 
		 public Row getRowObject(String workBookName, int sheetNum, int rownum) throws Exception {
			 
			 excelWSheetDynamic = getSheetObject(workBookName,sheetNum);
			 
			 return excelWSheetDynamic.getRow(rownum);
		 }
		 
		 public Row getRowObject(String sheetName, int rownum) {
			 
			 excelWSheet = getSheetObject(sheetName);
			 
			 return excelWSheet.getRow(rownum);
		 }
		 
		 
		 public Row getRowObject(int sheetNum, int rownum) {
			 
			 excelWSheet = getSheetObject(sheetNum);
			 
			 return excelWSheet.getRow(rownum);
		 }
		 
		 
		 public Row getRowObject(Sheet sheet, int rownum) {
			 
			
			 
			 return sheet.getRow(rownum);
		 }
		 
		 
		 public Row getRowObject(int rownum) {
			 
			
			 
			 return excelWSheetDefault.getRow(rownum);
		 }
		 
		 
		 /////////////////////
		 
		 
		 ////////
		 
		 
		 	// function to update the working after write operations
			public void updateFile(String status)  {
				
				if(status.equalsIgnoreCase("existing")) {
					
					
					try {
						out  = output(existence);
						
						excelWorkBook.write(out);
					} catch (FileNotFoundException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					
					
				}
				else if(status.equalsIgnoreCase("dynamic")) {
					
					try {
						out  = output("dynamic");
						
						excelWorkBookD.write(out);
					} catch (FileNotFoundException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
				
				
				else if(status.equalsIgnoreCase("newFile")) {
					
					try {
						out  = output("newFile");
						
						excelWorkBookD.write(out);
					} catch (FileNotFoundException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
				
				
				
			}
			
		 
		 //////////
		 
			//function to initialize the dynamic existence of workbook
			
			public void initDWbook(String workBookname) throws Exception {
				
				excelFileDynamic = getExcelFile(workBookname);
				
				
				excelWorkBookD = Init(excelFileDynamic,workBookname);
						
					
						this.existingDynamicFilePath = workBookname;
				
			}
		 
		 
	
}
