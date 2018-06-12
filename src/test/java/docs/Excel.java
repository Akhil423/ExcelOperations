package docs;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;

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

	
	String excelFolderPath,xmlFolderPath,name,existingFilePath;
	FileOutputStream out;
	
	Workbook excelWorkBook;
	
	File excelFolder,xmlFolder;
	boolean existence= false;
	
	// constructor when used existing excel file
		public Excel(String FilePath) throws Exception {
			
			
			FileInputStream ExcelFile = getExcelFile(FilePath);
			
			if((FilenameUtils.getExtension(FilePath)).equalsIgnoreCase("xlsx")) {
				
				
				 excelWorkBook = new XSSFWorkbook(ExcelFile);
				 System.out.println("workbook created");
				
			}
			else if(FilenameUtils.getExtension(FilePath).equalsIgnoreCase("xls")){
				
				excelWorkBook = new HSSFWorkbook(ExcelFile);
				
			}
				
				this.existingFilePath = FilePath;
				this.existence = true;
			
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
		public FileOutputStream output(boolean exist) throws FileNotFoundException {
			
			if(exist)
			return new FileOutputStream(existingFilePath);
			else
				return new FileOutputStream(excelFolderPath+File.separator+name);
		}
		
		
		//function to create new work book
		public void createBook(String name) {
			

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
			
			
			this.name   = name;
		    this.excelFolderPath  	= excelFolder.getPath();
		    this.xmlFolderPath	= xmlFolder.getPath();
		    
		    
			if((FilenameUtils.getExtension(name)).equalsIgnoreCase("xlsx")) {
				
					 excelWorkBook = new XSSFWorkbook();
					 System.out.println("workbook created");
				
			}
			else if(FilenameUtils.getExtension(name).equalsIgnoreCase("xls")){
				
					 excelWorkBook = new HSSFWorkbook();
				
			}
		}
		
		
		// function to create Sheet by passing sheet name
		public Sheet createSheet( String name) {
			try {

			
					out  = output(existence);
				
			Sheet ExcelWSheet = excelWorkBook.createSheet(name);
				
			excelWorkBook.write(out);
				return ExcelWSheet;
		
		} catch(Exception e) {
				return null;
		}
		}

		
		// function to get Sheet by passing sheet number
		public Sheet getSheetObject(int sheetNum) {
			
			
				Sheet ExcelWSheet;
				
				ExcelWSheet = excelWorkBook.getSheetAt(sheetNum);
				
				return ExcelWSheet;
			
		}
		
		
		
		// function to create rows in sheet by passing sheet 
		public void createRows(int Rows,int cols, Sheet sheet )  {
			
			try {
		
				out  = output(existence);
				
			 int rowNum = 0,Columns = cols,columnNum;
			 
			 try {
				 if(sheet.getLastRowNum()!= 0)
					 rowNum=sheet.getLastRowNum();
				 
				 if(sheet.getRow(0).getLastCellNum()!= 0)
				 Columns=sheet.getRow(0).getLastCellNum();
				 
				 columnNum=0;
			 }catch(Exception e) {
				 
				 e.printStackTrace();
			 }
			Rows= rowNum+Rows;
			
			for(int row=rowNum;row<Rows;++row){
				
				Row roow=sheet.createRow(row);
				
				for(columnNum=0;columnNum<Columns;columnNum++) {
					
					roow.createCell(columnNum);
		           
				}
				
				System.out.println(roow.getPhysicalNumberOfCells());
			}
			
			excelWorkBook.write(out);
			out.close();
			
			}
			catch(IOException e) {
				
				e.printStackTrace();
			}
			catch(Exception e) {
				
				e.printStackTrace();
			}
			
		}
		

		
		// function to fill rows in sheet by passing data Object,sheet 
		public void fillRows(Object data[][], Sheet sheet )  {
			
			try {
				
				out  = output(existence);
				
			int rowNum=0,columnNum=0;
			for(Object[] row:data){
				Row roow=sheet.createRow(rowNum++);
				for(Object field:row) {
					Cell cell = roow.createCell(columnNum++);
		            if (field instanceof String) {
		                cell.setCellValue((String) field);
		            } else if (field instanceof Integer) {
		                cell.setCellValue((Integer) field);
		            }
				}
			}
			
			excelWorkBook.write(out);
			out.close();
			}
			catch(Exception e) {
				
				e.printStackTrace();
			}
			
		}
		
		
		// function to initialize Header in sheet by passing data Object,sheet 
		public void headerRow(Object[] data, Sheet sheet) {
			
		        Row row = sheet.getRow(0);
		        int columns = 0;
		        for(Object field:data) {
		        	
		        	Cell cell= row.getCell(columns++);
		        	
		        //	System.out.println(cell.getCellType());
		        	
		        	if (field instanceof String) {
		                cell.setCellValue((String) field);
		            } else if (field instanceof Integer) {
		                cell.setCellValue((Integer) field);
		            }
		        	
		        }
		        
		     }

	
		
		// function to fill rows in sheet by passing xml file,sheet 
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
		
		 // function to get row count by passing sheet number
		public int getRowCount(int sheetNo) throws Exception {
			
				int totRowCount 					= 0;
				
				try {
					
						
							
							totRowCount 			= excelWorkBook.getSheetAt(sheetNo).getLastRowNum();
							excelWorkBook.close();
		
				
					
				} catch(Exception e) {
					throw new Exception("Unable to get row count");
				}
				
					return totRowCount;
				}
		
		// function to get row number of Sheet value by passing sheet and value
		@SuppressWarnings("deprecation")
		public int getRowNumber(Sheet sheet, String searchValue) {
				
				try {
					
					for (Row row : sheet) 
						for (Cell cell : row) 
							if (cell.getCellType() == Cell.CELL_TYPE_STRING) 
								if (cell.getRichStringCellValue().getString().trim().equalsIgnoreCase(searchValue)) 
									return row.getRowNum();
					
				} catch(Exception e) {
					e.printStackTrace();
				}
				
				return 0;
			}
		

		
		// function to set value in particular cell of Sheet by passing sheet,row number, column number, value
		 public void setCellData(Sheet sheet, int rownum, int columnum, String value)  {
			
				try {
				
					out  = output(existence);	
					// Get Row Number
			        Row row 						= sheet.getRow(rownum);
			        Cell cell 						= row.getCell(columnum);  
			        
			        cell.setCellValue(value);
			    	
			        excelWorkBook.write(out);
			        out.close();
					
				} 
				catch(IOException e) {
					
					e.printStackTrace();
				}
				catch (Exception e) {
					e.printStackTrace();
				}
				
				
			}
		 


		// function to get value in particular cell of Sheet by passing sheet,row number, column number
		public String getCellData(Sheet sheet,int rownum, int columnum) {
				
				String value  = "";
				try {
				
					out  = output(existence);
																					// Get Row Number
			        Row row 						= sheet.getRow(rownum);
			        Cell cell 						= row.getCell(columnum);
			        
			        value  =  cell.getStringCellValue();
			    	
			        
			        excelWorkBook.write(out);
			        out.close();
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

	
	
		// function to get number of sheets present in particular workbook
		public int getNumberOfSheets(String WorkBookName) {
		
				int numOfSheets 			= 0;
			
				if((FilenameUtils.getExtension(WorkBookName)).equalsIgnoreCase("xlsx")) {
				
					numOfSheets 			= excelWorkBook.getNumberOfSheets();
			
				} else if((FilenameUtils.getExtension(excelFolderPath)).equalsIgnoreCase("xls")) {
					
					numOfSheets 			= excelWorkBook.getNumberOfSheets();
				}
				
				
				return numOfSheets;
		}
	
	
		// function to get complete row in sheet by passing sheet, row number 
		public ArrayList<String> getRow(Sheet sheet, int rownum){
		
				ArrayList<String> list = new ArrayList<String>();
				
				Row row  = sheet.getRow(rownum);
				int numofColumns = sheet.getRow(0).getLastCellNum();
				String val = "";
				for(int colCounter = 0 ; colCounter < numofColumns; colCounter++) {
					
					Cell cell 							= row.getCell(colCounter);
					
					if(cell.getCellType()==0) {
						val 						= Double.toString(cell.getNumericCellValue());
			    	} else if(cell.getCellType()==1) {
			    		val 						= cell.getStringCellValue();
			    	} else if (cell.getCellType()==3) {
			    		val 						= "";
			    	}
					
					 list.add(val);
				}
			
			return list;
		}
	
	
		
		// function to get sheet data in hash map by passing sheet number
		public LinkedHashMap<String, ArrayList<String>> getSheetData( int sheetNum) {
		
				LinkedHashMap<String, ArrayList<String>> map 	= new LinkedHashMap<String, ArrayList<String>>();
				
					
					try {
						
						// Create object for sheet
						
						Sheet sheet 					= getSheetObject(sheetNum);
						Row row									= null;
						Cell cell 								= null;
						
						// Get total number of used columns
						
						int numOfColumns 						= sheet.getRow(0).getLastCellNum();
				
						// Loop for rows starting from 1 as 0 is header of row
						
						for(int counter = 1; counter <= sheet.getLastRowNum(); counter ++) {
							
							ArrayList<String> mapData 			= new ArrayList<String>();
							row 								= sheet.getRow(counter);
							String val 							= "";
							
							// Loop for columns
							
							for(int colCounter = 1 ; colCounter < numOfColumns; colCounter++) {
								
								cell 							= row.getCell(colCounter);
								
								if(cell.getCellType()==0) {
									val 						= Double.toString(cell.getNumericCellValue());
				            	} else if(cell.getCellType()==1) {
				            		val 						= cell.getStringCellValue();
				            	} else if (cell.getCellType()==3) {
				            		val 						= "";
					            	}
									
									mapData.add(val);
								}
								
								map.put(row.getCell(0).getStringCellValue(), mapData);
							}
							
						} catch(Exception e) {
							System.out.println(e.getMessage());
						}
					
					
					return map;
			
		}
	
	
}
