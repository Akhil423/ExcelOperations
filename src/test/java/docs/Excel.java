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
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class Excel {

	
	String Excelpath,xmlpath,name,ExistingFilePath;
	FileOutputStream out;
	
	XSSFWorkbook ExcelXWBook;
	HSSFWorkbook ExcelHWBook;
	File ExcelFolder,XmlFolder;
	
	
	// constructor when used existing excel file
		public Excel(String name) throws Exception {
			
			
			FileInputStream ExcelFile = getExcelPath(name);
			
			if((FilenameUtils.getExtension(name)).equalsIgnoreCase("xlsx")) {
				
				
				 ExcelXWBook = new XSSFWorkbook(ExcelFile);
				 System.out.println("workbook created");
				
			}
			else if(FilenameUtils.getExtension(name).equalsIgnoreCase("xls")){
				
					 ExcelHWBook = new HSSFWorkbook(ExcelFile);
				
			}
				
				this.ExistingFilePath = name;
			
		}
		
		
		// constructor if you want to create new workbook and do operations on it 
		public Excel() {
			// TODO Auto-generated constructor stub
		}


		// returns Input stream of particular excel
		public FileInputStream getExcelPath(String path) throws Exception{
			return new FileInputStream(new File(path));
		}
		
		
		// creates output stream of excel according file existence and vice versa
		public FileOutputStream output(boolean exist) throws FileNotFoundException {
			
			if(exist)
			return new FileOutputStream(ExistingFilePath);
			else
				return new FileOutputStream(Excelpath+File.separator+name);
		}
		
		
		//function to create new work book
		public void createXSSFBook(String name) {
			

			ExcelFolder= new File(new File("").getAbsolutePath()+File.separator+"ExcelOps");  // Folder to store excel files created dynamically
			
			XmlFolder= new File(new File("").getAbsolutePath()+File.separator+"Xml");    // Folder to store xml files created dynamically
			
			// creating folders according to their existence
			if(ExcelFolder.exists() ){
				
				System.out.println("file already exists");
			}
			else {
				
				ExcelFolder.mkdir();
			}
			
			if(XmlFolder.exists() ){
				
				System.out.println("file already exists");
			}
			else {
				
				XmlFolder.mkdir();
			}
			
			
			this.name   = name;
		    this.Excelpath  	= ExcelFolder.getPath();
		    this.xmlpath	= XmlFolder.getPath();
		    
		    
			if((FilenameUtils.getExtension(name)).equalsIgnoreCase("xlsx")) {
				
					 ExcelXWBook = new XSSFWorkbook();
					 System.out.println("workbook created");
				
			}
			else if(FilenameUtils.getExtension(name).equalsIgnoreCase("xls")){
				
					 ExcelHWBook = new HSSFWorkbook();
				
			}
		}
		
		
		// function to create XSSFSheet by passing sheet name
		public XSSFSheet CreateXSSFSheet( String name,boolean existence) {
			try {

			
					out  = output(existence);
				
			XSSFSheet ExcelWSheet = ExcelXWBook.createSheet(name);
				
			ExcelXWBook.write(out);
				return ExcelWSheet;
		
		} catch(Exception e) {
				return null;
		}
		}
		
		// function to create HSSFSheet by passing sheet name
		public HSSFSheet CreateHSSFSheet(String name,boolean existence) {
			try {
				

				out  = output(existence);
					
				HSSFSheet ExcelWSheet = ExcelHWBook.createSheet(name);
				
				ExcelHWBook.write(out);
				return ExcelWSheet;
		
		} catch(Exception e) {
				return null;
		}
		}
		
		
		
		// function to get XSSFSheet by passing sheet number
		public XSSFSheet getXSSFSheetObject(int sheetNum) {
			
			
				XSSFSheet ExcelWSheet;
				
				ExcelWSheet = ExcelXWBook.getSheetAt(sheetNum);
				
				return ExcelWSheet;
			
		}
		
		
		/** * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		 * Function 				- Get HSSF Sheet's Object
		 * @param ExcelWBook		- Excel work book
		 * @param sheet				- Sheet name(Also accept sheet number except 0)
		 * @return ExcelWSheet		- HSSF Sheet's Object
		 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */
		
		// function to get HSSFSheet by passing sheet number
		public HSSFSheet getHSSFSheetObject(int sheetNum) {
			
				HSSFSheet ExcelWSheet;
				
				ExcelWSheet = ExcelHWBook.getSheetAt(sheetNum);
				
				return ExcelWSheet;
				}
		
		// function to create rows in XSSFsheet by passing sheet 
		public void createXSSFRows(int Rows,int cols, XSSFSheet sheet,boolean existence )  {
			
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
			
			ExcelXWBook.write(out);
			out.close();
			
			}
			catch(IOException e) {
				
				e.printStackTrace();
			}
			catch(Exception e) {
				
				e.printStackTrace();
			}
			
		}
		
		// function to create rows in HSSFsheet by passing sheet 
		public void createHSSFRows(int Rows,int cols, HSSFSheet sheet, boolean existence )  {
			
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
				}
			ExcelHWBook.write(out);
			out.close();
			
			}
			catch(IOException e) {
				
				e.printStackTrace();
			}
			catch(Exception e) {
				
				e.printStackTrace();
			}
			
		}
		
		// function to fill rows in XSSFsheet by passing data Object,sheet 
		public void FillXSSFRows(Object data[][], XSSFSheet sheet, boolean existence )  {
			
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
			
			ExcelXWBook.write(out);
			out.close();
			}
			catch(Exception e) {
				
				e.printStackTrace();
			}
			
		}
		
		// function to fill rows in HSSFsheet by passing data Object,sheet 
		public void FillHSSFRows(Object[][] data, HSSFSheet sheet, boolean existence )  {
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
			
			ExcelHWBook.write(out);
			out.close();
			}
			catch(Exception e) {
						
						e.printStackTrace();
					}
				}
		
		// function to initialize Header in XSSFsheet by passing data Object,sheet 
		public void HeaderRow(Object[] data, XSSFSheet sheet) {
			
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
		
		// function to Initialize Header in XSSFsheet by passing data Object,sheet 
		public void HeaderRow(Object[] data, HSSFSheet sheet) {
			
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
		
		
		// function to fill rows in XSSFsheet by passing xml file,sheet 
		public void FillXSSFRows(String file, XSSFSheet sheet )  {
				
				try {
					
					if((FilenameUtils.getExtension(file)).equalsIgnoreCase("xml")) {
		
						
						File inputFile = new File(xmlpath+File.separator+file);
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
				
				ExcelXWBook.write(out);
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
					
					if((FilenameUtils.getExtension(name)).equalsIgnoreCase("xlsx")) {
						
							
							totRowCount 			= ExcelXWBook.getSheetAt(sheetNo).getLastRowNum();
							ExcelXWBook.close();
					}
					else if(FilenameUtils.getExtension(name).equalsIgnoreCase("xls")){
						
						totRowCount 			= ExcelHWBook.getSheetAt(sheetNo).getLastRowNum();
						ExcelHWBook.close();
					}
					
				} catch(Exception e) {
					throw new Exception("Unable to get row count");
				}
				
					return totRowCount;
				}
		
		// function to get row number of XSSFSheet value by passing sheet and value
		@SuppressWarnings("deprecation")
		public int getRowNumber(XSSFSheet sheet, String searchValue) {
				
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
		
		// function to get row number of HSSFSheet value by passing sheet and value
		@SuppressWarnings("deprecation")
		public int getRowNumber(HSSFSheet sheet, String searchValue) {
				
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
		
		// function to set value in particular cell of XSSFSheet by passing sheet,row number, column number, value
		 public void setCellData(XSSFSheet sheet, int rownum, int columnum, String value, boolean existence)  {
			
				try {
				
					out  = output(existence);	
					// Get Row Number
			        Row row 						= sheet.getRow(rownum);
			        Cell cell 						= row.getCell(columnum);  
			        
			        cell.setCellValue(value);
			    	
			        ExcelXWBook.write(out);
			        out.close();
					
				} 
				catch(IOException e) {
					
					e.printStackTrace();
				}
				catch (Exception e) {
					e.printStackTrace();
				}
				
				
			}
		 
		// function to set value in particular cell of HSSFSheet by passing sheet,row number, column number, value
		 public void setCellData(HSSFSheet sheet, int rownum, int columnum, String value, boolean existence)  {
				
				try {
				
					out  = output(existence);
			        
			       																									// Get Row Number
			        Row row 						= sheet.getRow(rownum);
			        Cell cell 						= row.getCell(columnum);  
			        cell.setCellValue(value);
			    
			    	
			        ExcelHWBook.write(out);
			        out.close();
					
				} 
				catch(IOException e) {
					
					e.printStackTrace();
				}
				catch (Exception e) {
					e.printStackTrace();
				}
				
				
			}


		// function to get value in particular cell of XSSFSheet by passing sheet,row number, column number
		public String getCellData(XSSFSheet sheet,int rownum, int columnum, boolean existence) {
				
				String value  = "";
				try {
				
					out  = output(existence);
																					// Get Row Number
			        Row row 						= sheet.getRow(rownum);
			        Cell cell 						= row.getCell(columnum);
			        
			        value  =  cell.getStringCellValue();
			    	
			        
			        ExcelXWBook.write(out);
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
		
		// function to get value in particular cell of HSSFSheet by passing sheet,row number, column number
		public String getCellData(HSSFSheet sheet,int rownum, int columnum, boolean existence) {
			
			String value  = "";
			try {
			
				out  = output(existence);
																				// Get Row Number
		        Row row 						= sheet.getRow(rownum);
		        Cell cell 						= row.getCell(columnum);
		        
		        value  =  cell.getStringCellValue();
		        
		    	
		        
		        ExcelXWBook.write(out);
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
				
					numOfSheets 			= ExcelXWBook.getNumberOfSheets();
			
				} else if((FilenameUtils.getExtension(Excelpath)).equalsIgnoreCase("xls")) {
					
					numOfSheets 			= ExcelHWBook.getNumberOfSheets();
				}
				
				
				return numOfSheets;
		}
	
	
		// function to get complete row in XSSFsheet by passing sheet, row number 
		public ArrayList<String> getRow(XSSFSheet sheet, int rownum){
		
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
	
		// function to get complete row in HSSFsheet by passing sheet, row number 
		public ArrayList<String> getRow(HSSFSheet sheet, int rownum){
			
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
				
				if(FilenameUtils.getExtension(name).equalsIgnoreCase("xlsx")) {
					
					try {
						
						// Create object for sheet
						
						XSSFSheet Sheet 					= getXSSFSheetObject(sheetNum);
						Row row									= null;
						Cell cell 								= null;
						
						// Get total number of used columns
						
						int numOfColumns 						= Sheet.getRow(0).getLastCellNum();
						
						// Loop for rows starting from 1 as 0 is header of row
						
						for(int counter = 1; counter <= Sheet.getLastRowNum(); counter ++) {
							
							ArrayList<String> mapData 			= new ArrayList<String>();
							row 								= Sheet.getRow(counter);
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
					}
					
					return map;
			
		}
	
	
		
		
	
	
}
