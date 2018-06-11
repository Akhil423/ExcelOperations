package xlsx;

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

	
	String path,path1,name;
	FileOutputStream out;
	
	XSSFWorkbook ExcelXWBook;
	HSSFWorkbook ExcelHWBook;
	File ExcelFolder,XmlFolder;
	
		public Excel(String name) {
			
			
			ExcelFolder= new File(new File("").getAbsolutePath()+File.separator+"ExcelOps");
			XmlFolder= new File(new File("").getAbsolutePath()+File.separator+"Xml");
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
		    this.path  	= ExcelFolder.getPath();
		    this.path1	= XmlFolder.getPath();
		    
		    
			if((FilenameUtils.getExtension(name)).equalsIgnoreCase("xlsx")) {
				
					 ExcelXWBook = new XSSFWorkbook();
					 System.out.println("workbook created");
				
			}
			else if(FilenameUtils.getExtension(name).equalsIgnoreCase("xls")){
				
					 ExcelHWBook = new HSSFWorkbook();
				
			}
			
		}
		
		
		public XSSFSheet CreateXSSFSheet( String name) {
			try {
				
			XSSFSheet ExcelWSheet = ExcelXWBook.createSheet(name);
				
				return ExcelWSheet;
		
		} catch(Exception e) {
				return null;
		}
		}
		
		public HSSFSheet CreateHSSFSheet(String name) {
			try {
				
				HSSFSheet ExcelWSheet = ExcelHWBook.createSheet(name);
				
				return ExcelWSheet;
		
		} catch(Exception e) {
				return null;
		}
		}
		
		
		public FileOutputStream output() throws FileNotFoundException {
			
			return new FileOutputStream(path+File.separator+name);
		}
		
		/** * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		 * Function 				- Get XSSF Sheet's Object
		 * @param ExcelWBook		- Excel work book
		 * @param sheet				- Sheet name(Also accept sheet number except 0)
		 * @return ExcelWSheet		- XSSF Sheet's Object
		 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */
		
		public XSSFSheet getXSSFSheetObject(String sheet) {
			
			
			
			try{
			
				XSSFSheet ExcelWSheet;
				
				ExcelWSheet = ExcelXWBook.getSheet(sheet);
				
			System.out.println(ExcelWSheet.toString());
				return ExcelWSheet;
				
			} catch(Exception e) {
				return null;
			}
		}
		
		
		/** * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
		 * Function 				- Get HSSF Sheet's Object
		 * @param ExcelWBook		- Excel work book
		 * @param sheet				- Sheet name(Also accept sheet number except 0)
		 * @return ExcelWSheet		- HSSF Sheet's Object
		 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */
		
		public HSSFSheet getHSSFSheetObject(String sheet) {
			
				HSSFSheet ExcelWSheet;
				
					ExcelWSheet 	= ExcelHWBook.getSheet(sheet);
				
				return ExcelWSheet;
				}
		
		public void createXSSFRows(int Rows,int cols, XSSFSheet sheet )  {
			
			try {
				
				
			 out  = output();
			 
			 int rowNum = 0,Columns = cols,columnNum;
			 
			 try {
			//rowNum=sheet.getLastRowNum();Columns=sheet.getRow(0).getLastCellNum();columnNum=0;
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
		
		public void createHSSFRows(int Rows, HSSFSheet sheet )  {
			
			try {
				
				
			 out  = output();
			int rowNum=sheet.getLastRowNum(),Columns=sheet.getRow(0).getLastCellNum(),columnNum=0;
			Rows= rowNum+Rows;
			
			for(int row=rowNum;row<Rows;++row){
				
				Row roow=sheet.createRow(row);
				
				for(columnNum=0;columnNum<Columns;columnNum++) {
					
					Cell cell = roow.createCell(columnNum);
					//cell.setCellValue("test");
		           
				}
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
		
		public void createXSSFRows(int Rows,Object data[][], XSSFSheet sheet )  {
			
			try {
				
				
			 out  = output();
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
			catch(IOException e) {
				
				e.printStackTrace();
			}
			catch(Exception e) {
				
				e.printStackTrace();
			}
			
		}
		
		public void createHSSFRows(int Rows, int Cols, Object[][] data, HSSFSheet sheet )  {
			try {
				
				 out  = output();
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
			catch(IOException e) {
						
				e.printStackTrace();
					}
					catch(Exception e) {
						
						e.printStackTrace();
					}
				}
		
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
		
		
		
		public void createXSSFRows(String file, XSSFSheet sheet )  {
				
				try {
					
					if((FilenameUtils.getExtension(file)).equalsIgnoreCase("xml")) {
		
						
						File inputFile = new File(path1+File.separator+file);
				         DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
				         DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
				         Document doc = dBuilder.parse(inputFile);
				         doc.getDocumentElement().normalize();
				         System.out.println("Root element :" + doc.getDocumentElement().getNodeName());
				         NodeList nList = doc.getElementsByTagName("student");
				         
				         
						out  = output();
				
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
		
		public int getRowCount(String name, int sheetNo) throws Exception {
			
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
		
		
		 public void setCellData(String FileName, int rownum, int columnum, String value)  {
			
				try {
				
			        // Open work book
			        XSSFSheet sheet 				= ExcelXWBook.getSheetAt(0);							// Open work sheet 																					// Get Row Number
			        Row row 						= sheet.getRow(rownum);
			        Cell cell 						= row.getCell(columnum);  
			        cell.setCellValue(value);
			        
			        out						= output();
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


		public String getCellData(int rownum, int columnum) {
				
				String value  = "";
				try {
				
				
					// Open work book
					
			        XSSFSheet sheet 				= ExcelXWBook.getSheetAt(0);							// Open work sheet 																					// Get Row Number
			        Row row 						= sheet.getRow(rownum);
			        Cell cell 						= row.getCell(columnum);
			        
			        value  =  cell.getStringCellValue();
			        
			        out							= output();
			        
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
	
	
	
	
		public int getNumberOfSheets(String WorkBookName) {
		
				int numOfSheets 			= 0;
			
				if((FilenameUtils.getExtension(WorkBookName)).equalsIgnoreCase("xlsx")) {
				
					numOfSheets 			= ExcelXWBook.getNumberOfSheets();
			
				} else if((FilenameUtils.getExtension(path)).equalsIgnoreCase("xls")) {
					
					numOfSheets 			= ExcelHWBook.getNumberOfSheets();
				}
				
				
				return numOfSheets;
		}
	
	
	
		public ArrayList<String> getRow(XSSFSheet sheetName, int rownum){
		
				ArrayList<String> list = new ArrayList<String>();
				
				Row row  = sheetName.getRow(rownum);
				int numofColumns = sheetName.getRow(0).getLastCellNum();
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
	
		public LinkedHashMap<String, ArrayList<String>> getSheetData( String sheetNum) {
		
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
