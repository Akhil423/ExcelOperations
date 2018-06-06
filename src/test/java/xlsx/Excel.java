package xlsx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	
	String path,name;
	FileOutputStream out;
	
	XSSFWorkbook ExcelXWBook;
	HSSFWorkbook ExcelHWBook;
	File f;
	
	public Excel(String name) {
		
		
		f= new File(new File("").getAbsolutePath()+File.separator+"ExcelOps");
		
		if(f.exists()){
			
			System.out.println("file already exists");
		}
		else {
			
			f.mkdir();
		}
		
		
		this.name   = name;
        this.path  	= f.getPath();
        
        
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
	
	public void createXSSFRows(int Rows,Object data[][], XSSFSheet sheet )  {
		
		try {
			
			
		 out  = new FileOutputStream(path+File.separator+name);
		int rowNum=0,columnNum=0;
		for(Object[] row:data){
			Row rOw=sheet.createRow(++rowNum);
			for(Object field:row) {
				Cell cell = rOw.createCell(++columnNum);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
			}
		}
		
		ExcelXWBook.write(out);
		
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
			
			 out  = new FileOutputStream(path+File.separator+name);
		int rowNum=0,columnNum=0;
		for(Object[] row:data){
			Row rOw=sheet.createRow(rowNum++);
			for(Object field:row) {
				
				Cell cell = rOw.createCell(columnNum++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
			}
		}
		
		ExcelHWBook.write(out);
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
		
		String FileName= path+File.separator+name;
		
		try {
			
			if((FilenameUtils.getExtension(FileName)).equalsIgnoreCase("xlsx")) {
				
					FileInputStream file1 	= new FileInputStream(new File(FileName));
					XSSFWorkbook ExcelWBook = new XSSFWorkbook(file1);
					totRowCount 			= ExcelWBook.getSheetAt(sheetNo).getLastRowNum();
					ExcelWBook.close();
			}
			else if(FilenameUtils.getExtension(FileName).equalsIgnoreCase("xls")){
				
				FileInputStream file1 	= new FileInputStream(new File(FileName));
				HSSFWorkbook ExcelWBook = new HSSFWorkbook(file1);
				totRowCount 			= ExcelWBook.getSheetAt(sheetNo).getLastRowNum();
				ExcelWBook.close();
			}
			
		} catch(Exception e) {
			throw new Exception("Unable to get row count");
		}
		
		return totRowCount;
	}
	
	@SuppressWarnings("deprecation")
	private int getRowNumber(XSSFSheet sheet, String searchValue) {
			
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
	public void setCellData(String FileName, String rowName, String columnName, String value)  {
		
		try {
		
			int rowNumber					= 0;
			int columnNumber 				= 0;
			FileOutputStream fos 			= null;
	        InputStream input 				= new FileInputStream(path + File.separator + FileName+ ".xlsx");

	        @SuppressWarnings("resource")
			XSSFWorkbook workbook 			= new XSSFWorkbook(input);							// Open work book
	        XSSFSheet sheet 				= workbook.getSheetAt(0);							// Open work sheet
	        rowNumber 						= getRowNumber(sheet, rowName);						// Get Row Number
	        Row headerRow 					= sheet.getRow(0);
	        Row row 						= sheet.getRow(rowNumber);
	        
	      //  System.out.println(row );
	        
	        for(int i = 0; i < headerRow.getLastCellNum(); i++) {
	        	
	            if(headerRow.getCell(i).getStringCellValue().trim().equals(columnName))
	                columnNumber 			= i;
	            
	        }
	       // System.out.println(columnNumber);
	        Cell cell 						= row.getCell(columnNumber);
	        
	        cell.setCellValue(value);
	        
	        fos 							= new FileOutputStream(path+ File.separator + FileName + ".xlsx");
	        
	        workbook.write(fos);
	        fos.close();
			
		} 
		catch(IOException e) {
			
			e.printStackTrace();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	
	public void setCellData(String FileName, int rownum, int columnum, String value)  {
		
		try {
		
			
			FileOutputStream fos 			= null;
	        InputStream input 				= new FileInputStream(path + File.separator + FileName+ ".xlsx");

	        @SuppressWarnings("resource")
			XSSFWorkbook workbook 			= new XSSFWorkbook(input);							// Open work book
	        XSSFSheet sheet 				= workbook.getSheetAt(0);							// Open work sheet 																					// Get Row Number
	        Row headerRow 					= sheet.getRow(0);
	        Row row 						= sheet.getRow(rownum);
	        
	      //  System.out.println(row );
	        
	      /*  for(int i = 0; i < headerRow.getLastCellNum(); i++) {
	        	
	            if(headerRow.getCell(i).getStringCellValue().trim().equals(columnName))
	                columnNumber 			= i;
	            
	        }
	       // System.out.println(columnNumber);
*/	        Cell cell 						= row.getCell(columnum);
	        
	        cell.setCellValue(value);
	        
	        fos 							= new FileOutputStream(path + File.separator + FileName + ".xlsx");
	        
	        workbook.write(fos);
	        fos.close();
			
		} 
		catch(IOException e) {
			
			e.printStackTrace();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		
		
	}
	
	
	public String getCellData(String FileName, int rownum, int columnum) {
		
		String value  = "";
		try {
		
			
			FileOutputStream fos 			= null;
	        InputStream input 				= new FileInputStream(path + File.separator + FileName+ ".xlsx");

	        @SuppressWarnings("resource")
			XSSFWorkbook workbook 			= new XSSFWorkbook(input);							// Open work book
	        XSSFSheet sheet 				= workbook.getSheetAt(0);							// Open work sheet 																					// Get Row Number
	        Row headerRow 					= sheet.getRow(0);
	        Row row 						= sheet.getRow(rownum);
	        
	   
	        Cell cell 						= row.getCell(columnum);
	        
	       value  =  cell.getStringCellValue();
	        
	        fos 							= new FileOutputStream(path+ File.separator + FileName + ".xlsx");
	        
	        workbook.write(fos);
	        fos.close();
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
	
	
	
	
	public int getNumberOfSheets(String name) throws IOException {
		
		int numOfSheets 			= 0;
		try {
		if((FilenameUtils.getExtension(name)).equalsIgnoreCase("xlsx")) {
			
			FileInputStream file1 	= new FileInputStream(path+File.separator+name);
			
			//System.out.println(new File("").getAbsolutePath()+File.separator+path);
			XSSFWorkbook WRKOBJ 	= new XSSFWorkbook(file1);
		
			numOfSheets 			= WRKOBJ.getNumberOfSheets();

		} else if((FilenameUtils.getExtension(path)).equalsIgnoreCase("xls")) {
			
			FileInputStream file1 	= new FileInputStream(path+File.separator+name);
			
			
			HSSFWorkbook WRKOBJ1 	= new HSSFWorkbook(file1);
			numOfSheets 			= WRKOBJ1.getNumberOfSheets();
		}
		
		}
		catch(IOException e) {
			
			e.printStackTrace();
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		
		
		return numOfSheets;
	}


	private XSSFWorkbook XSSFWorkbook(FileInputStream file1) {
		// TODO Auto-generated method stub
		
		return null;
	}
	
		
	
	
}
