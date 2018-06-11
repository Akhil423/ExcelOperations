package xlsx;



public class Ops {

	
	public static void main(String args[]) throws Exception {
		
		Excel obj   = new Excel("bookxml.xlsx");
		Object data[][]= {{"rollno","firstname","lastname","nickname","marks"}};
		
		Object data1[]= {"rollno","firstname","lastname","nickname","marks"};
		//obj.createXSSFRows(5, data, obj.CreateXSSFSheet("sheet1"));
		
		/*obj.createXSSFRows(5, data,obj.CreateXSSFSheet("sheet2"));
		obj.createXSSFRows(5, data,obj.CreateXSSFSheet("sheet3"));*/
		obj.CreateXSSFSheet("sheet1");
		obj.createXSSFRows(8,5,obj.getXSSFSheetObject("sheet1") );
		
		/*System.out.println(obj.getRowCount("bookxml.xlsx", 0));
		System.out.println(obj.getNumberOfSheets("bookxml.xlsx"));
		*/
		
		obj.HeaderRow(data1, obj.getXSSFSheetObject("sheet1"));
		
		obj.createXSSFRows("hello.xml", obj.getXSSFSheetObject("sheet1"));
		
		obj.setCellData("bookxml", 4, 3, "hello man");
		
		
	}
}
