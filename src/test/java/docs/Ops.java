package docs;



public class Ops {

	
	public static void main(String args[]) throws Exception {
		
		Excel obj   = new Excel("D:\\ExcelOperations\\ExcelOps\\bookxml.xlsx");
		
		Object data[][]= {{"rollno","firstname","lastname","nickname","marks"}};
		
		obj.CreateXSSFSheet("sheet2",true);
		
		obj.CreateXSSFSheet("sheet6", true);
		
		obj.FillXSSFRows(data, obj.getXSSFSheetObject(2), true);
		
		

		
		
	}
}
