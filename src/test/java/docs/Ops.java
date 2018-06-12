package docs;



public class Ops {

	
	public static void main(String args[]) throws Exception {
		
		Excel obj   = new Excel("D:\\ExcelOperations\\ExcelOps\\bookxml.xlsx");
		
		Object data[][]= {{"rollno","firstname","lastname","nickname","marks"}};
		
		obj.CreateSheet("sheet8");
		
		obj.CreateSheet("sheet9");
		
		obj.fillRows(data, obj.getSheetObject(2));
		
		
		

	}
}
