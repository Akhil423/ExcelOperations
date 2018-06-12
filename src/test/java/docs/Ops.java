package docs;



public class Ops {

	
	public static void main(String args[]) throws Exception {
		
		Excel obj   = new Excel();
		
		obj.createBook("akhil.xlsx");
		
		Object data[][]= {{"rollno","firstname","lastname","nickname","marks"}};
		
		obj.createSheet("sheet1");
		
		obj.createSheet("sheet2");
		
		obj.fillRows(data, obj.getSheetObject(1));
		
		
		

	}
}
