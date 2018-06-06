package xlsx;



public class Ops {

	
	public static void main(String args[]) throws Exception {
		
		Excel obj   = new Excel("book43.xlsx");
		Object data[][]= {{"Hi","Hello",9}};
		
		obj.createXSSFRows(5, data, obj.CreateXSSFSheet("sheet1"));
		
		obj.createXSSFRows(5, data,obj.CreateXSSFSheet("sheet2"));
		obj.createXSSFRows(5, data,obj.CreateXSSFSheet("sheet3"));
		
		System.out.println(obj.getRowCount("book43.xlsx", 0));
		System.out.println(obj.getNumberOfSheets("book43.xlsx"));
		
		
	}
}
