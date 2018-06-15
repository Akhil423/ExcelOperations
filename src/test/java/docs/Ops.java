package docs;

import java.io.File;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;

public class Ops {

	
	public static void main(String args[]) throws Exception {
		
		
		Object data[][]= {{"rollno","firstname","lastname","nickname","marks"},{"1","akhil","varma","akhi","34"},{"2","akhil","varma","akhi","38"}};
		
		
		String filePath="D:\\ExcelOperations\\ExcelOps\\akhil.xlsx";
		Excel obj   = new Excel(filePath);
		System.out.println(obj.getCellData(filePath, "sheet2", 1, 1));
		
		System.out.println(obj.getCellData("sheet2", 1, 1));
		
		System.out.println(obj.getCellData(1, 1, 1));
		
		obj.setSheetDefault("sheet2");
		
		System.out.println(obj.getCellData(1, 1));
	
		
		obj.createBook("varma89.xlsx");
		
		obj.createSheet(filePath, "sheet4");
		
		obj.insertData(data,filePath,"sheet4");
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		/*Object data[][]= {{"rollno","firstname","lastname","nickname","marks"},{"1","akhil","varma","akhi","34"}};
		
		obj.insertData(data, "sheet2");
	
		obj.getCellData(1, 1, 1);*/
		
		
		
		

	}
}
