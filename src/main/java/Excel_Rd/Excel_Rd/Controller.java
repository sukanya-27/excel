package Excel_Rd.Excel_Rd;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Iterator;


public class Controller {
	 public static void main(String[] args) {
		
	
     
		 int batchSize = 20;
     try{
    	
     	long start = System.currentTimeMillis();
     FileInputStream file = new FileInputStream(new File("student1.xlsx")); 
     // Create Workbook instance holding reference to .xlsx file 
     XSSFWorkbook workbook = new XSSFWorkbook(file);  
     // Get first/desired sheet from the workbook 
     XSSFSheet sheet = workbook.getSheetAt(0); 
     Iterator<Row> rowIterator = sheet.iterator(); 
     try {
			Class.forName("com.mysql.jdbc.Driver");
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
     Connection connection=DriverManager.getConnection("jdbc:mysql://localhost:3306/excel", "root", "root");
     connection.setAutoCommit(false);

     String sql = "INSERT INTO data1 (sid, sname, sloc) VALUES (?, ?, ?)";
           PreparedStatement statement = connection.prepareStatement(sql);    


     int count = 0;
     rowIterator.next();
     while (rowIterator.hasNext()) { 
         Row row = rowIterator.next();
         
             // For each row, iterate through all the columns 
             Iterator<Cell> cellIterator = row.cellIterator(); 

             while (cellIterator.hasNext()) { 
                 Cell cell = cellIterator.next(); 
                 int columnIndex =cell.getColumnIndex();
                // CellType b= cell.getCellType();
				// Check the cell type and format accordingly           
                 switch (cell.getColumnIndex()){
                 case 0:
                     int sid = (int)cell.getNumericCellValue();
                     statement.setInt(1,sid);
                     break;
                 case 1:
                     String sname = cell.getStringCellValue();
                     statement.setString(2,sname);
                 case 2:
                     String sloc = cell.getStringCellValue();
                     statement.setString(3,sloc);
                 }

             }
              
             statement.addBatch();
              
             if (count % batchSize == 0) {
                 statement.executeBatch();
             }              

         }

     
     
     workbook.close();
     statement.executeBatch();
     
     connection.commit();
     connection.close();
     long end = System.currentTimeMillis();
     System.out.printf("Import done in %d ms\n", (end - start));
     }catch(IOException e){
     	System.out.println("error in read file");
     	e.printStackTrace();
     	}catch(SQLException e1){
     		System.out.println("database error");
     		e1.printStackTrace();
     	}
}
}


