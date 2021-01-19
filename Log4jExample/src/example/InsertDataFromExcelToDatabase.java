package example;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.*;
public class InsertDataFromExcelToDatabase extends InsertDataFromExcelToDb {
	// jdbc connection with postgresql
	static String  JdbcConnection(String data1,String data2,String data3,String data4) {
		try {
			log.debug(data1+" "+data2+" "+data3+" "+data4);
			Class.forName("org.postgresql.Driver");
			Connection con=DriverManager.getConnection("jdbc:postgresql://localhost:5433/postgres","shanmukha","asdfghjkl");
			String sql_insert="insert into insertdata values(?,?,?,?)";
			PreparedStatement pstmt=con.prepareStatement(sql_insert);
			pstmt.setString(1,data1);
			pstmt.setString(2, data2);
			pstmt.setString(3, data3);
			pstmt.setString(4, data4);
			int j=pstmt.executeUpdate();
			con.close();
			log.trace("data inserted");
			
		}
		catch(Exception e) {
			System.out.println(e.getMessage());
		}
		return null;
	}
	//get data from the database table  and insert the data into the excel file
	static void CreateExcelFunction() {
		try   
		{  
			log.trace("getting data from the table and insert the data into the excel file ");
		//String filename1 = "C:\\Users\\varma\\Desktop\\DataInsert.xlsx";  
		FileOutputStream fileOut = new FileOutputStream("C:\\Users\\varma\\Desktop\\DataInsert.xlsx");  
		fileOut.close();  
		System.out.println("Excel file has been generated successfully.");  
		//String filename = "C:\\Users\\varma\\Desktop\\output.xlsx"; 
		Class.forName("org.postgresql.Driver");
		Connection con=DriverManager.getConnection("jdbc:postgresql://localhost:5433/postgres","shanmukha","asdfghjkl");
		Statement stmt = con.createStatement();
		ResultSet data=stmt.executeQuery("select * from insertdata");
		HSSFWorkbook workbook = new HSSFWorkbook(); 
		int i=0;
		HSSFSheet sheet = workbook.createSheet("Excel data");  
		while(data.next()) { 
		HSSFRow row = sheet.createRow((int)i);  
		row.createCell(0).setCellValue(data.getString(1)); 
		row.createCell(1).setCellValue(data.getString(2));  
		row.createCell(2).setCellValue(data.getString(3));  
		row.createCell(3).setCellValue(data.getString(4));  
		log.debug(data.getString(1)+" "+data.getString(2)+" "+data.getString(3)+" "+data.getString(4));
	i++;
		log.debug("data inserted ");
	    }
		
		workbook.write(fileOut);  
	 
		fileOut.close();  
 
		workbook.close();  
		
		}   
		catch (Exception e)   
		{ 
		 }

	}
	private static org.apache.log4j.Logger log=Logger.getLogger(InsertDataFromExcelToDatabase.class.getName());
	public static void main(String[] a){
		InsertDataFromExcelToDatabase insert_data_from_excel_to_database=new InsertDataFromExcelToDatabase();
		try {
			FileInputStream flp=new FileInputStream("C:\\Users\\Varma\\Desktop\\EmployeeData.xlsx");
			XSSFWorkbook xlsx_work_book=new XSSFWorkbook(flp);
			BasicConfigurator.configure();
			log.info("xlsx read sucessfully");
			XSSFSheet sheet_data=xlsx_work_book.getSheetAt(0);
			Iterator<Row> itr=sheet_data.iterator();
			while(itr.hasNext()) {
				Row row=itr.next();
				int i=0;
				Iterator<Cell> cellIterator=row.cellIterator();
				while(cellIterator.hasNext()) {
					Cell cell=cellIterator.next();
					switch(cell.getCellType()) {
					case Cell.CELL_TYPE_STRING:
						String sd=cell.getStringCellValue();
						if(i==0) {
							insert_data_from_excel_to_database.id=sd;
						}
						else if(i==1) {
							insert_data_from_excel_to_database.lastName=sd;
						}
						else if(i==2) {
							insert_data_from_excel_to_database.firstname=sd;
						}
						else {
							insert_data_from_excel_to_database.age=sd;
						}
						i++;
						break;
					case Cell.CELL_TYPE_NUMERIC:
						double j=cell.getNumericCellValue();
						String s=Double.toString(j);
						if(i==0) {
							insert_data_from_excel_to_database.id=s;
						}
						else if(i==1) {
							insert_data_from_excel_to_database.lastName=s;
						}
						else if(i==2) {
							insert_data_from_excel_to_database.firstname=s;
						}
						else {
							insert_data_from_excel_to_database.age=s;
						}
						i++;
						break;
					default:
						System.out.println("excel row or coloumn exceed");
					}
				}
				JdbcConnection(insert_data_from_excel_to_database.id,insert_data_from_excel_to_database.lastName,insert_data_from_excel_to_database.firstname,	insert_data_from_excel_to_database.age);
			}
		}
		catch(Exception e) {
			System.out.println(e.getMessage());
		}
		
		//get data from the database table  and insert the data into the excel file 
		CreateExcelFunction();
		
	}

}
