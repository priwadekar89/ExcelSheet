package com.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Iterator;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;


import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MainController extends HttpServlet {
	private static final long serialVersionUID = 1L;
  
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		
        try {
			Class.forName ("oracle.jdbc.OracleDriver");
		} catch (ClassNotFoundException e4) {
			// TODO Auto-generated catch block
			e4.printStackTrace();
		} 
        Connection conn = null;
		try {
			conn = DriverManager.getConnection("jdbc:oracle:thin:@localhost:1521/xe", "system", "Pr891996");
		} catch (SQLException e3) {
			// TODO Auto-generated catch block
			e3.printStackTrace();
		}
		PreparedStatement sql_statement=null;
        String jdbc_insert_sql = "INSERT INTO GR7_QUESTIONS"+ " VALUES" + "(?,?,?,?,?,?,?,?,?,?,?,?,?)";
        try {
			sql_statement = conn.prepareStatement(jdbc_insert_sql);
		} catch (SQLException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}
		 
        
        // We should now load excel objects and loop through the worksheet data 
		FileInputStream input_document = new FileInputStream(new File("C:\\Users\\priyanka\\Desktop\\demo.xlsx"));
        // Load workbook 
		XSSFWorkbook my_xls_workbook = new XSSFWorkbook(input_document);
        /* Load worksheet */
		XSSFSheet my_worksheet = my_xls_workbook.getSheetAt(0);
        // we loop through and insert data
        Iterator<Row> rowIterator = my_worksheet.rowIterator();
        int count=0;
        while(rowIterator.hasNext()) {
                Row row = rowIterator.next(); 
                Iterator<Cell> cellIterator = row.cellIterator();
                
                        while(cellIterator.hasNext()) {
                        	
                                Cell cell = cellIterator.next();
                                String s=String.valueOf(cell.getCellType());
                               /* for(int i=cell.getColumnIndex();i<=count;i++) {
                                	if(s.equals("NUMERIC") && i==0) {
                                		try {
											sql_statement.setDouble(1, cell.getNumericCellValue());
										} catch (SQLException e) {
											// TODO Auto-generated catch block
											e.printStackTrace();
										}
                                	}
                                	else if(s.equals("NUMERIC") && i==2) {
                                		try {
											sql_statement.setDouble(3, cell.getNumericCellValue());
										} catch (SQLException e) {
											// TODO Auto-generated catch block
											e.printStackTrace();
										}
                                	}
                                	else {
                                		try {
											sql_statement.setString(cell.getColumnIndex()+1, cell.getStringCellValue());
										} catch (SQLException e) {
											// TODO Auto-generated catch block
											e.printStackTrace();
										}
                                	}
                                	count++;
                            	}
                                */
                                if(s.equals("NUMERIC") && (cell.getColumnIndex()==0)) { 
                                	   //System.out.println("lol");
                                	try {
                                		sql_statement.setDouble(1, cell.getNumericCellValue());
										
										
									} catch (SQLException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}                                                                                     
                                }
                                else if("STRING".equals(s) && (cell.getColumnIndex()==1)){
                                	try {
                                		sql_statement.setString(2, cell.getStringCellValue());
									} catch (SQLException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
                                }
                                else if("NUMERIC".equals(s) && (cell.getColumnIndex()==2)){
                                	try {
                                		sql_statement.setDouble(3, cell.getNumericCellValue());
									} catch (SQLException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
                                }
                                else if("STRING".equals(s) && (cell.getColumnIndex()==3)){
                                	try {
                                		sql_statement.setString(4, cell.getStringCellValue());
									} catch (SQLException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
                                }
                                else if("STRING".equals(s) && (cell.getColumnIndex()==4)){
                                	try {
                                		sql_statement.setString(5, cell.getStringCellValue());
									} catch (SQLException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
                                }
                                else if("STRING".equals(s) && (cell.getColumnIndex()==5)){
                                	try {
                                		sql_statement.setString(6, cell.getStringCellValue());
									} catch (SQLException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
                                }
                                else if("STRING".equals(s) && (cell.getColumnIndex()==6)){
                                	try {
                                		sql_statement.setString(7, cell.getStringCellValue());
									} catch (SQLException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
                                }
                                else if("STRING".equals(s) && (cell.getColumnIndex()==7)){
                                	try {
                                		sql_statement.setString(8, cell.getStringCellValue());
									} catch (SQLException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
                                }
                                else if("STRING".equals(s) && (cell.getColumnIndex()==8)){
                                	try {
                                		sql_statement.setString(9, cell.getStringCellValue());
									} catch (SQLException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
                                }
                                else if("STRING".equals(s) && (cell.getColumnIndex()==9)){
                                	try {
                                		sql_statement.setString(10, cell.getStringCellValue());
									} catch (SQLException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
                                }
                                else if("STRING".equals(s) && (cell.getColumnIndex()==10)){
                                	try {
                                		sql_statement.setString(11, cell.getStringCellValue());
									} catch (SQLException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
                                }
                                else if("STRING".equals(s) && (cell.getColumnIndex()==11)){
                                	try {
                                		sql_statement.setString(12, cell.getStringCellValue());
									} catch (SQLException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
                                }
                                else if("STRING".equals(s) && (cell.getColumnIndex()==12)){
                                	try {
                                		sql_statement.setString(13, cell.getStringCellValue());
									} catch (SQLException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
                                }
                                
                        }
                        try {
							sql_statement.executeUpdate();
						} catch (SQLException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
                }
        /* Close input stream */
        input_document.close();
        
        try {
        	/* Close prepared statement */
			sql_statement.close();
			/* COMMIT transaction */
	        conn.commit();
	        /* Close connection */
	        conn.close();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        
        
	}

	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		
	}

}
