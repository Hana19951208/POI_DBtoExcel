package com.offcn.utils;

import java.io.FileOutputStream;
import java.sql.DriverManager;
import java.sql.ResultSet;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.mysql.jdbc.Connection;
import com.mysql.jdbc.Statement;

public class DBtoExcel {
	 public final static String outputFile="e:\\country.xlsx";
	 
	    public final static String url="jdbc:mysql://localhost:3306/qingzhu";
	 
	    public final static String user="root";
	 
	    public final static String password="123";
	 
	    public static void main(String[] args) {
	        try {
	            Class.forName("com.mysql.jdbc.Driver");
	            Connection conn=(Connection) DriverManager.getConnection(url, user, password);
	            Statement stat = (Statement) conn.createStatement();
	            ResultSet resultSet = stat.executeQuery("select * from student");
	            XSSFWorkbook workbook=new XSSFWorkbook();
	            XSSFSheet sheet=workbook.createSheet("sheet1");
	            XSSFRow row = sheet.createRow((short)0);
	            XSSFCell cell=null;
	            cell=row.createCell((short)0);
	            cell.setCellValue("id");
	            cell=row.createCell((short)1);
	            cell.setCellValue("name");
	            cell=row.createCell((short)2);
	            cell.setCellValue("score");
	          
	            int i=1;
	            while(resultSet.next())
	            {
	                row=sheet.createRow(i);
	                cell=row.createCell(0);
	                cell.setCellValue(resultSet.getString("id"));
	                cell=row.createCell(1);
	                cell.setCellValue(resultSet.getString("name"));
	                cell=row.createCell(2);
	                cell.setCellValue(resultSet.getString("score"));
	                i++;
	             }
	            FileOutputStream FOut = new FileOutputStream(outputFile);
	            workbook.write(FOut);
	            FOut.flush();
	            FOut.close();
	        } catch (Exception e) {
	            e.printStackTrace();
	        }
	    }
	 

}
