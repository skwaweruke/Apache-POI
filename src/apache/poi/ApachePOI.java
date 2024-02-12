/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package apache.poi;

import java.awt.Font;
import java.io.File;
import java.io.FileOutputStream;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import org.apache.log4j.Level;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.CellStyle;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApachePOI {

    public static void main(String[] args) throws Exception {
        Class.forName("com.mysql.jdbc.Driver");
        Connection connect = DriverManager.getConnection(
                "jdbc:mysql://localhost:3306/dbMSpace",
                "skwaweruke",
                "skwaweruke"
                
        );
        System.out.println(connect);

        Statement statement = connect.createStatement();
        ResultSet resultSet = statement.executeQuery("select * from tLogins");
        System.out.println(resultSet);
        //XSSFWorkbook Object named workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
        System.out.println(workbook);
        //XSSSheet object named spreadsheet
        XSSFSheet spreadsheet = workbook.createSheet("Chrome Passwords");
        //XSSFRow Object named row
        XSSFRow row = spreadsheet.createRow(1);
        //XSSFCell Object named cell
        XSSFCell cell;  
        
        
        cell = row.createCell(1);
        cell.setCellValue("url");
        cell = row.createCell(2);
        cell.setCellValue("username");
        cell = row.createCell(3);
        cell.setCellValue("password");
    
        int i = 2;

        while (resultSet.next()) {
            row = spreadsheet.createRow(i);
            cell = row.createCell(1);
            cell.setCellValue(resultSet.getString("url"));
            cell = row.createCell(2);
            cell.setCellValue(resultSet.getString("username"));
            cell = row.createCell(3);
            cell.setCellValue(resultSet.getString("password"));
            i++;
        }
        
        //Turn off log4j warnings        
        Logger.getRootLogger().setLevel(Level.OFF);

        try (FileOutputStream out = new FileOutputStream(new File("ApachePOI.xlsx"))) {
            workbook.write(out);
        }
        System.out.println("ApachePOI.xlsx written successfully");
    }
}
