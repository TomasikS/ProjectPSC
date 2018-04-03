/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package excel;




import java.io.FileInputStream;

import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;

import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
 


  

    //Main function is calling readExcel function to read data from excel file

    

   public static void main(String[] args) throws Exception {

    try {

        Class forName = Class.forName("com.mysql.jdbc.Driver");
        Connection con = null;
        con = DriverManager.getConnection("jdbc:mysql://localhost/excel", "root", "root");
        
        FileInputStream input = new FileInputStream("D:\\PSCobci.xlsx");
      
        Workbook workbook= new XSSFWorkbook(input);
          
  
        Sheet sheet = workbook.getSheetAt(0);
        Row row;
       for (int i = 1; i <= sheet.getLastRowNum(); i++) {
        //for (int i = 1; i <= 2; i++) {
            row = (Row) sheet.getRow(i);
            String Obec = row.getCell(0).getStringCellValue();
            String Okres  = row.getCell(1).getStringCellValue();

            String  PSC =  row.getCell(2).getStringCellValue();
 
           
System.out.println(Obec);System.out.println(Okres);System.out.println(PSC);

//String sql =  "INSERT INTO excel (Obec, Okres, PSC) VALUES(" + Obec + "," +Okres + "," + PSC + ")";
 String query = " insert into excel (Obec, Okres, PSC)"
        + " values (?, ?, ?)"; 
         //  Statement st = con.createStatement(); 
        //   st.execute(sql);
        PreparedStatement preparedStmt = con.prepareStatement(query);
      preparedStmt.setString (1, Obec);
      preparedStmt.setString (2, Okres);
      preparedStmt.setString   (3, PSC);
      
            preparedStmt.execute();
            System.out.println("Import rows " + i);
        }
        
        
        con.close();
        input.close();
        System.out.println("Success import excel to mysql table");
     } catch (IOException e) {
     }
   }
}
   
    

 