package TomasikServer.PSC.db;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.List;

public class Database {
	 private Connection conn;
	    private String driver = "com.mysql.jdbc.Driver";
	    private String url="jdbc:mysql://localhost:3306/excel";
	    private String username="root";
	    private String password="root";
	    public List<String> getPSC(String city){
	        List<String> zipCodes=new ArrayList<>();
	        try {
	            Class.forName(driver).newInstance();
	            conn = DriverManager.getConnection(url, username, password);
	            String s="SELECT PSC from excel where Obec like  '" + city + "'";

	            PreparedStatement preparedStatement=conn.prepareStatement(s);
	            ResultSet resultSet=preparedStatement.executeQuery();

	            while(resultSet.next()){
	                String p=resultSet.getString("PSC");
	                zipCodes.add(p);
	            }
	            conn.close();
	        }
	        catch (Exception e){
	            System.out.println("Err " +e.getMessage());
	        }
	        return zipCodes;
	    }

	    public List<String> getCity(String zip){
	        List<String> cities=new ArrayList<>();
	        try {
	            Class.forName(driver).newInstance();
	            conn = DriverManager.getConnection(url, username, password);
	            String s="SELECT Obec from excel where PSC  like  '" + zip + "'";

	            PreparedStatement preparedStatement=conn.prepareStatement(s);
	            ResultSet resultSet=preparedStatement.executeQuery();

	            while(resultSet.next()){
	                String c=resultSet.getString("Obec");
	                cities.add(c);
	            }
	            conn.close();
	        }
	        catch (Exception e){
	            System.out.println("Err " +e.getMessage());
	        }
	        return cities;
	    }
}


