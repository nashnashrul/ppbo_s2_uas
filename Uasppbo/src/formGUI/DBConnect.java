package formGUI;

import java.io.*;
import java.sql.*;

public class DBConnect {
    private static Connection con = null;
    private static String jdbcDriver = "com.mysql.jdbc.Driver";
    private static String url = "jdbc:mysql://localhost:3306/NamaDataBase";
    private static String user = "UserDataBase";
    private static String pswd = "PasswordDataBase";
    
    public static Connection GetConnection()
    {
        if (con == null)
        {
            try
            {
                Class.forName(jdbcDriver);
                System.out.println("Trying connection ...");
                con = DriverManager.getConnection(url, user, pswd);
            }
            
            catch(Exception e)
            {
                e.printStackTrace();
            }
        }
        
        if (con != null)
            System.out.println("Connection Success!");
        
        return con;
    }
    
    public static void CloseConnection()
    {
        try
        {
            System.out.println("Clossing connection ...");
            con.close();
        }
        
        catch(Exception e)
        {
            e.printStackTrace();
        }
    }
}









