// 本类只用于数据库连接及关闭操作
package dbc;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.*;
import java.util.Enumeration;
import java.util.Properties;

public class DataBaseConnection {
	static String driver = "com.mysql.jdbc.Driver";
	// URL指向要访问的数据库名scutcs
	static String url = "jdbc:mysql://127.0.0.1:3306/jkkb?autoReconnect=true&amp;autoReconnectForPools=true&useUnicode=true&characterEncoding=utf8";
	// MySQL配置时的用户名
	static String user = "root";
	// MySQL配置时的密码
	static String password = "*******";
	private static Connection conn = null;
	public static ConnectionPool connPool = null;
	private DataBaseConnection() {
	}

	public static Connection getConnection() {
		
		if(connPool == null) {
		 connPool = new ConnectionPool(driver
		 ,url
		 ,user 
		 ,password); 
		 try {
			connPool .createPool();
		} catch (Exception e) {
			e.printStackTrace();
		} 
		}
		
		Connection conn = null;
		try {
			conn = connPool .getConnection();
		} catch (SQLException e) {
			e.printStackTrace();
		} 
		return conn;
	}
	
	 public static void returnConnection(Connection conn) {
		 connPool.returnConnection(conn);
	 }
};
