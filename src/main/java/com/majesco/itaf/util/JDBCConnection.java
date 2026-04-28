package com.majesco.itaf.util;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Properties;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.majesco.itaf.main.Config;

public class JDBCConnection {

	private static final Logger log = LogManager.getLogger(JDBCConnection.class);
	public static Boolean sqlexception = false;
	public static String description = null;

	// Generic Connection Builders
	
	public static Connection establishDBConn() throws SQLException, ClassNotFoundException {
		String dbURL = "jdbc:oracle:thin:@172.16.244.237:1521/stgngbid";
		String user = "stg_release";
		String pass = "RELEASEDUJUE";
		return establishDBConn(dbURL, user, pass);
	}

	public static Connection establishDSConnection() throws SQLException, ClassNotFoundException {
		String dbURL = getConnectionUrl();
		String user = Config.applicationdatabaseusername;
		String pass = Config.applicationdatabasepassword;
		return establishDBConn(dbURL, user, pass);
	}

	private static String getConnectionUrl() {
		if (Config.databaseType.equalsIgnoreCase("MsSQL")) {
			return "jdbc:sqlserver://" + Config.databaseHost + ":" + Config.databasePort + ";databaseName="
					+ Config.databaseName;
		} else if (Config.databaseType.equalsIgnoreCase("Oracle")) {
			return "jdbc:oracle:thin:@" + Config.databaseHost + ":" + Config.databasePort + "/" + Config.databaseSID;
		} else if (Config.databaseType.equalsIgnoreCase("Postgres")) {
			return "jdbc:postgresql://" + Config.databaseHost + ":" + Config.databasePort + "/" + Config.databaseName;
		} else {
			log.error("No database type selected in config sheet.");
			return null;
		}
	}

	public static Connection establishPASDBConn() throws SQLException, ClassNotFoundException {
		String dbURL = getConnectionUrl();
		String user = Config.applicationdatabaseusername;
		String pass = Config.applicationdatabasepassword;
		return establishDBConn(dbURL, user, pass);
	}

	public static Connection establishHTML5BillingDBConn() throws SQLException, ClassNotFoundException {
		String dbURL;
		if (Config.databaseType.equalsIgnoreCase("MsSQL")) {
			dbURL = "jdbc:sqlserver://" + Config.databaseHost + ":" + Config.databasePort + ";databaseName="
					+ Config.databaseName;
		} else if (Config.databaseType.equalsIgnoreCase("Oracle")) {
			if ("TRUE".equalsIgnoreCase(Config.securedJdbcConnection)) {
				dbURL = "jdbc:oracle:thin:@(DESCRIPTION=(ADDRESS=(PROTOCOL=tcps)(HOST=" + Config.databaseHost
						+ ")(PORT=" + Config.databasePort + "))(CONNECT_DATA=(SERVICE_NAME=" + Config.databaseSID
						+ ")))";
			} else {
				dbURL = "jdbc:oracle:thin:@" + Config.databaseHost + ":" + Config.databasePort + "/"
						+ Config.databaseSID;
			}
		} else {
			log.error("No database type selected in config sheet.");
			return null;
		}

		String user = Config.applicationdatabaseusername;
		String pass = Config.applicationdatabasepassword;
		return establishDBConn(dbURL, user, pass);
	}

	public static Connection establishHTML5BillingCoreDBConn() throws SQLException, ClassNotFoundException {
		String dbURL;
		if (Config.databaseType.equalsIgnoreCase("MsSQL")) {
			if (Config.jbeamdatabaseName == null) {
				dbURL = "jdbc:sqlserver://" + Config.jbeamHost + ":" + Config.databasePort + ";databaseName="
						+ Config.jbeamdatabaseusername;
			} else {
				dbURL = "jdbc:sqlserver://" + Config.jbeamHost + ":" + Config.jbeamPort + ";databaseName="
						+ Config.jbeamdatabaseName;
			}
		} else if (Config.databaseType.equalsIgnoreCase("Oracle")) {
			if ("TRUE".equalsIgnoreCase(Config.securedJdbcConnection)) {
				dbURL = "jdbc:oracle:thin:@(DESCRIPTION=(ADDRESS=(PROTOCOL=tcps)(HOST=" + Config.jbeamHost + ")(PORT="
						+ Config.databasePort + "))(CONNECT_DATA=(SERVICE_NAME=" + Config.jbeamSID + ")))";
			} else {
				dbURL = "jdbc:oracle:thin:@" + Config.jbeamHost + ":" + Config.databasePort + "/" + Config.jbeamSID;
			}
		} else {
			log.error("No database type selected in config sheet.");
			return null;
		}

		String user = Config.jbeamdatabaseusername;
		String pass = Config.jbeamdatabasepassword;
		return establishDBConn(dbURL, user, pass);
	}

	// Core Connection Establishment

	public static Connection establishDBConn(String dbURL, String user, String pass)
			throws SQLException, ClassNotFoundException {

		if (dbURL == null || user == null || pass == null) {
			throw new SQLException("Invalid DB connection parameters. Please verify configuration.");
		}

		try {
			Connection conn = null;
			if (Config.databaseType.equalsIgnoreCase("MsSQL")) {
				Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
			} else if (Config.databaseType.equalsIgnoreCase("Oracle")) {
				Class.forName("oracle.jdbc.driver.OracleDriver");
			} else if (Config.databaseType.equalsIgnoreCase("Postgres")) {
				Class.forName("org.postgresql.Driver");
			} else {
				throw new SQLException("Unsupported database type: " + Config.databaseType);
			}

			conn = DriverManager.getConnection(dbURL, user, pass);
			log.info("DB Connection established successfully: " + dbURL);
			return conn;

		} catch (SQLException se) {
			log.error("SQL Exception while connecting: {}", se.getMessage(), se);
			throw new SQLException("Failed to connect establishDBConn -> " + se.getLocalizedMessage(), se);
		} catch (Exception e) {
			log.error("General Exception while connecting: {}", e.getMessage(), e);
			throw new SQLException("Failed while connecting establishDBConn -> " + e.getLocalizedMessage(), e);
		}
	}

	// Connection Close

	public static void closeConnection(Connection con) {
		if (con != null) {
			try {
				con.close();
				log.info("DB Connection closed successfully.");
			} catch (Exception e) {
				log.warn("Exception while closing JDBC Connection: {}", e.getMessage());
			}
		}
	}

	// Utility Method (Legacy Support)

	public static ResultSet establishDBConn_LnA(String filePath, String query)
			throws SQLException, IOException, ClassNotFoundException {

		Properties p = new Properties();
		try (InputStream input = new FileInputStream("./Database.properties")) {
			p.load(input);
		}

		String host = p.getProperty("Server");
		String dbName = p.getProperty("DatabaseName");
		String user = p.getProperty("UserId");
		String pass = p.getProperty("Password");
		String port = p.getProperty("Port");

		String dbURL = "jdbc:sqlserver://" + host + ":" + port + ";DatabaseName=" + dbName;
		Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");

		log.info("Connecting to host: {} | Port: {} | DB: {}", host, port, dbName);

		// Use try-with-resources for safety
		try (Connection conn = DriverManager.getConnection(dbURL, user, pass);
				Statement stmt = conn.createStatement();
				ResultSet rs = stmt.executeQuery(query)) {

			log.info("Query executed successfully: {}", query);
			// Instead of returning ResultSet (which will close), copy to a temporary
			// list/map if needed.
			// For now, just log count.
			int count = 0;
			while (rs.next())
				count++;
			log.info("Fetched {} rows from DB.", count);
			return null; // returning null since ResultSet will auto-close
		}
	}
}
