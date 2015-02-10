/**
 * 
 */
package com.ask.inv.db;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.log4j.Logger;

/**
 * @author iling
 *
 */
public class DBOThreads 
{
	private static final int INT_TRY_COUNT = 2;
	
	private static DBConnectionManager DBCMInstance;
	private static Logger logger;
	
	static
	{
		logger = Logger.getLogger (DBOThreads.class.getName());
		DBOThreads.DBCMInstance = DBConnectionManager.getInstance(DBOThreads.class.getName());
	}
	
	/**
	 * 
	 */
	public DBOThreads() 
	{
	}

	/**
	 * @param args
	 */
	public static void main(String[] args) 
	{
		
	}

	public void release()
	{
		DBOThreads.DBCMInstance.release(DBOThreads.class.getName());
	}
	
	/**	execute a query
	 * 
	 * @param strSql			execute sql.
	 * @throws NullConnectionException 
	 * @throws SQLException
	 */
	public static ResultSet execQuery(String strDBTag, String strSql) throws NullConnectionException
	{
		Connection conn = DBOThreads.DBCMInstance.getConnection(strDBTag);
		
		Statement 	sm = null;
		ResultSet 	rs = null;
		
		int iRetryCount = INT_TRY_COUNT;
		
		try {
			sm = conn.createStatement();
			
	        do {
	            try {
	            	rs = sm.executeQuery(strSql);
	               
	            	iRetryCount = 0;
	            } catch (SQLException e) {
	            	iRetryCount--;
	                if (e.getSQLState().equalsIgnoreCase("08003") || e.getSQLState().equalsIgnoreCase("08S01") || e.getSQLState().equalsIgnoreCase("40001") ) {
	                 	rs = sm.executeQuery(strSql);
	                } else {
	                	iRetryCount = 0;
	                	  
	                    logger.error("---------------> Query SQL: " + strSql);
	                    logger.error("---------------> SQLState : " + e.getSQLState());
	                    
	                    e.printStackTrace();
	                }
	            }
	        } while (iRetryCount > 0);
		} catch (SQLException e1) {
			logger.error("---------------> Query SQL: " + strSql);
			e1.printStackTrace();
		}finally{
			DBOThreads.DBCMInstance.freeConnection(strDBTag, conn);
		}
		
		
        return rs;
    }
	
	/**	execute a query
	 * 
	 * @param strSql			execute sql.
	 * @throws NullConnectionException 
	 * @throws SQLException
	 */
	public static ResultSet execQueryReadOnly(String strDBTag, String strSql) throws NullConnectionException
	{
		Connection conn = DBOThreads.DBCMInstance.getConnection(strDBTag);
		
		Statement 	sm = null;
		ResultSet 	rs = null;
		
		int iRetryCount = INT_TRY_COUNT;
		
		try {
			conn.setReadOnly(true);
			sm = conn.createStatement(ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
			
			do {
	            try {
	            	rs = sm.executeQuery(strSql);
	               
	            	iRetryCount = 0;
	            } catch (SQLException e) {
	            	iRetryCount--;
	                if (e.getSQLState().equalsIgnoreCase("08003") || e.getSQLState().equalsIgnoreCase("08S01") || e.getSQLState().equalsIgnoreCase("40001") ) {
	                 	rs = sm.executeQuery(strSql);
	                } else {
	                	iRetryCount = 0;
	                	  
	                    logger.error("---------------> Query SQL: " + strSql);
	                    logger.error("---------------> SQLState : " + e.getSQLState());
	                    
	                    e.printStackTrace();
	                }
	            }
	        } while (iRetryCount > 0);
	        
	        conn.setReadOnly(false);
		} catch (SQLException e1) {
			logger.error("---------------> Query SQL: " + strSql);
			e1.printStackTrace();
		}finally{
			DBOThreads.DBCMInstance.freeConnection(strDBTag, conn);
		}
				
        return rs;
    }
	
	/**	execute a query
	 * 
	 * @param strSql			execute sql.
	 * @throws SQLException
	 */
	public static ResultSet execQuery(Connection conn, String strSql) throws SQLException 
	{
		Statement 	sm = null;
		ResultSet 	rs = null;
		
		int iRetryCount = INT_TRY_COUNT;
		
		sm = conn.createStatement();
		
        do {
            try {
            	rs = sm.executeQuery(strSql);
               
            	iRetryCount = 0;
            } catch (SQLException e) {
            	iRetryCount--;
                if (e.getSQLState().equalsIgnoreCase("08003") || e.getSQLState().equalsIgnoreCase("08S01") || e.getSQLState().equalsIgnoreCase("40001") ) {
                 	rs = sm.executeQuery(strSql);
                } else {
                	iRetryCount = 0;
                	  
                    logger.error("---------------> Query SQL: " + strSql);
                    logger.error("---------------> SQLState : " + e.getSQLState());
                    
                    e.printStackTrace();
                }
            }
        } while (iRetryCount > 0);
		
        return rs;
    }
	
	/**	execute a update sql
	 * 
	 * @param strSql			execute sql.
	 * @param lDelayTime		retry delay time.
	 * @throws SQLException
	 * @return update count.
	 */
	public static int execUpdate(Connection conn, String strSql) throws SQLException 
	{
		Statement 	sm 		= null;
		int iRetryCount 	= INT_TRY_COUNT;
		int iReturn 		= 0;
		
		sm = conn.createStatement();
		
        do {
            try {
            	iReturn = sm.executeUpdate(strSql);
               
            	iRetryCount = 0;
            } catch (SQLException e) {
            	iRetryCount--;
                if (e.getSQLState().equalsIgnoreCase("08003") || e.getSQLState().equalsIgnoreCase("08S01") || e.getSQLState().equalsIgnoreCase("40001") ) {
  					iReturn = sm.executeUpdate(strSql);
                } else {
                	iRetryCount = 0;
                	  
                    logger.error("---------------> Query SQL: " + strSql);
                    logger.error("---------------> SQLState : " + e.getSQLState());
                    
                    e.printStackTrace();
                }
            }
        } while (iRetryCount > 0);
        
        sm.close();
        
        return iReturn;
    }
	
	/**	execute a update sql
	 * 
	 * @param strSql			execute sql.
	 * @param lDelayTime		retry delay time.
	 * @throws SQLException
	 * @return update count.
	 * @throws NullConnectionException 
	 */
	public static int execUpdate(String strDBTag, String strSql) throws NullConnectionException 
	{
		Connection conn = DBOThreads.DBCMInstance.getConnection(strDBTag);
		int iReturn 		= 0;
		
		if(conn == null){
			logger.error("connection is null.");
		}else{
			Statement 	sm 		= null;
			int iRetryCount 	= INT_TRY_COUNT;
		
			try {
				sm = conn.createStatement();
			
		        do {
		            try {
		            	iReturn = sm.executeUpdate(strSql);
		               
		            	iRetryCount = 0;
		            } catch (SQLException e) {
		            	iRetryCount--;
		                if (e.getSQLState().equalsIgnoreCase("08003") || e.getSQLState().equalsIgnoreCase("08S01") || e.getSQLState().equalsIgnoreCase("40001") ) {
		  					iReturn = sm.executeUpdate(strSql);
		                } else {
		                	iRetryCount = 0;
		                	  
		                    logger.error("---------------> Query SQL: " + strSql);
		                    logger.error("---------------> SQLState : " + e.getSQLState());
		                    
		                    e.printStackTrace();
		                }
		            }
		        } while (iRetryCount > 0);
		        
		        
			} catch (SQLException e1) {
				logger.error("---------------> Query SQL: " + strSql);
				
				if(sm != null)
					try {
						sm.close();
					} catch (SQLException e) {
					}
				e1.printStackTrace();
			}finally{
				DBOThreads.DBCMInstance.freeConnection(strDBTag, conn);
			}
		}
		
        return iReturn;
    }
}
