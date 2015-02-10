/* Title:		connection pool manager, use program for single thread.
 *
 * Project:     public
 *
 * @ Link: 		http://...
 *
 * @ Email:		ivan_ling@mezimedia.com
 *
 * @ Copyright: Copyright (c) 2008 mezimedia
 *
 * @ Author 	Ivan.ling
 *
 * @ Version 	1.0
 
 * @ last change time 2008-08-20 	
 */
package com.ask.inv.db;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.apache.log4j.Logger;

public class DBO
{
	public static enum EnumStatementType {ONLY_READ, READ_WRITE};
	
	private Logger logger = Logger.getLogger (DBO.class.getName());

	private int iTryCount = 2;
	private String strDBTag;
	
	private DBConnectionManager dbcmInstance = null;
	
	private Connection con;
	private Statement sm;
	private ResultSet rs;
	
	public static void main(String args[])
	{
		try {
			//logger.info("===========66===============" + dbLogic.getDataDailyAVG("MAS", "89", "2", 7));
			//logger.info("===" + dbLogic.isInvoiceMerchant(conn, "9981","2"));
		} catch (Exception e) {
			
			e.printStackTrace();
		}
		
		//dbLogic.release();
	}

	public DBO(String strDBTagParam)
	{
		this.strDBTag = strDBTagParam;
		
		dbcmInstance = DBConnectionManager.getInstance(DBO.class.getName());
		
		this.init();
	}

	public DBO(String strDBTagParam, boolean isInit) throws SQLException
	{
		this.strDBTag = strDBTagParam;
		
		dbcmInstance = DBConnectionManager.getInstance(DBO.class.getName());
		
		if(isInit)
			this.init();
	}

	public void release()
	{
		dbcmInstance.freeConnection(this.strDBTag, this.con);
		dbcmInstance.release(DBO.class.getName());
		
		dbcmInstance 	= null;
		this.con 		= null;
	}
	
	public void init(int iTryCount, String strNetTimeout, EnumStatementType estStatementType) throws SQLException 
	{
		switch(estStatementType){
			case ONLY_READ : 	this.sm = this.con.createStatement(ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
			case READ_WRITE : 	this.sm = this.con.createStatement();
			
			default: this.sm = this.con.createStatement();
		}
		
		this.iTryCount = iTryCount;
		
        this.sm.executeUpdate("SET SESSION net_read_timeout = " + strNetTimeout);
        this.sm.executeUpdate("SET SESSION net_write_timeout = " + strNetTimeout);
 	}

	public void init()
	{
		try {
			this.con = dbcmInstance.getConnection(this.strDBTag);
			this.sm = this.con.createStatement();
			this.sm.executeUpdate("SET SESSION net_read_timeout = 1200");
	        this.sm.executeUpdate("SET SESSION net_write_timeout = 1200");
		} catch (SQLException e) {
			logger.fatal("DBO initialize failure.");
			
			e.printStackTrace();
		} catch (NullConnectionException e) {
			logger.fatal("DBO initialize failure. message : " + e.getMessage());
			e.printStackTrace();
		}
 	}
	
	/**	execute a query
	 * 
	 * @param strSql			execute sql.
	 * @throws SQLException
	 */
	public ResultSet execQuery(String strSql) throws SQLException 
	{
       int iRetryCount = this.iTryCount;
        do {
            try {
            	this.rs = this.sm.executeQuery(strSql);
               
            	iRetryCount = 0;
            } catch (SQLException e) {
            	iRetryCount--;
                if (e.getSQLState().equalsIgnoreCase("08003") || e.getSQLState().equalsIgnoreCase("08S01") || e.getSQLState().equalsIgnoreCase("40001") ) {
                 	this.rs = this.sm.executeQuery(strSql);
                } else {
                	iRetryCount = 0;
                	  
                    logger.error("---------------> Query SQL: " + strSql);
                    logger.error("---------------> SQLState : " + e.getSQLState());
                    
                   e.printStackTrace();
                }
            }
        } while (iRetryCount > 0);
        
        return this.rs;
    }
	
	/**	execute a query
	 * 
	 * @param strSql			execute sql.
	 * @param lDelayTime		retry delay time.
	 * @throws SQLException
	 */
	public ResultSet execQuery(String strSql, long lDelayTime) throws SQLException 
	{
       int iRetryCount = this.iTryCount;
        do {
            try {
            	this.rs = this.sm.executeQuery(strSql);
               
            	iRetryCount = 0;
            } catch (SQLException e) {
            	iRetryCount--;
                if (e.getSQLState().equalsIgnoreCase("08003") || e.getSQLState().equalsIgnoreCase("08S01") || e.getSQLState().equalsIgnoreCase("40001") ) {
                 	try {
						Thread.sleep(lDelayTime);
					} catch (InterruptedException e1) {
						e1.printStackTrace();
					}
                	
                	this.rs = this.sm.executeQuery(strSql);
                } else {
                	iRetryCount = 0;
                	  
                    logger.error("---------------> Query SQL: " + strSql);
                    logger.error("---------------> SQLState : " + e.getSQLState());
                   
                     e.printStackTrace();
                }
            }
        } while (iRetryCount > 0);
        
        return this.rs;
    }
	
	/**	execute a update sql
	 * 
	 * @param strSql			execute sql.
	 * @throws SQLException
	 * @return update count.
	 */	
	public int execUpdate(String strSql) throws SQLException 
	{
          return this.sm.executeUpdate(strSql);
    }
	
	/**	execute a update sql
	 * 
	 * @param strSql			execute sql.
	 * @param lDelayTime		retry delay time.
	 * @throws SQLException
	 * @return update count.
	 */
	public int execUpdate(String strSql, long lDelayTime) throws SQLException 
	{
       int iRetryCount = this.iTryCount;
       int iReturn = 0;
        do {
            try {
            	iReturn = this.sm.executeUpdate(strSql);
               
            	iRetryCount = 0;
            } catch (SQLException e) {
            	iRetryCount--;
                if (e.getSQLState().equalsIgnoreCase("08003") || e.getSQLState().equalsIgnoreCase("08S01") || e.getSQLState().equalsIgnoreCase("40001") ) {
                 	try {
						Thread.sleep(lDelayTime);
					} catch (InterruptedException e1) {
						e1.printStackTrace();
					}
                	
					iReturn = this.sm.executeUpdate(strSql);
                } else {
                	iRetryCount = 0;
                	  
                    logger.error("---------------> Query SQL: " + strSql);
                    logger.error("---------------> SQLState : " + e.getSQLState());
                    
                    e.printStackTrace();
                }
            }
        } while (iRetryCount > 0);
        
        return iReturn;
    }
	
	public void vdStartTransaction() throws SQLException
    {
    	if(this.con !=null){
    		this.con.setAutoCommit(false);
    	}
    }
    
    public void vdEndTransaction() throws SQLException
    {
    	if(this.con !=null){
    		this.con.commit();
    	}
    }
    
    public void vdRollback() throws SQLException 
    {
    	if(this.con !=null){
	    	this.con.rollback();
    	}
    }
    
    public void close() 
    {
        try {
        	if(this.rs != null )
        		this.rs.close();
        	
        	this.rs = null;
        } catch (SQLException e) {
             logger.error("Close ResultSet exception : ");
             e.printStackTrace();
        }
        
        try {
        	if(this.sm != null )
        		this.sm.close();
        	
        	this.sm = null;
         } catch (SQLException e) {
             logger.error("Close Statemetn exception : ");
             e.printStackTrace();
        }
        
        this.release();
    }
}