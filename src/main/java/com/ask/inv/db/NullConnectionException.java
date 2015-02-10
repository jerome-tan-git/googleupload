/**
 * 
 */
package com.ask.inv.db;

/**
 * @author ivan
 *
 */
public class NullConnectionException extends Exception
{
	private static final long serialVersionUID = -5173752001457575671L;
	
	private DBConnectionManager dbcmInstance; 	//-- only instance
	private String connPoolName;
	
	/**
	 * 
	 */
	public NullConnectionException() 
	{
		
	}

	/**
	 * 
	 */
	public NullConnectionException(String strMessage) 
	{
		super(strMessage);
	}
	
	/**
	 * 
	 */
	public NullConnectionException(String strMessage, DBConnectionManager dbcm, String connPoolName) 
	{
		super(strMessage);
		
		this.dbcmInstance 	= dbcm;
		this.connPoolName 	= connPoolName;
	}
	
	public String getConnPoolName()
	{
		return connPoolName; 
	}
	
	public DBConnectionManager getDBConnManager()
	{
		return dbcmInstance;
	}
}
