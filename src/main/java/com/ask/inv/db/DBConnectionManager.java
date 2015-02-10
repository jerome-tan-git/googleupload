/* @ Title:		connection pool manager
 *
 * @ Project:   public
 *
 * @ Link: 		http://...
 *
 * @ Email:		ivan_ling@hotmail.com
 *
 * @ Copyright 	Copyright (c) 2008 mezimedia
 *
 * @ Author 	Ivan.ling
 *
 * @ Version 	4.0
 *
 * @ JDK 		1.6 or later
 *
 * @ last change time 2008-08-20 	
 */
package com.ask.inv.db;

import java.io.ByteArrayOutputStream;
import java.io.DataOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.SQLException;
import java.rmi.server.UID;
import java.util.Enumeration;
import java.util.HashSet;
import java.util.Hashtable;
import java.util.Timer;
import java.util.TimerTask;
import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.TimeUnit;
import javax.sql.XAConnection;
import javax.transaction.xa.Xid;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import org.apache.log4j.Logger;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import com.mysql.jdbc.jdbc2.optional.MysqlXADataSource;
import com.mysql.jdbc.jdbc2.optional.MysqlXid;

public class DBConnectionManager
{
	protected static DBConnectionManager dbcmInstance; 	//-- only instance
	protected static boolean isUseTimeControl;			//-- to use in web, please set true

	protected static String STRConfigFileDir;
	
	private static Logger logger;

	private static Hashtable<String, Integer> htClientConnection;
	
	private static Hashtable<String, DBXAConnectionPool> hbXAConnPooled;
	
	static
	{
		logger = Logger.getLogger (DBConnectionManager.class.getName());
		
		htClientConnection 	= new Hashtable<String, Integer>();
		hbXAConnPooled 		= new Hashtable<String, DBXAConnectionPool>();
	}
	
	public static void main(String args[])
	{
		//DBConnectionManager dbc = getInstance(DBConnectionManager.class.getName());
		//boolean isInit = dbc.init();
		//logger.debug("init result : " + isInit);
	}
	
	/* return only instance
	 * @param clientClassName			# client class name.
	 */
	public static synchronized DBConnectionManager getInstance(String clientClassName) 
	{
    	if (dbcmInstance == null) {
			dbcmInstance = new DBConnectionManager(true);
		}
		
    	if (dbcmInstance != null) {
	    	if(DBConnectionManager.htClientConnection.containsKey(clientClassName)){
	    		DBConnectionManager.htClientConnection.put(clientClassName, DBConnectionManager.htClientConnection.get(clientClassName) + 1);
	    	}else{
	    		DBConnectionManager.htClientConnection.put(clientClassName, 1);
	    	}
    	}
    	
		return dbcmInstance;
	}

	/* return only instance
	 * @param isUseTC 					# is use time control ?
	 * @param clientClassName			# client class name.
	 */
	public static synchronized DBConnectionManager getInstance(boolean isUseTC, String clientClassName) 
	{
    	if (dbcmInstance == null) {
			dbcmInstance = new DBConnectionManager(isUseTC);
		}
		
    	if (dbcmInstance != null) {
	       	if(DBConnectionManager.htClientConnection.containsKey(clientClassName)){
	    		DBConnectionManager.htClientConnection.put(clientClassName, DBConnectionManager.htClientConnection.get(clientClassName) + 1);
	    	}else{
	    		DBConnectionManager.htClientConnection.put(clientClassName, 1);
	    	}
    	}
		return dbcmInstance;
	}
	
	/*-- protected method 
	 * @param boolean isUseTC 			# is use time control ?
	 * @param log4jXMLconfigureFile		# log4j XML configure file name.
	 */	
	protected DBConnectionManager(boolean isUseTC) 
	{
		if(init()){
			isUseTimeControl = isUseTC;

			if(isUseTC)
				timerControl(21600000); //-- timer，unit：millisecond(default: 6 hour)

			if(dbcmInstance == null)
				dbcmInstance = this;
		}else{
			logger.fatal("init db config error. not create XA connection pool.");
		}
	}

	/* 	timer，some time update
	 * @param 	spaceTime 	#spaceTime
	 * @ return void
	 */
	private void timerControl(int spaceTime)
	{
		Timer timer = new Timer(true);
		timer.schedule(new TimerTask(){
		   	public void run(){
		   		vdRefreshAllConnection();
		   	}
		},spaceTime,spaceTime);
	}

	/* return connection
	 * @param 	poolName	# XA connection pool name
	 * @param 	xaConn 		# XA connection
	 * @ return if success, other false
	 */
	public boolean freeXAConnection(String poolName, XAConnection xaConn)
	{
		boolean isReturn = false;
		
		DBXAConnectionPool XAPool = hbXAConnPooled.get(poolName);
		
		if (xaConn == null) {
			logger.warn("free null XA connection.");
		}else if (XAPool == null) {
			logger.warn("isn't existed the connecton pool for " + poolName);
		}else if (!XAPool.hsCheckOutXAConnection.contains(xaConn)) {
			logger.warn("can't find this connection in XA connection pool for " + poolName);
		}else{
			isReturn = XAPool.freeXAConnection(xaConn);
		}
		
		return isReturn;
	}

	/* return connection
	 * @param 	poolName 	# connection pool name
	 * @param 	conn 		# connection
	 * @return true if success, other false
	 */
	public boolean freeConnection(String poolName, Connection conn)
	{
		boolean isReturn = false;
		
		DBXAConnectionPool XAPool = hbXAConnPooled.get(poolName);
		if (conn == null) {
			logger.warn("free null connection.");
		}else if (XAPool == null) {
			logger.warn("isn't existed the connecton pool for " + poolName);
		}else if (!XAPool.hsCheckOutConnection.contains(conn)) {
			logger.warn("can't find this connection in connection pool for " + poolName);
		}else{
			isReturn = XAPool.freeConnection(conn);
		}
		
		return isReturn;
	}

	/* fetch usable connection
	 * @param poolName 	# connect pool name
	 * @return Connection
	 */
	public XAConnection getXAConnection(String poolName) throws NullConnectionException {
		DBXAConnectionPool XAPool = hbXAConnPooled.get(poolName);
		if (XAPool == null) {
			throw new NullConnectionException("not existed this connection pool for " + poolName);
		}else{
			return XAPool.getXAConnection();
		}
	}

	/* fetch usable connection, not have usable connection for wait some time
	 * @param poolName 	# connection name
	 * @param timeout 	# wait time
	 * @return Connection or null(not usable connection)
	 */
	public XAConnection getXAConnection(String poolName, long timeout) throws NullConnectionException 
	{
		DBXAConnectionPool XAPool = hbXAConnPooled.get(poolName);
		if (XAPool == null) {
			throw new NullConnectionException("not existed this connection pool for " + poolName);
		}else{
			return XAPool.getXAConnection(timeout);
		}
	}
	
	/* fetch usable connection
	 * @param poolName 	# connect pool name
	 * @return Connection
	 */
	public Connection getConnection(String poolName) throws NullConnectionException 
	{
		DBXAConnectionPool XAPool = hbXAConnPooled.get(poolName);
		if (XAPool == null) {
			throw new NullConnectionException("not existed this connection pool for " + poolName);
		}else{
			return XAPool.getConnection();
		}
	}

	/* fetch usable connection, not have usable connection for wait some time
	 * @param poolName 		# connection name
	 * @param timeout		# wait time
	 * @ return Connection or null(not usable connection)
	 */
	public Connection getConnection(String poolName, long timeout) throws NullConnectionException
	{
		DBXAConnectionPool XAPool = hbXAConnPooled.get(poolName);
		if (XAPool == null) {
			throw new NullConnectionException("not existed this connection pool for " + poolName);
		}else{
			return XAPool.getConnection(timeout);
		}
	}
	
	/* close all connection, cancel all driver register
	 * @ return void
	*/
	public synchronized void vdRefreshAllConnection()
	{
		/*-- refresh all XA connection --*/
		Enumeration<DBXAConnectionPool> allPools = hbXAConnPooled.elements();
		while (allPools.hasMoreElements()) {
			DBXAConnectionPool XAPool = (DBXAConnectionPool) allPools.nextElement();
			XAPool.releaseAll();
		}
	}
	
	/* close all connection, cancel all driver register
	 * @param 	String	strClientClassName 	# client class name
	 * @ return void
	*/
	public synchronized void release(String strClientClassName)
	{
		/*-- wait,till last client end --*/
		if(DBConnectionManager.htClientConnection.containsKey(strClientClassName)){
			int iClientCount =  DBConnectionManager.htClientConnection.get(strClientClassName);
			if(iClientCount > 1){
				iClientCount--;
				DBConnectionManager.htClientConnection.put(strClientClassName, iClientCount);
				logger.debug("Client release success for connection by " + strClientClassName);
			}else{
				DBConnectionManager.htClientConnection.remove(strClientClassName);
				logger.debug("Client release success for connection by " + strClientClassName);
			}
		}else{
			logger.warn("client release client failure for " + strClientClassName);
			logger.warn("Client connection info to " + DBConnectionManager.htClientConnection.toString());
		}
		
		if(DBConnectionManager.htClientConnection.isEmpty()){
			Enumeration<DBXAConnectionPool> allPools = hbXAConnPooled.elements();
			while (allPools.hasMoreElements()) {
				DBXAConnectionPool XAPool = (DBXAConnectionPool) allPools.nextElement();
				XAPool.releaseAll();
			}
		}else{
			logger.debug("Client connection info to " + DBConnectionManager.htClientConnection.toString());
		}
	}
	
	/* get max connection size
	 * @ return void
	 */
	public int strGetMaxConnSize(String strPoolName) throws NullConnectionException
	{
		DBXAConnectionPool XAPool = hbXAConnPooled.get(strPoolName);
		if (XAPool == null) {
			throw new NullConnectionException("not existed this pool for " + strPoolName);
		}else{
			return XAPool.maxConn;
		}
	}
	
	/* fetch property, initialize connection and register drivers.
	 * @ return void
	 */
	public void vdOutDebugInfo()
	{
		logger.info("Client Connection = " + htClientConnection);
		logger.info("XA Connection Pooled = " + hbXAConnPooled);
		
		Enumeration<DBXAConnectionPool> allPools = hbXAConnPooled.elements();
		while (allPools.hasMoreElements()) {
			DBXAConnectionPool XAPool = (DBXAConnectionPool) allPools.nextElement();
			
			logger.debug("pool name 					= " + XAPool.name);
			logger.debug("connection url 				= " + XAPool.connUrl);
			logger.debug("connection user				= " + XAPool.user);
			logger.debug("encoding 						= " + XAPool.xaDS.getEncoding());
			logger.debug("Used Connection count 		= " + XAPool.iCheckedOut);
			logger.debug("Used XA Connection count 		= " + XAPool.iXACheckedOut);
			logger.debug("Max Connection  count 		= " + XAPool.maxConn);
			
			logger.debug("Connection Pooled size   		= " + XAPool.bkqConnctionPooled.size());
			logger.debug("XAConnection Pooled size 		= " + XAPool.bkqXAConnctionPooled.size());
		}
	}
	
	/* initialize pool to use custom configure file name.
	 * @param 	strConfigFilePathName 	# log4j XML configure file name.
	 * @ return void
	 */
	private boolean init() 
	{
    	boolean isReturn = true;
		try {
    	    String strPoolName, strConnUrl, strUser, strPassword, strMaxConn, strEncoding;
    	    
    	    Document 	docDoc;
        	DocumentBuilderFactory dbfDocFactory = DocumentBuilderFactory.newInstance(); 
    		DocumentBuilder dbDocBuilder = dbfDocFactory.newDocumentBuilder();
    		InputStream inputStream = DBConnectionManager.class.getClassLoader().getResourceAsStream("dbconfig.xml");
    		docDoc = dbDocBuilder.parse(inputStream);
    		logger.debug("Create a XA connection pool instance by config file: dbconfig.xml");
    		NodeList nlPoolNode = docDoc.getElementsByTagName("DBConfigure");
    		NodeList nlDataNode;
    		int iDataNodeCount;
    		
    		if(nlPoolNode != null && nlPoolNode.getLength() > 0){
    			int iColumnNum = nlPoolNode.getLength();
    			Node nodPoolCache, nodDataCache;
    			
    			for (int i = 0; i < iColumnNum; i++){
    				strPoolName 	= null;
    				strConnUrl		= null;
    				strUser			= null;
    				strPassword 	= null;
    				strMaxConn		= null;
    				strEncoding 	= null;
    				
    				try{
	    				nodPoolCache = nlPoolNode.item(i);
	
	    				nlDataNode 		= nodPoolCache.getChildNodes();
	    				iDataNodeCount 	= nlDataNode.getLength();
	    				for (int j = 0; j < iDataNodeCount; j++){
	    					nodDataCache = nlDataNode.item(j);
	    					
	    					if(nodDataCache.getNodeName().equalsIgnoreCase("tag")){
	    						strPoolName = nodDataCache.getTextContent();
	    					}
	    					
	    					if(nodDataCache.getNodeName().equalsIgnoreCase("url")){
	    						strConnUrl = nodDataCache.getTextContent();
	    					}
	    					
	    					if(nodDataCache.getNodeName().equalsIgnoreCase("user")){
	    						strUser = nodDataCache.getTextContent();
	    					}
	    					
	    					if(nodDataCache.getNodeName().equalsIgnoreCase("password")){
	    						strPassword = nodDataCache.getTextContent();
	    					}
	    					
	    					if(nodDataCache.getNodeName().equalsIgnoreCase("maxconn")){
	    						strMaxConn = nodDataCache.getTextContent();
	    					}
	    					
	    					if(nodDataCache.getNodeName().equalsIgnoreCase("encoding")){
	    						strEncoding = nodDataCache.getTextContent();
	    						
	    						if(strEncoding.equals(""))
	    							strEncoding = null;
	    					}
	    				}
	    				
	    				if(strPoolName == null || strConnUrl == null || strUser == null || strPassword == null){
		    				logger.warn("create connection pool " + strPoolName + " fail.");
		                    
		    				logger.warn("pool name : " + strPoolName);
		    				logger.warn("con url   : " + strConnUrl);
		    				logger.warn("user      : " + strUser);
		    				logger.warn("password  : " + strPassword);
		    				
		    				break;
		    			}
		    			
		    			if(strMaxConn == null || strMaxConn.equals("") || !strMaxConn.matches("[\\d.]+")){
		    				strMaxConn = "25";	//--> set default connection count
		    			}
		    			
		                DBXAConnectionPool XAPool = new DBXAConnectionPool(strPoolName, strConnUrl, strUser, strPassword, Integer.parseInt(strMaxConn), strEncoding);
		                hbXAConnPooled.put(strPoolName, XAPool);
		                
		                logger.debug("Create a XA connection pool instance to " + strPoolName);
    				}catch(Exception e){
    		        	try{
	    					logger.fatal("init config failure. Exception : " + e.toString());
	    		        	
		    				logger.warn("pool name : " + strPoolName);
		    				logger.warn("con url   : " + strConnUrl);
		    				logger.warn("user      : " + strUser);
		    				logger.warn("password  : " + strPassword);
    		        	}catch(Exception ee){
    		        		ee.printStackTrace();
    		        	}
    		        	e.printStackTrace();
    		        	isReturn = false;
    				}
     			}
    		}
    	}catch(Exception e){
        	logger.fatal("init config failure." );
        	
        	e.printStackTrace();
        	isReturn = false;
		}
    	
    	return isReturn;
	}

	public Xid createXid() throws IOException 
	{
		ByteArrayOutputStream gtridOut = new ByteArrayOutputStream();
		DataOutputStream dataOut = new DataOutputStream(gtridOut);
		new UID().write(dataOut);
		
		final byte[] gtrid = gtridOut.toByteArray();
		
		ByteArrayOutputStream bqualOut = new ByteArrayOutputStream();
		dataOut = new DataOutputStream(bqualOut);
		
		new UID().write(dataOut);
		
		final byte[] bqual = bqualOut.toByteArray();
		
		Xid xid = new MysqlXid(gtrid, bqual, 3306);
		return xid;
	}

	public Xid createXid(Xid xidToBranch) throws IOException 
	{
		ByteArrayOutputStream bqualOut 	= new ByteArrayOutputStream();
		DataOutputStream dataOut 		= new DataOutputStream(bqualOut);
		
		new UID().write(dataOut);
		
		final byte[] bqual = bqualOut.toByteArray();
		
		Xid xid = new MysqlXid(xidToBranch.getGlobalTransactionId(), bqual, 3306);
		
		return xid;
	}
	
	/* create XAconnection pool class.
	 */
	class DBXAConnectionPool 
	{
		private final int INT_WAIT_TIMEOUT 			= 30; //--> unit: second, default timeout for fetch XAconnection
		private final int INT_RECONNECTION_LIMIT 	= 5; //--> if connection is null, reconnection count.  
		
		private int 	iXACheckedOut;
		private int 	iCheckedOut;
		private int 	maxConn;	//-- max connection count
		private String 	name;
		private String 	password;
		private String 	connUrl;
		private String 	user;
		
		private HashSet<Connection> hsCheckOutConnection;
		private HashSet<XAConnection> hsCheckOutXAConnection;
		
		private BlockingQueue<XAConnection> bkqXAConnctionPooled;
		private BlockingQueue<Connection> 	bkqConnctionPooled;
		
		private MysqlXADataSource xaDS;
		
		private Hashtable<Connection, XAConnection> htCheckOutXAConnByConn;
		
		/* structure method
		 * 
		 * @param name 	connection pool name
		 * @param connUrl 	JDBC connUrl
		 * @param user 	DB user
		 * @param password DB password
		 * @param maxConn 	max connection count
		 */
		public DBXAConnectionPool(String name, String connUrl, String user, String password, int maxConn, String strCharEncoding) 
		{
			this.name 		= name;
			this.connUrl 	= connUrl;
			this.user 		= user;
			this.password 	= password;
			this.maxConn 	= maxConn;
			
			this.xaDS = new MysqlXADataSource();
			
			this.xaDS.setUrl(this.connUrl);
			this.xaDS.setUser(this.user);
			this.xaDS.setPassword(this.password);
			this.xaDS.setRollbackOnPooledClose(true);
			this.xaDS.setZeroDateTimeBehavior("convertToNull");
			this.xaDS.setAutoReconnect(true);
			
			if(strCharEncoding != null)
				this.xaDS.setEncoding(strCharEncoding);
		
			this.bkqXAConnctionPooled 		= new ArrayBlockingQueue<XAConnection>(this.maxConn);
			this.bkqConnctionPooled 		= new ArrayBlockingQueue<Connection>(this.maxConn);
			
			this.htCheckOutXAConnByConn 	= new Hashtable<Connection, XAConnection>();
			
			this.hsCheckOutConnection 		= new HashSet<Connection>(this.maxConn + 5);
			this.hsCheckOutXAConnection 	= new HashSet<XAConnection>(this.maxConn + 5);
		}

		/* fetch property, initialize connection and register drivers.
		 * @ return void
		 */
		public void vdOutNowStatus()
		{
			logger.debug("Client Connection = " + htClientConnection);
			logger.debug("XA Connection Pooled = " + hbXAConnPooled);
			
			logger.debug("pool name 					= " + this.name);
			logger.debug("connection url 				= " + this.connUrl);
			logger.debug("connection user				= " + this.user);
			logger.debug("encoding 						= " + this.xaDS.getEncoding());
			logger.debug("Used Connection count 		= " + this.iCheckedOut);
			logger.debug("Used XA Connection count 		= " + this.iXACheckedOut);
			logger.debug("Max Connection  count 		= " + this.maxConn);
			
			logger.debug("Connection Pooled size   		= " + this.bkqConnctionPooled.size());
			logger.debug("XAConnection Pooled size 		= " + this.bkqXAConnctionPooled.size());
			
		}
		/* reclaim connection
		 * @param conXA 	# used connection
		 */
		public synchronized boolean freeXAConnection(XAConnection conXA) 
		{
			boolean isReturn = false;
			
			try {
				isReturn = this.bkqXAConnctionPooled.offer(conXA, this.INT_WAIT_TIMEOUT, TimeUnit.SECONDS);
			} catch (InterruptedException e) {
				logger.warn("interrupted while waiting free XAConnection.");
				e.printStackTrace();
			}finally{
				this.iXACheckedOut--;
					
				this.hsCheckOutXAConnection.remove(conXA);	//--> remove check out XAconnection record.
			}
			
			return isReturn;
		}

		/* reclaim connection
		 * @param con 	# used connection
		 */
		public synchronized boolean freeConnection(Connection con) 
		{
			boolean isReturn = false;
			
			try {
				if(con.isClosed()){
					logger.warn("free invalid connection. connection be closed.");
				}else if(con.isValid(5)){
					con.setAutoCommit(true);

					if(con.isReadOnly())
						con.setReadOnly(false);
					
					isReturn = this.bkqConnctionPooled.offer(con, this.INT_WAIT_TIMEOUT, TimeUnit.SECONDS);
				}else{
					logger.warn("free invalid connection. unknown reason.");
				}
			} catch (InterruptedException e) {
				logger.warn("interrupted while waiting free Connection.");
				e.printStackTrace();
			} catch (SQLException e) {
				logger.warn("Set auto commit or Read Only failure.");
				e.printStackTrace();
			}finally{
				this.iCheckedOut--;
				
				this.hsCheckOutConnection.remove(con);	//--> remove check out connection record.
				
				if(!isReturn){
					if(htCheckOutXAConnByConn.containsKey(con)){// invalid content
						XAConnection conXA = htCheckOutXAConnByConn.get(con);
						
						htCheckOutXAConnByConn.remove(con);
						
						this.iXACheckedOut--;
						
						this.hsCheckOutXAConnection.remove(conXA);	//--> remove check out XAconnection record.
					}
				}
			}
			
			return isReturn;
		}
		
		/* fetch usable XA connection.
	 	 * @ return Connection or null(not usable connection)
		 */
		public synchronized XAConnection getXAConnection() throws NullConnectionException 
		{
			XAConnection conXA = null;
			
			int iLoopCache = 0;
			
			while(conXA == null){
				if (this.bkqXAConnctionPooled.size() > 0) {
					try {
						conXA = this.bkqXAConnctionPooled.poll(this.INT_WAIT_TIMEOUT, TimeUnit.SECONDS);
					} catch (InterruptedException e) {
						logger.warn("interrupted in waiting fetch XAConnection.");
						e.printStackTrace();
					}
				}else if (this.maxConn == 0 || this.iXACheckedOut < this.maxConn) {
					conXA = newXAConnection();
				}else{
					throw new NullConnectionException("full XA connection pool to " + this.name);
				}
				
				if (conXA != null) {
					this.iXACheckedOut++;
					
					this.hsCheckOutXAConnection.add(conXA);	//--> record check out XAconnection.
				}
				
				iLoopCache++;
				if(iLoopCache > INT_RECONNECTION_LIMIT)
					break;
			}
			
			if(conXA == null)
				throw new NullConnectionException("can't fetch XA connection.");
			else
				return conXA;
		}

		/* fetch usable connection, if not usable connection, wait some time.
		 * 
		 * @param timeout 		wait time, unit:millisecond
	 	 * @ return Connection or null
		 */

		public synchronized XAConnection getXAConnection(long timeout) throws NullConnectionException 
		{
			XAConnection conXA = null;
			
			int iLoopCache = 0;
			
			while(conXA == null){
				if (this.bkqXAConnctionPooled.size() > 0) {
					try {
						conXA = this.bkqXAConnctionPooled.poll(timeout, TimeUnit.MILLISECONDS);
					} catch (InterruptedException e) {
						logger.warn("interrupted while waiting fetch XAConnection.");
						e.printStackTrace();
					}
				}else if (this.maxConn == 0 || this.iXACheckedOut < this.maxConn) {
					conXA = newXAConnection();
				}else{
					throw new NullConnectionException("full XA connection pool to " + this.name);
				}
				
				if (conXA != null) {
					this.iXACheckedOut++;
					
					this.hsCheckOutXAConnection.add(conXA);	//--> record check out XAconnection.
				}
				
				iLoopCache++;
				if(iLoopCache > INT_RECONNECTION_LIMIT)
					break;
			}

			if(conXA == null)
				throw new NullConnectionException("can't fetch XA connection.");
			else
				return conXA;
		}
		
		/* fetch usable XA connection.
	 	 * @ return Connection or null(not usable connection)
		 */
		public synchronized Connection getConnection() throws NullConnectionException 
		{
			Connection con = null;
			
			int iLoopCache = 0;
			
			while(con == null){
				if (this.bkqConnctionPooled.size() > 0) {
					try {
						con = this.bkqConnctionPooled.poll(this.INT_WAIT_TIMEOUT, TimeUnit.SECONDS);
					} catch (InterruptedException e) {
						logger.warn("interrupted while waiting fetch Connection.");
						e.printStackTrace();
					}
				}else if (this.maxConn == 0 || this.iCheckedOut < this.maxConn) {
					con = newConnection();
				}else{
					throw new NullConnectionException("full connection pool for " + this.name);
				}
				
				if (con != null) {
					try {
						if (con.isValid(5)){
							this.iCheckedOut++;
							
							this.hsCheckOutConnection.add(con);	//--> record check out connection.
						}else if (!con.isClosed()) {
							logger.warn("Connection be closed.");
							con.close();
						}else{
							con = null;
						}
					} catch (SQLException e) {
						con = null;
						
						logger.warn("Connection be closed or invalid.");
						e.printStackTrace();
					}
				}
				
				iLoopCache++;
				if(iLoopCache > INT_RECONNECTION_LIMIT)
					break;
			}

			if(con == null)
				throw new NullConnectionException("can't fetch connection.");
			else
				return con;
		}
		
		/* fetch usable connection, if not usable connection, wait some time.
		 * 
		 * @param timeout 		wait time, unit:millisecond
	 	 * @ return Connection or null
		 */
		public synchronized Connection getConnection(long timeout) throws NullConnectionException 
		{
			Connection con = null;
			
			int iLoopCache = 0;
			
			while(con == null){
				if (this.bkqConnctionPooled.size() > 0) {
					try {
						con = this.bkqConnctionPooled.poll(timeout, TimeUnit.MILLISECONDS);
					} catch (InterruptedException e) {
						logger.warn("interrupted while waiting fetch Connection.");
						e.printStackTrace();
					}
				}else if (this.maxConn == 0 || this.iCheckedOut < this.maxConn) {
					con = newConnection();
				}else{
					throw new NullConnectionException("full connection pool for " + this.name);
				}
				
				if (con != null) {
					try {
						if (con.isValid(5)){
							this.iCheckedOut++;
							
							this.hsCheckOutConnection.add(con);	//--> record check out connection.
						}else if (!con.isClosed()) {
							logger.warn("Connection be closed.");
							con.close();
						}else{
							con = null;
						}
					} catch (SQLException e) {
						con = null;
						
						logger.warn("Connection be closed or invalid.");
						e.printStackTrace();
					}
				}
				
				iLoopCache++;
				if(iLoopCache > INT_RECONNECTION_LIMIT)
					break;
			}

			if(con == null)
				throw new NullConnectionException("can't fetch connection.");
			else
				return con;
		}

		/* close all connection
		 */
		public synchronized void release()
		{
			Connection con;
			while (this.bkqConnctionPooled.size() > 0) {
				try {
					con = this.bkqConnctionPooled.poll(this.INT_WAIT_TIMEOUT, TimeUnit.SECONDS);
					
					if(htCheckOutXAConnByConn.containsKey(con)){
						freeXAConnection(htCheckOutXAConnByConn.get(con));
					}
					
					con.close();
				}catch (SQLException e) {
					logger.error("can't close connection for connection url " + this.connUrl);
					e.printStackTrace();
				} catch (InterruptedException e) {
						logger.warn("interrupted while waiting poll Connection.");
						e.printStackTrace();
				}catch(Exception e){
					logger.warn("issue unknow exception in release all connection.");
					e.printStackTrace();
				}
			}
			
			this.bkqConnctionPooled.clear();
			
			logger.info("release connection pool success for " + this.name);
		}

		/* close all connection
		 */
		public synchronized void releaseAll() 
		{
			this.release();
			this.releaseXA();
			vdOutDebugInfo();
		}
		
		/* close all connection
		 */
		public synchronized void releaseXA() 
		{
			XAConnection conXA;
			while (this.bkqXAConnctionPooled.size() > 0) {
				try {
					conXA = this.bkqXAConnctionPooled.poll(this.INT_WAIT_TIMEOUT, TimeUnit.SECONDS);
					conXA.close();
				}catch (SQLException e) {
					logger.error("can't close connection for XA connection url " + this.connUrl);
					e.printStackTrace();
				} catch (InterruptedException e) {
					logger.warn("interrupted while waiting poll XA Connection.");
					e.printStackTrace();
				}catch(Exception e){
					logger.warn("issue unknow exception in release all XA connection.");
					e.printStackTrace();
				}
			}
			
			this.bkqXAConnctionPooled.clear();
			logger.info("release XA connection pool success for " + this.name);
		}
		
	   /* create new connect
	 	* @ return Connection or null
	    */
		private XAConnection newXAConnection() 
		{
			XAConnection conXA = null;
			try {
				conXA = this.xaDS.getXAConnection();
			}catch (SQLException e) {
				logger.error("can't create XA connection for connection URL " + this.connUrl);
				e.printStackTrace();
			}catch(Exception e){
				logger.error("can't create XA connection for connection URL " + this.connUrl);
				e.printStackTrace();
			}
			
			return conXA;
		}
		
		/* create new connect
	 	* @ return Connection or null
	    */
		private Connection newConnection() 
		{
			Connection con 		= null;
			XAConnection conXA 	= null;
			try {
				conXA = this.getXAConnection();
				if(conXA != null){
					con = conXA.getConnection();
					
					htCheckOutXAConnByConn.put(con, conXA);
				}
			}catch(NullConnectionException e){
				logger.error("can't create connection for connection URL " + this.connUrl);
				e.printStackTrace();
			}catch(SQLException e){
				logger.error("can't create connection for connection URL " + this.connUrl);
				e.printStackTrace();
			}catch(Exception e){
				logger.error("can't create connection for connection URL " + this.connUrl);
				e.printStackTrace();
			}
			
			return con;
		}
	}
}
