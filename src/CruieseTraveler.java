import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.Socket;
import java.net.URL;
import javax.net.ssl.SSLSocket;
import javax.net.ssl.SSLSocketFactory;

import org.apache.commons.lang3.time.DateUtils;
import javax.xml.namespace.QName;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.tempuri.CruisesTravelerWebService;
import org.tempuri.CruisesTravelerWebServiceSoap;
import org.tempuri.PSSApiResponse;
import jakarta.xml.bind.DatatypeConverter;
import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
//import com.microsoft.sqlserver.jdbc.SQLServerDataSource;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class CruieseTraveler {
	public static void SocketSend() {
		
	
		 String host = "tpnet.twport.com.tw"; // 替換為服務器地址
	        int port = 443; // 替換為正確的端口
	        Date today=new Date();
	        SimpleDateFormat sm = new SimpleDateFormat("yyyy-mm-dd");
	        // SOAP 請求
	        String soapRequest = 
	        		"<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n"+
	        "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"+
	          "<soap:Body>\r\n"+
	            "<GetExcel xmlns=\"http://tempuri.org/\">\r\n"+
	              "<sDate>"+sm.format(today)+"</sDate>\r\n"+
	              "<eDate>"+sm.format(DateUtils.addDays(today,7))+"</eDate>\r\n"+
	              "<cruMode></cruMode>\r\n"+
	              "<shipType></shipType>\r\n"+
	              "<cruCate></cruCate>\r\n"+
	              "<vesType>1</vesType>\r\n"+
	              "<firstCru></firstCru>\r\n"+
	              "<dirRoute></dirRoute>\r\n"+
	              "<vesselCname></vesselCname>\r\n"+
	              "<vesselNo></vesselNo>\r\n"+
	              "<callSign></callSign>\r\n"+
	              "<port>KEL</port>\r\n"+
	            "</GetExcel>\r\n"+
	          "</soap:Body>\r\n"+
	        "</soap:Envelope>";
	        //Socket socket=null;
	        try {
	        	SSLSocketFactory factory=(SSLSocketFactory)SSLSocketFactory.getDefault();
	    		SSLSocket socket=(SSLSocket)factory.createSocket(host, port);
	        	//SocketAddress address = new InetSocketAddress(host, port);
	        	//Socket socket = new Socket(host, port);
	        	//socket.connect(address);
	        	 //socket.connect(null);
	             OutputStream output = socket.getOutputStream();
	             BufferedReader reader = new BufferedReader(new InputStreamReader(socket.getInputStream()));

	            // 發送 HTTP 請求
	            String httpRequest = "POST /HOPWS/WS/CruisesTravelerWebService.asmx HTTP/1.1\r\n" + // 替換為實際的服務端點
	                    "Host: " + host + "\r\n" +
	                    "Content-Type: text/xml; charset=utf-8\r\n" +
	                    "Content-Length: " + soapRequest.length() + "\r\n" +
	                    "SOAPAction: \"http://tempuri.org/GetExcel\"\r\n" +
	                    "\r\n" + 
	                    soapRequest;
	            System.out.println(httpRequest);
	            output.write(httpRequest.getBytes());
	            output.flush();

	            // 讀取響應
	            String line;
	            while ((line = reader.readLine()) != null) {
	                System.out.println(line);
	            }
	            socket.close();

	        } catch (Exception e) {
	            e.printStackTrace();
	        }
	}
	public static void ReadData(List<String> row) 
	{
		// Create a variable for the connection string.
        String connectionUrl = "jdbc:sqlserver://<server>:<port>;databaseName=AdventureWorks;user=<user>;password=<password>";

        try {
        	Connection conn=null;
        	Statement stmt =null;
        	String vesselEName=row.get(9);
        	String vesselCName=row.get(10);
        	String sDate=row.get(15);
        	String sTime=row.get(17);
        	String eDate=row.get(19);
        	String eTime=row.get(21);
        	String status=row.get(row.size()-2);
        	//System.out.println(row);
        	System.out.println(String.format("英文船名:%-20s,中文船名:%-10s,到港日期:%-16s,離港日期:%-16s,狀態:%-4s",vesselEName.trim(),vesselCName.trim(),sDate+" "+sTime,eDate+" "+eTime,status));
            //conn = DriverManager.getConnection(connectionUrl);
            //stmt = conn.createStatement();
            //createTable(stmt);
			//String SQL = "SELECT * FROM Production.Product;";
           // ResultSet rs = stmt.executeQuery(SQL);
            //displayRow("PRODUCTS", rs);
        }
        // Handle any errors that may have occurred.
        //catch(SQLException e) {
        //	e.printStackTrace();
        //}
        catch ( Exception e) {
            e.printStackTrace();
        }
		
		
	}
	protected static Workbook getWorkbook(String path) throws FileNotFoundException, IOException {
		Workbook wb = null;
		if(path == null) return null;
		String extString = path.substring(path.lastIndexOf("."));
		InputStream is = new FileInputStream(path);
		if(".xls".equals(extString)){
		    wb = new HSSFWorkbook(is);
		}else if(".xlsx".equals(extString)){
		    wb = new XSSFWorkbook(is);
		}
		return wb;
	}
	protected static Sheet getSheet(Workbook workbook, int sheetNo){
		return (Sheet) workbook.getSheetAt(sheetNo);
	}
	protected static String readCell(Cell cell){
		return (String) getCellFormatValue(cell);
	}
	protected static Object getCellFormatValue(Cell cell){
		Object cellValue = null;
	        if(cell!=null){
	            switch(cell.getCellType()){
		            case NUMERIC:
		            	cellValue 	= cell.getNumericCellValue();  
		                break;
		            case FORMULA:
		                if(DateUtil.isCellDateFormatted(cell)){
		                    cellValue = cell.getDateCellValue();
		                }else{
		                    cellValue = String.valueOf(cell.getNumericCellValue());
		                }
		                break;
		            case STRING:
		                cellValue = cell.getRichStringCellValue().getString();
		                break;
		            default:
		                cellValue = "";
	            }
		    }else{
	            cellValue = "";
	        }
		return cellValue;
	}

	protected static List<List<String>> readFields(Workbook workbook, int sheetNo, int firstRow, int firstCol) throws Exception {
		org.apache.poi.ss.usermodel.Sheet sheet =workbook.getSheetAt(sheetNo);
		Row row =  sheet.getRow(0); 
		int rownum = (int )sheet.getPhysicalNumberOfRows();
		int colnum = row.getPhysicalNumberOfCells();
		List<List<String>> list = new ArrayList<>();
		List<String> _inner;
		for(int i = firstRow; i < rownum; i++){
			row = (Row)sheet.getRow(i);
			 _inner = new ArrayList<>();
			if(row != null){
				for(int j = firstCol; j < colnum; j++){
					_inner.add(readCell(row.getCell(j)));
				}
				list.add(_inner);
			}else{
				break;
			}
		}
		return list;
	}
    public static void main(String[] args) {
    	//SocketSend();
    	Date today=new Date();
        SimpleDateFormat sm = new SimpleDateFormat("yyyy-MM-dd");
        //System.out.
    	String requrl = "https://tpnet.twport.com.tw/HOPWS/WS/CruisesTravelerWebService.asmx";
    	String soapRequest = 
        		"<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n"+
        "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"+
          "<soap:Body>\r\n"+
            "<GetExcel xmlns=\"http://tempuri.org/\">\r\n"+
              //"<sDate>2024/10/07</sDate>\r\n"+
              //"<eDate>2024/10/15</eDate>\r\n"+
              "<sDate>"+sm.format(today)+"</sDate>\r\n"+
              "<eDate>"+sm.format(DateUtils.addDays(today,10))+"</eDate>\r\n"+
              "<cruMode></cruMode>\r\n"+
              "<shipType></shipType>\r\n"+
              "<cruCate></cruCate>\r\n"+
              "<vesType>1</vesType>\r\n"+
              "<firstCru></firstCru>\r\n"+
              "<dirRoute></dirRoute>\r\n"+
              "<vesselCname></vesselCname>\r\n"+
              "<vesselNo></vesselNo>\r\n"+
              "<callSign></callSign>\r\n"+
              "<port>KEL</port>\r\n"+
            "</GetExcel>\r\n"+
          "</soap:Body>\r\n"+
        "</soap:Envelope>";
    	String reqcontentType = "text/xml; charset=utf-8";
    	String reqsoapAction = "http://tempuri.org/GetExcel";
    	String resxml = SOAPClient.sendSoapPost(requrl, soapRequest, reqcontentType, reqsoapAction);
    	//System.out.println(resxml);
    	PSSApiResponse rsp=SOAPClient.getvaluefromxml(resxml);
    	if(!rsp.isErrorHappend()&&rsp.isProcessStatus())
    	{
    		//System.out.println(rsp.getMsgObj());
    		byte[] data = DatatypeConverter.parseBase64Binary(rsp.getMsgObj().toString());
    		try {
    			String file="./CruieseTravelerList.xls";
    		    FileOutputStream stream = new FileOutputStream(file) ;
    		    stream.write(data);
    		    stream.close();
    		    Thread.sleep(1000);
    		    System.out.println("file output complete...");
    		    Workbook wb=getWorkbook(file);
    		    if(wb!=null) {
    		    //Sheet sheet1= getSheet(wb,0);
    		     List<List<String>> mylist = readFields(wb,0,0,0);
    		     //System.out.println(mylist);
    		     
    		     for(int i=2;i<mylist.size();i++) {
    		    	 List<String> collist= mylist.get(i);
    		    	 //System.out.println(collist);
    		    	 if(i>=2) {
    		    		 ReadData(collist);
    		    	 }
    		     }
    		     //QName SERVICE_NAME = new QName("http://tempuri.org/", "CruisesTravelerWebService");

    		     //URL wsdlURL = CruisesTravelerWebService.WSDL_LOCATION;
    		     //CruisesTravelerWebService svc=new CruisesTravelerWebService(wsdlURL,SERVICE_NAME);
    		     //CruisesTravelerWebServiceSoap soap=svc.getCruisesTravelerWebServiceSoap();
    		     //PSSApiResponse rspResult=soap.getExcel("2024/10/11", "2024/10/18", "", "", "", "1", "", "", "", "", "", "KEL");
    		     //String colString=String.join(",",(CharSequence[]) collist.toArray());
    		     
    		    }
    		
    		}catch(FileNotFoundException e) {
    			
    		}catch(Exception e) {
    			
    		}
    	}
    }
    

    
}
