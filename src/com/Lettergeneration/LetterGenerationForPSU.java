package com.Lettergeneration;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Formatter;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Properties;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.DatabaseConnectivity.DatabaseConnectivity;


public class LetterGenerationForPSU {

	static private Properties properties_NcbMailConfirmation = new Properties();
	
	public LetterGenerationForPSU() 
	{
		
		Connection conn = null;
		Formatter monthNameFormatter = null;
		try 
		{
			try 
			{
				InputStream ipStream = LetterGenerationForPSU.class.getResourceAsStream("PropertiesSource.properties");
				properties_NcbMailConfirmation.load(ipStream);
			}
			catch(Exception ex) 
			{
				System.out.println("Properties file is not found.");
			
			}
			
			/* For Letter Generation,
			 * The Condition will be Stage = Stage1 or Stage3 and Status = Successfully Completed. Than it will execute 
			*/
			int req_Id = 0;
			int src_Id = 0;
			conn = DatabaseConnectivity.getDatabaseConnection();
			
			String str_SelectRequestQueryForLetterGeneration = 
					"select ncbsourceentry.* from ncbsourceentry inner Join ncbrequestentry On ncbrequestentry.Id = ncbsourceentry.Req_Id Where \r\n" + 
					"ncbsourceentry.Stage=? AND ncbsourceentry.Status=? AND ncbrequestentry.Status=?";
			PreparedStatement preparedStatement_SelectRequestQueryForLetterGeneration =  conn.prepareStatement(str_SelectRequestQueryForLetterGeneration);
			preparedStatement_SelectRequestQueryForLetterGeneration.setString(1, "Stage 1");
			preparedStatement_SelectRequestQueryForLetterGeneration.setString(2, "Successfully Completed");
			preparedStatement_SelectRequestQueryForLetterGeneration.setString(3, "unprocess");

			ResultSet rs_SelectRequestQueryForLetterGeneration = preparedStatement_SelectRequestQueryForLetterGeneration.executeQuery();
			while(rs_SelectRequestQueryForLetterGeneration.next()) 
			{
				req_Id = rs_SelectRequestQueryForLetterGeneration.getInt("Req_Id");
				src_Id = rs_SelectRequestQueryForLetterGeneration.getInt("Id");
			
				
				if(req_Id != 0 && src_Id != 0) 
				{
					String str_SampleDocxFilepath = properties_NcbMailConfirmation.getProperty("SampleDocumentPath");
					String str_ReplaceDocxFilepath = properties_NcbMailConfirmation.getProperty("ReplaceDocumentOutputpath");
					String str_LetterOutputpathForStage1 = properties_NcbMailConfirmation.getProperty("LetterOutputpathForStage1");
					
					monthNameFormatter = new Formatter();
					Calendar cal = Calendar.getInstance();
					monthNameFormatter.format("%tB", cal);
					int year = cal.get(Calendar.YEAR);
					SimpleDateFormat sdFormat = new SimpleDateFormat("dd-MM-yyyy");
					Date now = new Date();
					String day = sdFormat.format(now);
					System.out.println(day);
			
					String str_PDFFilePath =  str_LetterOutputpathForStage1+"\\" + year + "\\" + monthNameFormatter + "\\" + day + "\\";
					File dir1 = new File(str_PDFFilePath);
					if (!dir1.exists())
						dir1.mkdirs();
					
					callingLetterGenerationForPSU(req_Id,src_Id,str_SampleDocxFilepath,str_ReplaceDocxFilepath,str_PDFFilePath);
					
					/* After generating the all the letter for PSU at first STAGE1 OR Third STAGE3
					 * Update query should trigger saying Stage=Stage2 OR Stage3 and Status = Successfully Completed
					*/
					java.sql.Timestamp setDateTime = new java.sql.Timestamp(new java.util.Date().getTime());
	
					String str_UpdateQueryForLetteGeneration = "UPDATE ncbsourceentry SET Stage=? ,EndTime=? WHERE Id=? AND Req_Id=?";
					PreparedStatement preparedStatement_UpdateQueryForLetteGeneration = conn.prepareStatement(str_UpdateQueryForLetteGeneration);
					preparedStatement_UpdateQueryForLetteGeneration.setString(1, "Stage 2");
					preparedStatement_UpdateQueryForLetteGeneration.setTimestamp(2, setDateTime);
					preparedStatement_UpdateQueryForLetteGeneration.setInt(3, src_Id);
					preparedStatement_UpdateQueryForLetteGeneration.setInt(4, req_Id);
					preparedStatement_UpdateQueryForLetteGeneration.execute();
					preparedStatement_UpdateQueryForLetteGeneration.close();
				}
			}

			preparedStatement_SelectRequestQueryForLetterGeneration.close();
			rs_SelectRequestQueryForLetterGeneration.close();
			
				java.sql.Timestamp setDateTime = new java.sql.Timestamp(new java.util.Date().getTime());

				String str_UpdateQueryForRequestEntry = 
						"update ncbrequestentry join ncbsourceentry on ncbrequestentry.Id = ncbsourceentry.Req_Id set ncbrequestentry.Status=? ,ncbrequestentry.EndTime=? \r\n" + 
						"where ncbsourceentry.Stage=? AND ncbsourceentry.Status=? AND ncbrequestentry.Status=?;";
				PreparedStatement preparedStatemen_UpdateQueryForRequestEntry = conn.prepareStatement(str_UpdateQueryForRequestEntry);
				preparedStatemen_UpdateQueryForRequestEntry.setString(1, "Successfully Completed");
				preparedStatemen_UpdateQueryForRequestEntry.setTimestamp(2, setDateTime);
				preparedStatemen_UpdateQueryForRequestEntry.setString(3, "Stage 2");
				preparedStatemen_UpdateQueryForRequestEntry.setString(4, "Successfully Completed");
				preparedStatemen_UpdateQueryForRequestEntry.setString(5, "unprocess");
				preparedStatemen_UpdateQueryForRequestEntry.execute();
				preparedStatemen_UpdateQueryForRequestEntry.close();
			
		}catch(Exception ex) {ex.printStackTrace();}
		finally {
			try {
				if(conn != null) {conn.close();}
				if(monthNameFormatter != null) {monthNameFormatter.close();}
			}
			catch(Exception ex) {ex.printStackTrace();}
		}

	}
	

	private void callingLetterGenerationForPSU(int req_Id,int src_Id,String str_SampleDocxFilePath,String str_ReplaceDocxFielpath,String str_PDFFilePath) 
	{
		Connection conn = null;
		try 
		{			
			conn = DatabaseConnectivity.getDatabaseConnection();
			
			String str_SelectQueryForPSUEntry  = "SELECT * FROM ncbdataentry WHERE Sector=? AND Req_Id=? AND Src_Id=?";
			PreparedStatement preparedStatement_SelectQueryForPSUEntry = conn.prepareStatement(str_SelectQueryForPSUEntry);
			preparedStatement_SelectQueryForPSUEntry.setString(1, "PSU");
			preparedStatement_SelectQueryForPSUEntry.setInt(2, req_Id);
			preparedStatement_SelectQueryForPSUEntry.setInt(3, src_Id);
			ResultSet rs_SelectQueryForPSUEntry = preparedStatement_SelectQueryForPSUEntry.executeQuery();
			while(rs_SelectQueryForPSUEntry.next()) 
			{
				int db_req_Id = rs_SelectQueryForPSUEntry.getInt("Req_Id");
				int db_src_Id = rs_SelectQueryForPSUEntry.getInt("Src_Id");
				int db_dt_Id = rs_SelectQueryForPSUEntry.getInt("Id");
				
				String dbStr_PolicyNumber = rs_SelectQueryForPSUEntry.getString("PolicyNumber");
				dbStr_PolicyNumber = rs_SelectQueryForPSUEntry.wasNull()?"":dbStr_PolicyNumber;
				
				String dbStr_PreviousInsurerName = rs_SelectQueryForPSUEntry.getString("PreviousInsurerName");
				dbStr_PreviousInsurerName = rs_SelectQueryForPSUEntry.wasNull()?"":dbStr_PreviousInsurerName;
				
				String dbStr_Address1 = rs_SelectQueryForPSUEntry.getString("AddressLine1");
				dbStr_Address1 = rs_SelectQueryForPSUEntry.wasNull()?"":dbStr_Address1;
				
				String dbStr_Address2 = rs_SelectQueryForPSUEntry.getString("AddressLine2");
				dbStr_Address2 = rs_SelectQueryForPSUEntry.wasNull()?"":dbStr_Address2;
				
				String dbStr_Address3 = rs_SelectQueryForPSUEntry.getString("AddressLine3");
				dbStr_Address3 = rs_SelectQueryForPSUEntry.wasNull()?"":dbStr_Address3;

				String dbStr_City = rs_SelectQueryForPSUEntry.getString("City");
				dbStr_City = rs_SelectQueryForPSUEntry.wasNull()?"":dbStr_City;

				String dbStr_State = rs_SelectQueryForPSUEntry.getString("State"); 
				dbStr_State = rs_SelectQueryForPSUEntry.wasNull()?"":dbStr_State;

				String dbStr_Pincode = rs_SelectQueryForPSUEntry.getString("Pincode");
				dbStr_Pincode = rs_SelectQueryForPSUEntry.wasNull()?"":dbStr_Pincode;

				/* If address is not available, need to refer to column AN- Branch Name.
				 *  On the basis of the branch name, refer to the below excel and 
				 *  get the correct address pertaining to PSU and respective branch
				*/
				
				String dbStr_InsuredName = rs_SelectQueryForPSUEntry.getString("InsuredName");
				dbStr_InsuredName = rs_SelectQueryForPSUEntry.wasNull()?"":dbStr_InsuredName;

				String dbStr_VehicleNo = rs_SelectQueryForPSUEntry.getString("RegistrationNumber");
				dbStr_VehicleNo = rs_SelectQueryForPSUEntry.wasNull()?"":dbStr_VehicleNo;

				String dbStr_ChassioNo = rs_SelectQueryForPSUEntry.getString("ChassioNo");
				dbStr_ChassioNo = rs_SelectQueryForPSUEntry.wasNull()?"":dbStr_ChassioNo;

				String dbStr_EngineNo = rs_SelectQueryForPSUEntry.getString("EngineNo");
				dbStr_EngineNo = rs_SelectQueryForPSUEntry.wasNull()?"":dbStr_EngineNo;

				String dbStr_YourPolicy_CovernoteNo = rs_SelectQueryForPSUEntry.getString("PreviousYearPolicyNumber");
				dbStr_YourPolicy_CovernoteNo = rs_SelectQueryForPSUEntry.wasNull()?"":dbStr_YourPolicy_CovernoteNo;

				String dbStr_PreviousYearNcb = rs_SelectQueryForPSUEntry.getString("PreviousYearNcbDsc");
				dbStr_PreviousYearNcb = rs_SelectQueryForPSUEntry.wasNull()?"":dbStr_PreviousYearNcb;
				
				Map<String,String> _map_ReplaceKeyAndValue = new LinkedHashMap<>();
				_map_ReplaceKeyAndValue.put("$date$", new SimpleDateFormat("dd-MM-yyyy").format(new Date()));
				_map_ReplaceKeyAndValue.put("$refNo$", dbStr_PolicyNumber);
				
				
				String [] list_SubAddress = {dbStr_Address1,dbStr_Address2,dbStr_Address3,dbStr_City,dbStr_Pincode,dbStr_State};
				String str_customerAddress = "";
				for(String sub_address : list_SubAddress) 
				{
					if(sub_address.trim() != "" && sub_address.trim().length() != 0 ) 
					{
						str_customerAddress = str_customerAddress + sub_address + "\n";
					}
				}
				
//				String str_customerAddress = dbStr_Address1+"\n"+dbStr_Address2+"\n"+dbStr_Address3+"\n"+dbStr_City+" "+dbStr_Pincode+"\n"+dbStr_State;
				
//				System.out.println("Customer Address:"+str_customerAddress);
				
				System.out.println(db_dt_Id + "- policy Number= "+dbStr_PolicyNumber);
				
				_map_ReplaceKeyAndValue.put("$customerAddress$", str_customerAddress);
				_map_ReplaceKeyAndValue.put("$insuredName$", dbStr_InsuredName);
				_map_ReplaceKeyAndValue.put("$vehicleNo$", dbStr_VehicleNo);
				_map_ReplaceKeyAndValue.put("$engineNo$", dbStr_EngineNo);
				_map_ReplaceKeyAndValue.put("$chassisNo$", dbStr_ChassioNo);
				_map_ReplaceKeyAndValue.put("$yourPolicy$", dbStr_YourPolicy_CovernoteNo);
				_map_ReplaceKeyAndValue.put("$ncb$", dbStr_PreviousYearNcb);
				
				
				
				
				String str_DocxFilename = dbStr_PolicyNumber + ".docx";
				String str_PDFFileName = dbStr_PolicyNumber + ".pdf";
								
				String str_Result = replaceTextAndCreateDocument(
						str_SampleDocxFilePath, str_ReplaceDocxFielpath, str_DocxFilename, 
						str_PDFFilePath, str_PDFFileName, 
						_map_ReplaceKeyAndValue);
				System.out.println("The Response of the method: "+str_Result);
				
				String str_Status = null;
				if(str_Result.equalsIgnoreCase("Operation Successfully Completed")) 
				{
					str_Status = "Successfully Completed";
				}else {
					str_Status = "Failed";
				}
				
				String str_InsertQueryForLetterDetails = "INSERT INTO ncbletterlogsdetails ("
						+ "Req_Id,	Src_Id,	Dt_Id,	InsuranceCompanyName,	ReferenceNumber,	InsuredName,	VehicleName,	ChassisNo,	EngineNo,"
						+ "PsuAddress,	YourPolicyCoverNote,	PreviousYearNCB,	LetterName,	LetterLocation,	Status,	Remark,	LetterGenerateDate)"
						+ "VALUES("
						+ "?,?,?,?,?,?,?,?,?,"
						+ "?,?,?,?,?,?,?,?"
						+ ")";
				java.sql.Timestamp setDateTime = null;
				 setDateTime = new java.sql.Timestamp(new java.util.Date().getTime());

				PreparedStatement preparedStatement_InsertQueryForLetterDetails = conn.prepareStatement(str_InsertQueryForLetterDetails);
				preparedStatement_InsertQueryForLetterDetails.setInt(1, db_req_Id);
				preparedStatement_InsertQueryForLetterDetails.setInt(2, db_src_Id);
				preparedStatement_InsertQueryForLetterDetails.setInt(3, db_dt_Id);
				preparedStatement_InsertQueryForLetterDetails.setString(4, dbStr_PreviousInsurerName);
				preparedStatement_InsertQueryForLetterDetails.setString(5, dbStr_PolicyNumber);
				preparedStatement_InsertQueryForLetterDetails.setString(6, dbStr_InsuredName);
				preparedStatement_InsertQueryForLetterDetails.setString(7, dbStr_VehicleNo);
				preparedStatement_InsertQueryForLetterDetails.setString(8, dbStr_ChassioNo);
				preparedStatement_InsertQueryForLetterDetails.setString(9, dbStr_EngineNo);
				preparedStatement_InsertQueryForLetterDetails.setString(10, str_customerAddress);
				preparedStatement_InsertQueryForLetterDetails.setString(11, dbStr_YourPolicy_CovernoteNo);
				preparedStatement_InsertQueryForLetterDetails.setString(12, dbStr_PreviousYearNcb);
				preparedStatement_InsertQueryForLetterDetails.setString(13, str_PDFFileName);
				preparedStatement_InsertQueryForLetterDetails.setString(14, str_PDFFilePath+"\\"+str_PDFFileName);
				preparedStatement_InsertQueryForLetterDetails.setString(15, str_Status);
				preparedStatement_InsertQueryForLetterDetails.setString(16, str_Result);
				preparedStatement_InsertQueryForLetterDetails.setTimestamp(17, setDateTime);
				preparedStatement_InsertQueryForLetterDetails.execute();
				
			}
			rs_SelectQueryForPSUEntry.close();
			preparedStatement_SelectQueryForPSUEntry.close();
		}
		catch(Exception ex) {ex.printStackTrace();}
		finally 
		{
			try 
			{
				if(conn != null) 
				{
					conn.close();
				}
			}
			catch(Exception ex) {ex.printStackTrace();}
		}	
		
	}
	
	public  String replaceTextAndCreateDocument(
			String str_SampleDocxFilePath,String str_ReplaceDocxFielPath,String str_DocxFilename,String str_PDFFilePath,String str_PDFFileName,Map<String, String> _map_ReplaceKeyAndValue) 
	{
		XWPFDocument doc = null;
		try
		{
			doc = new XWPFDocument(OPCPackage.open(str_SampleDocxFilePath));			
			for(Entry<String,String> entry: _map_ReplaceKeyAndValue.entrySet()) 
			{	
				String str_KeyForReplace = entry.getKey();
				String str_ValueForReplace = entry.getValue();
				
				for (XWPFParagraph p : doc.getParagraphs())
				{
			        List<XWPFRun> runs = p.getRuns();
			        if (runs != null)
			        {
				         for (XWPFRun r : runs)
				         {
					          String text = r.getText(0);
			        
			                 if (text != null && text.trim().toLowerCase().equalsIgnoreCase(str_KeyForReplace.trim().toLowerCase())) 
					          {
			                	  if (str_ValueForReplace.contains("\n")) {
						                String[] lines = str_ValueForReplace.split("\n");
						                r.setText(lines[0], 0); // set first line into XWPFRun
						                for(int i=1;i<lines.length;i++){
						                    // add break and insert new text
						                    r.addBreak();
						                    r.setText(lines[i]);
						                }
			                	  }else
			                	  {
						           text = text.replace(str_KeyForReplace, str_ValueForReplace);//your content
						           r.setText(text, 0);
			                	  }
					          }
				         }
			       	}
				}
			       for (XWPFTable tbl : doc.getTables()) 
			       {
				        for (XWPFTableRow row : tbl.getRows()) 
				        {
					         for (XWPFTableCell cell : row.getTableCells()) 
					         {
						          for (XWPFParagraph p : cell.getParagraphs()) 
						          {
							           for (XWPFRun r : p.getRuns()) 
							           {
								            String text = r.getText(0);
								            if (text != null && text.trim().toLowerCase().equalsIgnoreCase(str_KeyForReplace.trim().toLowerCase())) 
								            {
								            	 text = text.replace(str_KeyForReplace, str_ValueForReplace);//your content
										           r.setText(text, 0);		           
								            }
							           }
						          }
					         }
				        }
			       }
			}		         	
			FileOutputStream fos = new FileOutputStream(new File(str_ReplaceDocxFielPath,str_DocxFilename));
			doc.write(fos);
			
			PdfOptions options = PdfOptions.create();
            OutputStream out = new FileOutputStream(new File(str_PDFFilePath,str_PDFFileName));
            PdfConverter.getInstance().convert(doc, out, options);
            return "Operation Successfully Completed";
		}	
		catch(Exception ex) 
		{
			return ex.getLocalizedMessage();
		}
	}

	public static void main(String[] args) 
	{
		
		new LetterGenerationForPSU();
		
//		Map<String,String> _map_ReplaceKeyAndValue = new LinkedHashMap<>();
//		_map_ReplaceKeyAndValue.put("$date$", new SimpleDateFormat("dd-MM-yyyy").format(new Date()));
//		_map_ReplaceKeyAndValue.put("$refNo$", "201250030118800200000000");
//		
//		String str_customerAddress = "New India"+"\n"+"Kolkata"+"\n"+"West Bengal";
//		
//		
//		_map_ReplaceKeyAndValue.put("$customerAddress$", str_customerAddress);
//		_map_ReplaceKeyAndValue.put("$insuredName$", "NITHIN ANTONY");
//		_map_ReplaceKeyAndValue.put("$vehicleNo$", "KL07CK7383");
//		_map_ReplaceKeyAndValue.put("$engineNo$", "U3S5C1HE044430");
//		_map_ReplaceKeyAndValue.put("$chassisNo$", "ME3U3S5C1HE880502");
//		_map_ReplaceKeyAndValue.put("$yourPolicy$", "71070031170160011488");
//		_map_ReplaceKeyAndValue.put("$ncb$", "0");
//
//		
//		String str_SampleDocxFilePath,str_ReplaceDocxFielPath,str_DocxFilename,str_PDFFilePath,str_PDFFileName;
//		
//		str_SampleDocxFilePath = "D:\\A\\LGI\\More Project BRD\\Sample NCB Confirmation letter PSU Insurer format.docx";
//		str_ReplaceDocxFielPath = "D:\\A\\LGI\\More Project BRD\\";
//		str_DocxFilename = "ReplaceDocx.docx";
//		str_PDFFilePath = "D:\\A\\LGI\\More Project BRD\\";
//		str_PDFFileName = "PdfDocx.pdf";
		
//		String str_Result = replaceTextAndCreateDocument(str_SampleDocxFilePath, str_ReplaceDocxFielPath, str_DocxFilename, str_PDFFilePath, str_PDFFileName, _map_ReplaceKeyAndValue);
//		System.out.println(str_Result);
	}
	
}
