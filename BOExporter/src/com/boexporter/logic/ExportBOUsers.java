package com.boexporter.logic;

import java.util.*;
import com.crystaldecisions.sdk.plugin.desktop.usergroup.*;
import com.crystaldecisions.sdk.properties.IProperties;
import com.crystaldecisions.sdk.properties.IProperty;
import com.crystaldecisions.sdk.plugin.desktop.user.*;
import com.crystaldecisions.sdk.plugin.desktop.common.*;
import com.crystaldecisions.sdk.plugin.desktop.folder.IFolder;
import com.crystaldecisions.sdk.occa.infostore.*;
import com.crystaldecisions.sdk.framework.IEnterpriseSession;
import com.crystaldecisions.sdk.framework.CrystalEnterprise;
import com.crystaldecisions.sdk.exception.SDKException;
import java.sql.*;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import java.io.*;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;

public class ExportBOUsers {
	
	String boUser = "";
	String boPassword = ""; 
	String boCmsName = "";
	String boAuthType = "secEnterprise";
	String excelPath = "BODataExport.xls";
	
	public ExportBOUsers(String boCmsName, String boUser, String boPassword) {
		this.boUser = boUser;
		this.boPassword = boPassword;
		this.boCmsName = boCmsName;
	}
	
	public ExportBOUsers(String boCmsName, String boUser, String boPassword,String excelPath) {
		this.boUser = boUser;
		this.boPassword = boPassword;
		this.boCmsName = boCmsName;
		this.excelPath = excelPath;
	}
	
	public IInfoStore getBoInforStore() {
		String boUser = this.boUser;
		String boPassword = this.boPassword;
		String boCmsName = this.boCmsName;
		String boAuthType = this.boAuthType;
		ArrayList<BOUserInfo> userLists;
		// Declare BO Variables
		IInfoStore boInfoStore=null;
		IInfoObjects boInfoObjects=null;
		IInfoObjects boInfoObjects2=null;
		IInfoObject boInfoObject=null;
		SDKException failure = null;
		IEnterpriseSession boEnterpriseSession = null;
		
		// Logon and obtain an Enterprise Session
		try {
			boEnterpriseSession = CrystalEnterprise.getSessionMgr().logon( boUser, boPassword, boCmsName, boAuthType);
			boInfoStore = (IInfoStore) boEnterpriseSession.getService("", "InfoStore");
		} catch (SDKException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return boInfoStore;
		
	}
	
	public IInfoObjects queryBoInfoStore(String querystring) {
		IInfoStore boInfoStore=null;
		IInfoObjects queryResult = null;
		boInfoStore = this.getBoInforStore();
		try {
			queryResult =  boInfoStore.query(querystring);
		} catch (SDKException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return queryResult;
	}
	
	
	public ArrayList getBoUsers() throws Exception{
		String userListQueryString = "SELECT * FROM CI_SYSTEMOBJECTS WHERE SI_KIND = 'User' ORDER BY SI_ID ASC";
		IInfoObjects userListResult = this.queryBoInfoStore(userListQueryString);
		ArrayList<BOUserInfo> boUserList= new ArrayList();
		if(userListResult.size() > 0) {
			
			for(Iterator boUsers = userListResult.iterator(); boUsers.hasNext();) {
				BOUserInfo boUser = new BOUserInfo();
				IUser boUserAccount = (IUser) boUsers.next();
				String userName=boUserAccount.getTitle();
				String userFullName=boUserAccount.getFullName();
				String userEmailAddress=boUserAccount.getEmailAddress();
				Set userGroupsObject = (Set)boUserAccount.getGroups();
				Set userGroups = new HashSet();
				boUser.setUserAccount(userName);
				boUser.setUserFullName(userFullName);
				boUser.setUserEmail(userEmailAddress);
				
				
				for (int i = 0; i < userGroupsObject.size(); i++) {
					String groupId = ((Integer)userGroupsObject.toArray()[i]).toString();
					String groupName = this.getGroupNameById(groupId);
					userGroups.add(groupName);
				}
				boUser.setUserGroups(userGroups);
				boUserList.add(boUser);
				
			}
			
			
		}
		
		
		return boUserList;
	}
	
	
	public Map<String, String> getAllReportFolders() throws SDKException{
		IInfoObjects folderListResult=null;
		IInfoObject boInfoObject=null;
		
		String querystring= "SELECT SI_ID,SI_NAME,SI_PARENTID,SI_PATH FROM CI_INFOOBJECTS WHERE 	SI_KIND in ( 'Folder') 	";
		
		Map<String,String> map = new HashMap<String,String>();
		
		folderListResult = this.queryBoInfoStore(querystring);
		if(folderListResult.size() > 0) {
			
			for(Iterator boFolders = folderListResult.iterator(); boFolders.hasNext();) {
				IInfoObject  item =  (IInfoObject) boFolders.next();
				String pathValue = "";
				IProperties itemProperties = (IProperties) item.properties();
				String pathKey = itemProperties.get("SI_ID").toString();
				
				if(item.getKind().equals("Folder")) {
					
					 IFolder iifolder=(IFolder)item;
					 String finalFolderPath="";
					 if(iifolder.getPath()!= null)
					   {
					    String path[]=iifolder.getPath();
					    for(int fi=0;fi<path.length;fi++)
					    {
					     finalFolderPath = path[fi] + "/" + finalFolderPath;
					    }
					    finalFolderPath = finalFolderPath + iifolder.getTitle();
					   }
					   else
					   {
					    finalFolderPath=finalFolderPath+iifolder.getTitle();
					   }
					 pathValue=finalFolderPath;
				}
				
				//String folderName = itemProperties.get("SI_NAME").toString();
				//IFolder childProperty = (IFolder) itemProperties.getProperty("SI_PATH");
				
				map.put(pathKey, pathValue);
			}
			
		}
		
		return map;
	}

	public String getReportPathById(String pathKey) {
		
		
		
		return "";
		
	}
	
	
	public ArrayList getBoReports() throws Exception{
		String queryString = "SELECT SI_ID,SI_NAME,SI_KIND,SI_PARENT_FOLDER,SI_UPDATE_TS,SI_CREATION_TIME,SI_PATH FROM CI_INFOOBJECTS WHERE SI_KIND in ('Webi' , 'CrystalReport','Publication')";
		IInfoObjects listResult = this.queryBoInfoStore(queryString);
		System.out.print(listResult);
		
		ArrayList<BOReportInfo> boReportList= new ArrayList();
		if(listResult.size() > 0) {
			
			Map<String,String> pathMap = this.getAllReportFolders(); 
			for(Iterator boReports = listResult.iterator(); boReports.hasNext();) {
				
				
				IInfoObject  item =  (IInfoObject) boReports.next();
				IProperties itemProperties = (IProperties) item.properties();
				if(itemProperties != null) {
					BOReportInfo boReport = new BOReportInfo();
					String reportId = itemProperties.get("SI_ID").toString();
					String reportName = itemProperties.get("SI_NAME").toString();
					String reportType = itemProperties.get("SI_KIND").toString();
					IProperty childProperty = itemProperties.getProperty("SI_CREATION_TIME");
					String reportCreateTime = childProperty.getValue().toString();
					childProperty = itemProperties.getProperty("SI_UPDATE_TS");
					String reportUpdateTime = childProperty.getValue().toString();
					String reportParentId = itemProperties.get("SI_PARENT_FOLDER").toString();
					String reportPath = pathMap.get(reportParentId);
					//String reportPath = itemProperties.get("SI_PARENT_CUID").toString();
					
					boReport.setReportId(reportId);
					boReport.setReportName(reportName);
					boReport.setReportType(reportType);
					boReport.setReportCreateTime(reportCreateTime);
					boReport.setReportUpdateTime(reportUpdateTime);
					boReport.setReportParentId(reportParentId);
					boReport.setReportPath(reportPath);
					
//					System.out.println(boReport.getDetailString());
					
					//System.out.println(logstr);
					boReportList.add(boReport);

					
				}
				
				
			}
		}
		
		return boReportList;
		
	}
	
	
	
	public String getGroupNameById(String groupId) {
		
		IInfoObjects boInfoObjects2=null;
		IInfoObject boInfoObject=null;
		String querystring= "Select * from CI_SYSTEMOBJECTS where SI_ID = '" + groupId + "'";
		
		boInfoObjects2 = this.queryBoInfoStore(querystring);
		boInfoObject = (IInfoObject) boInfoObjects2.get(0);
		String groupName = boInfoObject.getTitle();
		//System.out.println(groupName);
		return groupName;
		
	}
	/**
	 * old function,just reference
	 * @throws Exception
	 */
	private void exportBoUsers() throws Exception {
		String boUser = this.boUser;
		String boPassword = this.boPassword;
		String boCmsName = this.boCmsName;
		String boAuthType = this.boAuthType;
		// Set some values
		int max_id = 0;
		// Declare Variables
		IInfoStore boInfoStore=null;
		IInfoObjects boInfoObjects=null;
		IInfoObjects boInfoObjects2=null;
		IInfoObject boInfoObject=null;
		SDKException failure = null;
		IEnterpriseSession boEnterpriseSession = null;
		HSSFWorkbook wb = new HSSFWorkbook();
		// modify 20181221
		HSSFSheet sheet = wb.createSheet("user_list");
		HSSFRow rowhead = sheet.createRow((short)0);
		// modify 20181221
		rowhead.createCell((short) 1).setCellValue("User");
		rowhead.createCell((short) 0).setCellValue("User_Groups");
		rowhead.createCell((short) 2).setCellValue("User_Full_Name");
		rowhead.createCell((short) 3).setCellValue("User_Email");
		int index=1;
		try{
		// Logon and obtain an Enterprise Session
		boEnterpriseSession = CrystalEnterprise.getSessionMgr().logon( boUser, boPassword, boCmsName, boAuthType);
		boInfoStore = (IInfoStore) boEnterpriseSession.getService("", "InfoStore");
		// Loop through all users
		for(;;) {
		boInfoObjects = (IInfoObjects)boInfoStore.query("SELECT TOP 1000 * FROM CI_SYSTEMOBJECTS WHERE SI_KIND = 'User' And SI_ID > " + max_id + " ORDER BY SI_ID ASC");
		if(boInfoObjects.size() == 0)
		break;
		for(Iterator boUsers = boInfoObjects.iterator(); boUsers.hasNext();) {
			IUser boUserAccount = (IUser) boUsers.next();
			System.out.println("--------------------------------------<BR>");
			System.out.println(boUserAccount.getTitle() + " (" + boUserAccount.getID() + ")<BR>");
			String userName=boUserAccount.getTitle();
			HSSFRow row11 = sheet.createRow((short)index);
			row11.createCell((short) 1).setCellValue(userName);
			String userFullName=boUserAccount.getFullName();
			row11.createCell((short) 2).setCellValue(userFullName);
			String userEmailAddress=boUserAccount.getEmailAddress();
			row11.createCell((short) 3).setCellValue(userEmailAddress);
			// modify 20181221
			//index++;
			// Retrieve all the groups for this user
			java.util.Set boGroups = (java.util.Set)boUserAccount.getGroups();
			for (int i = 0; i < boGroups.size(); i++) {
			boInfoObjects2 = boInfoStore.query("Select * from CI_SYSTEMOBJECTS where SI_ID = '" + ((Integer)boGroups.toArray()[i]).toString() + "'");
			boInfoObject = (IInfoObject)boInfoObjects2.get(0);
			System.out.println((i+1)+" :"+boInfoObject.getTitle() + "<BR>");
			HSSFRow row = sheet.createRow((short)index);
			// modify 20181221
			row.createCell((short) 0).setCellValue(boInfoObject.getTitle());
			
			row.createCell((short) 1).setCellValue(userName);
			row.createCell((short) 2).setCellValue(userFullName);
			row.createCell((short) 3).setCellValue(userEmailAddress);
			index++;
		}
		System.out.println("<BR>");
		max_id = boUserAccount.getID();
		}
		}
		}
		catch(SDKException e)
		{
		System.out.println(e.getMessage());
		}
		finally
		{
		FileOutputStream fileOut = new FileOutputStream(excelPath);
		wb.write(fileOut);
		fileOut.close();
		System.out.println(" Excel file created successfully");
		boEnterpriseSession.logoff();
		}
	}
	
	public boolean saveBoUsersAndReportListToExcel() throws Exception {
		ArrayList<BOUserInfo> userList = this.getBoUsers();
		ArrayList<BOReportInfo> reportList = this.getBoReports();
		String excelPath = this.excelPath;
		//////////////////////////��ArrayList�е�����д�뵽����excel��///////////////////////////        
		//��һ��������һ��workbook��Ӧһ��excel�ļ�
		HSSFWorkbook workbook = new HSSFWorkbook();
		
		HSSFCellStyle styleBorderThin= workbook.createCellStyle();
		 
		styleBorderThin.setBorderBottom(HSSFCellStyle.BORDER_THIN); //�±߿�
		styleBorderThin.setBorderLeft(HSSFCellStyle.BORDER_THIN);//��߿�
		styleBorderThin.setBorderTop(HSSFCellStyle.BORDER_THIN);//�ϱ߿�
		styleBorderThin.setBorderRight(HSSFCellStyle.BORDER_THIN);//�ұ߿�
		
		//  ##### process Users 
		//�ڶ�������workbook�д���һ��sheet��Ӧexcel�е�sheet
		HSSFSheet sheet = workbook.createSheet("BOUserList");
		//����������sheet������ӱ�ͷ��0�У��ϰ汾��poi��sheet������������
		HSSFRow row = sheet.createRow(0);
		//���Ĳ���������Ԫ�����ñ�ͷ
		HSSFCell cell = row.createCell(0);
		cell.setCellValue("Account");
		cell = row.createCell(1);
		cell.setCellValue("Email");
		cell=row.createCell(2);
		cell.setCellValue("FullName");
		cell=row.createCell(3);
		cell.setCellValue("Groups");
		
		for (Cell cell2 : row) {
			cell2.setCellStyle(styleBorderThin);
		}
		
		
		//���岽��д��ʵ�����ݣ�ʵ��Ӧ������Щ���ݴ����ݿ�õ�,�����װ���ݣ����ϰ����󡣶��������ֵ��Ӧ���ÿ�е�ֵ
		for (int i = 0; i <userList.size(); i++) 
		{
				HSSFRow row1 = sheet.createRow(i+1);
				BOUserInfo mo=userList.get(i);
				//������Ԫ����ֵ
				row1.createCell(0).setCellValue(mo.getUserAccount());
				row1.createCell(1).setCellValue(mo.getUserEmail());
				row1.createCell(2).setCellValue(mo.getUserFullName());
				row1.createCell(3).setCellValue(mo.getUserGroups());
				for (Cell cell2 : row1) {
					cell2.setCellStyle(styleBorderThin);
				}
		}
		
		
		
	//  ##### process report 
		//�ڶ�������workbook�д���һ��sheet��Ӧexcel�е�sheet
		HSSFSheet sheetReport = workbook.createSheet("BOReportList");
		//����������sheet������ӱ�ͷ��0�У��ϰ汾��poi��sheet������������
		HSSFRow rowReport = sheetReport.createRow(0);
		//���Ĳ���������Ԫ�����ñ�ͷ
		HSSFCell cellReport = rowReport.createCell(0);
		cellReport.setCellValue("Report ID");
		cellReport = rowReport.createCell(1);
		cellReport.setCellValue("Report Name");
		cellReport = rowReport.createCell(2);
		cellReport.setCellValue("Report Type");
		cellReport=rowReport.createCell(3);
		cellReport.setCellValue("Create Time");
		cellReport=rowReport.createCell(4);
		cellReport.setCellValue("Update Time");
		cellReport=rowReport.createCell(5);
		cellReport.setCellValue("Folder ID");
		cellReport=rowReport.createCell(6);
		cellReport.setCellValue("Folder Path");
		
		for (Cell cell2 : rowReport) {
			cell2.setCellStyle(styleBorderThin);
		}
		//���岽��д��ʵ�����ݣ�ʵ��Ӧ������Щ���ݴ����ݿ�õ�,�����װ���ݣ����ϰ����󡣶��������ֵ��Ӧ���ÿ�е�ֵ
		for (int i = 0; i <reportList.size(); i++) 
		{
				HSSFRow row1 = sheetReport.createRow(i+1);
				BOReportInfo mo=reportList.get(i);
				//������Ԫ����ֵ
				row1.createCell(0).setCellValue(mo.getReportId());
				String reporNameStr = new String(mo.getReportName().getBytes(),"gbk");
				System.out.println(reporNameStr);
				row1.createCell(1).setCellValue(reporNameStr);
				
				
				row1.createCell(2).setCellValue(mo.getReportType());
				row1.createCell(3).setCellValue(mo.getReportCreateTime());
				row1.createCell(4).setCellValue(mo.getReportUpdateTime());
				row1.createCell(5).setCellValue(mo.getReportParentId());
				row1.createCell(6).setCellValue(mo.getReportPath());
				
				for (Cell cell2 : row1) {
					cell2.setCellStyle(styleBorderThin);
				}
		}
		
		// ������������Ӧ
		for (int i = 0; i <= 3; i++)
	      {
			sheet.autoSizeColumn(i);
	      }
		
		for (int i = 0; i <= 6; i++)
	      {
			sheetReport.autoSizeColumn(i);
	      }
		//���ļ����浽ָ����λ��
		try 
		{
			FileOutputStream fos = new FileOutputStream(excelPath);
			workbook.write(fos);
			System.out.println("��ϲ����д��ɹ�������������");
			fos.close();
		} catch (IOException e) 
		{
			System.out.println("д���ļ���������");
			e.printStackTrace();
			return false;
		}

		return true;
	}
	
	public boolean saveBoReportListToExcel() {
		
		return true;
	}
	
	public boolean saveBoUsersToExcel() throws Exception {
		ArrayList<BOUserInfo> userList = this.getBoUsers();
		String excelPath = this.excelPath;
		//////////////////////////��ArrayList�е�����д�뵽����excel��///////////////////////////        
		//��һ��������һ��workbook��Ӧһ��excel�ļ�
		HSSFWorkbook workbook = new HSSFWorkbook();
		//�ڶ�������workbook�д���һ��sheet��Ӧexcel�е�sheet
		HSSFSheet sheet = workbook.createSheet("BOUserList");
		//����������sheet������ӱ�ͷ��0�У��ϰ汾��poi��sheet������������
		HSSFRow row = sheet.createRow(0);
		//���Ĳ���������Ԫ�����ñ�ͷ
		HSSFCell cell = row.createCell(0);
		cell.setCellValue("Account");
		cell = row.createCell(1);
		cell.setCellValue("Email");
		cell=row.createCell(2);
		cell.setCellValue("FullName");
		cell=row.createCell(3);
		cell.setCellValue("Groups");
		
		//���岽��д��ʵ�����ݣ�ʵ��Ӧ������Щ���ݴ����ݿ�õ�,�����װ���ݣ����ϰ����󡣶��������ֵ��Ӧ���ÿ�е�ֵ
		for (int i = 0; i <userList.size(); i++) 
		{
				HSSFRow row1 = sheet.createRow(i+1);
				BOUserInfo mo=userList.get(i);
				//������Ԫ����ֵ
				row1.createCell(0).setCellValue(mo.getUserAccount());
				row1.createCell(1).setCellValue(mo.getUserEmail());
				row1.createCell(2).setCellValue(mo.getUserFullName());
				row1.createCell(3).setCellValue(mo.getUserGroups());
		}
		//���ļ����浽ָ����λ��
		try 
		{
			FileOutputStream fos = new FileOutputStream(excelPath);
			workbook.write(fos);
			System.out.println("��ϲ����д��ɹ�������������");
			fos.close();
		} catch (IOException e) 
		{
			System.out.println("д���ļ���������");
			e.printStackTrace();
			return false;
		}

		
		return true;
	}


}
