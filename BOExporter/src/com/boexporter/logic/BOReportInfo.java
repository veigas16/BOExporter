package com.boexporter.logic;

import java.io.UnsupportedEncodingException;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.Set;

import com.crystaldecisions.celib.properties.URLDecoder;

public class BOReportInfo {
	private static final SimpleDateFormat sdf = new SimpleDateFormat("yyyy.MM.dd.HH.mm.ss");
	
	private String reportId = "";
	private String reportName= "";
	private String reportType = "";
	private String reportCreateTime = "";
	private String reportUpdateTime = "";
	private String reportParentId = "";
	private String reportPath="";
	public String getReportId() {
		return reportId;
	}
	public void setReportId(String reportId) {
		this.reportId = reportId;
	}
	public String getReportName() {
		return URLDecoder.decode(reportName);
				
	}
	public void setReportName(String reportName) {
		this.reportName = reportName;
	}
	public String getReportType() {
		return reportType;
	}
	public void setReportType(String reportType) {
		this.reportType = reportType;
	}
	public String getReportCreateTime() {
		
		return reportCreateTime;
	}
	
	public String convertTs(String tsString) {
		long longTimestamp = Long.parseLong(tsString);
		Timestamp timestamp = new Timestamp(longTimestamp);
		return "";
		
	}
	public void setReportCreateTime(String reportCreateTime) {
		this.reportCreateTime = reportCreateTime;
	}
	public String getReportUpdateTime() {
		return reportUpdateTime;
	}
	public void setReportUpdateTime(String reportUpdateTime) {
		this.reportUpdateTime = reportUpdateTime;
	}
	public String getReportParentId() {
		return reportParentId;
	}
	public void setReportParentId(String reportParentId) {
		this.reportParentId = reportParentId;
	}
	public String getReportPath() {
		return reportPath;
	}
	public void setReportPath(String reportPath) {
		this.reportPath = reportPath;
	}
	
	public String getDetailString() {
		return this.reportId + "," + this.getReportName() + "," + this.reportType + ","+this.reportCreateTime + ","+ this.reportUpdateTime + ","+ this.reportParentId + ","+ this.reportPath;
	}
	


	

}
