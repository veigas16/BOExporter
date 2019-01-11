package com.boexporter.logic;

import java.util.Set;

public class BOUserInfo {
	
	private String userAccount="";
	private Set userGroups=null;
	private String userFullName="";
	private String userEmail="";
	public String getUserAccount() {
		return userAccount;
	}
	public void setUserAccount(String userAccount) {
		this.userAccount = userAccount;
	}
	public String getUserGroups() {
		String userGroupsString="";
		if(userGroups.size() > 0) {
			
			userGroupsString = userGroups.toString();
			
				
		}
		return userGroupsString;
		
	}
	public void setUserGroups(Set userGroups) {
		this.userGroups = userGroups;
	}
	public String getUserFullName() {
		return userFullName;
	}
	public void setUserFullName(String userFullName) {
		this.userFullName = userFullName;
	}
	public String getUserEmail() {
		return userEmail;
	}
	public void setUserEmail(String userEmail) {
		this.userEmail = userEmail;
	}
	
	

}
