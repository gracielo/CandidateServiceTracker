package com.tcs.trade.model;

import java.util.Date;

public class CandidateExcel {
		
	private int candidateId;
	private String name;
	private String email;	
	private String phone;	
	private String profile;		
	private int yearsOfExperience;	
	private String englishLevel;	
	private String status;	
	private Date creationDate;
	private String skills;
	private int aging;
	private String manager;
	private boolean relocated;
	
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getEmail() {
		return email;
	}
	public void setEmail(String email) {
		this.email = email;
	}
	public String getPhone() {
		return phone;
	}
	public void setPhone(String phone) {
		this.phone = phone;
	}
	public String getProfile() {
		return profile;
	}
	public void setProfile(String profile) {
		this.profile = profile;
	}
	public int getYearsOfExperience() {
		return yearsOfExperience;
	}
	public void setYearsOfExperience(int yearsOfExperience) {
		this.yearsOfExperience = yearsOfExperience;
	}
	public String getEnglishLevel() {
		return englishLevel;
	}
	public void setEnglishLevel(String englishLevel) {
		this.englishLevel = englishLevel;
	}
	public String getStatus() {
		return status;
	}
	public void setStatus(String status) {
		this.status = status;
	}
	public Date getCreationDate() {
		return creationDate;
	}
	public void setCreationDate(Date creationDate) {
		this.creationDate = creationDate;
	}
	
	public String getSkills() {
		return skills;
	}
	public void setSkills(String skills) {
		this.skills = skills;
	}
	public int getAging() {
		return aging;
	}
	public void setAging(int aging) {
		this.aging = aging;
	}
	public int getCandidateId() {
		return candidateId;
	}
	public void setCandidateId(int candidateId) {
		this.candidateId = candidateId;
	}
	public String getManager() {
		return manager;
	}
	public void setManager(String manager) {
		this.manager = manager;
	}
	public boolean isRelocated() {
		return relocated;
	}
	public void setRelocated(boolean relocated) {
		this.relocated = relocated;
	}
	
	
	
}
