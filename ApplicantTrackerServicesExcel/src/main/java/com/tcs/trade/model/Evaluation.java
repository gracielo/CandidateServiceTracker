package com.tcs.trade.model;

import java.util.Date;

public class Evaluation {
	
	private int evaluationId;
	private int candidateId;
	private int evaluatorId;
	private String feedback;
	private int grade;
	private Date evaluationDate;
	
	
	
	public int getEvaluationId() {
		return evaluationId;
	}
	public void setEvaluationId(int evaluationId) {
		this.evaluationId = evaluationId;
	}
	public int getCandidateId() {
		return candidateId;
	}
	public void setCandidateId(int candidateId) {
		this.candidateId = candidateId;
	}
	public int getEvaluatorId() {
		return evaluatorId;
	}
	public void setEvaluatorId(int evaluatorId) {
		this.evaluatorId = evaluatorId;
	}
	public String getFeedback() {
		return feedback;
	}
	public void setFeedback(String feedback) {
		this.feedback = feedback;
	}
	public int getGrade() {
		return grade;
	}
	public void setGrade(int grade) {
		this.grade = grade;
	}
	public Date getEvaluationDate() {
		return evaluationDate;
	}
	public void setEvaluationDate(Date evaluationDate) {
		this.evaluationDate = evaluationDate;
	}
	
	
	
}
