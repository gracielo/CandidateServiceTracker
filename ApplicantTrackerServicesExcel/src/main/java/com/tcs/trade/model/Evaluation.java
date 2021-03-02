package com.tcs.trade.model;

import java.util.Date;

public class Evaluation {
	
	private int evaluationId;
	private int candidateId;
	private String evaluatorId;
	private String feedback;
	private int grade;
	private Date evaluationDate;
	private String status;
	private String evaluationType;
	
	
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
	public String getEvaluatorId() {
		return evaluatorId;
	}
	public void setEvaluatorId(String evaluatorId) {
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
	public String getStatus() {
		return status;
	}
	public void setStatus(String status) {
		this.status = status;
	}
	public String getEvaluationType() {
		return evaluationType;
	}
	public void setEvaluationType(String evaluationType) {
		this.evaluationType = evaluationType;
	}
	@Override
	public String toString() {
		return "Evaluation [evaluationId=" + evaluationId + ", candidateId=" + candidateId + ", evaluatorId="
				+ evaluatorId + ", feedback=" + feedback + ", grade=" + grade + ", evaluationDate=" + evaluationDate
				+ ", status=" + status + ", evaluationType=" + evaluationType + "]";
	}	
	
	
	
}
