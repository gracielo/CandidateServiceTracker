package com.tcs.trade.model;

import java.util.Date;
import java.util.List;

public class CandidateUpdates {

	private CandidateExcel candidateInfo;
	private List<Update> canidateUpdate;
	
	public CandidateExcel getCandidateInfo() {
		return candidateInfo;
	}
	public void setCandidateInfo(CandidateExcel candidateInfo) {
		this.candidateInfo = candidateInfo;
	}
	public List<Update> getCanidateUpdate() {
		return canidateUpdate;
	}
	public void setCanidateUpdate(List<Update> canidateUpdate) {
		this.canidateUpdate = canidateUpdate;
	}
	
}
