package com.tcs.trade.utils;

import java.util.function.Predicate;

import com.tcs.trade.model.Evaluation;

public class Utilities {
	
	public static Predicate<Evaluation> filterByCandidateId(Integer id){
		return (Evaluation e)->{
			return new Integer(e.getCandidateId()).equals(id);
		};
	}
}
