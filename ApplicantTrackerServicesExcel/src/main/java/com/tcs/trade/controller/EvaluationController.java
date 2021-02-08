package com.tcs.trade.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import com.tcs.trade.model.*;

@RestController
@RequestMapping("/evaluations")
public class EvaluationController {

	
	Workbook candidatesWorkBook;
	FileInputStream excelFile;
	//@Value("${excel.path}")
	private String path= "/app/src/main/resources/files/CandidatesTracker.xlsx";
	//@Value("${excel.sheetNameEvaluations}")
	private String sheetName = "Evaluations";
	
	@GetMapping("/getEvaluationsByCandidate/{id}")
	private ResponseEntity<Object> getEvaluationsByCandidate(@PathVariable("id") int id){
		try {
			List<Evaluation> evaluations = new ArrayList<>();
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			while(data.hasNext()) {
				Row currentRow = data.next();
				if ((int)currentRow.getCell(0).getNumericCellValue() == id) {
					Evaluation ev = new Evaluation();
					ev.setEvaluationId((int)currentRow.getCell(0).getNumericCellValue());
					ev.setCandidateId(id);
					ev.setEvaluatorId((int)currentRow.getCell(2).getNumericCellValue());
					ev.setFeedback(currentRow.getCell(3).getStringCellValue());
					ev.setGrade((int)currentRow.getCell(4).getNumericCellValue());
					ev.setEvaluationDate(currentRow.getCell(5).getDateCellValue());
					evaluations.add(ev);
				}
			}
			return new ResponseEntity<>(evaluations,HttpStatus.OK);
			
		}catch(Exception e) {
			return new ResponseEntity<>("404 NOT Found",HttpStatus.NOT_FOUND);
		}
	}
	
	@PostMapping("/registerEvaluations")
	private ResponseEntity<Object> registerEvaluation(@RequestBody Evaluation evaluation){
		try {
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			int evId = worksheet.getLastRowNum() + 1;
			Row row = worksheet.createRow(evId);
			Cell cell = row.createCell(0);
			cell.setCellValue(evId);
			cell = row.createCell(1);
			cell.setCellValue(evaluation.getCandidateId());
			cell = row.createCell(2);
			cell.setCellValue(evaluation.getEvaluatorId());
			cell = row.createCell(3);
			cell.setCellValue(evaluation.getFeedback());
			cell = row.createCell(4);
			cell.setCellValue(evaluation.getGrade());
			cell = row.createCell(5);
			cell.setCellValue(getSystemDate());
			
			FileOutputStream output = new FileOutputStream(path);
			candidatesWorkBook.write(output);
			candidatesWorkBook.close();
			return new ResponseEntity<>(evaluation, HttpStatus.CREATED);
		}catch(Exception e) {
			return new ResponseEntity<>("ERROR: "+e.getMessage(),HttpStatus.INTERNAL_SERVER_ERROR);
		}
	}
	
	private String getSystemDate() {
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
		LocalDateTime now = LocalDateTime.now();
		return dtf.format(now);
	}
}
