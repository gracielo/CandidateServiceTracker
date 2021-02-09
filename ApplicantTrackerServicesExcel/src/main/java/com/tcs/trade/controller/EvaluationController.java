package com.tcs.trade.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
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
	//private String path = "C:\\Users\\ALEJANDROBARRETOJIME\\git\\CandidateServiceTracker\\ApplicantTrackerServicesExcel\\src\\main\\resources\\files\\CandidatesTracker.xlsx";
	//@Value("${excel.sheetNameEvaluations}")
	private String sheetName = "Evaluations";
	private String sheetNameCandidates = "Candidates";
	private String sheetNameEvaluators = "Evaluators";
	
	@GetMapping("/getEvaluationsByCandidateId/{id}")
	private ResponseEntity<Object> getEvaluationsByCandidate(@PathVariable("id") int id){
		try {
			List<Evaluation> evaluations = new ArrayList<>();
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			data.next();
			while(data.hasNext()) {
				Row currentRow = data.next();
				if ((int)currentRow.getCell(0).getNumericCellValue() == id) {
					Evaluation ev = new Evaluation();
					ev.setEvaluationId((int)currentRow.getCell(0).getNumericCellValue());
					ev.setCandidateId(id);
					ev.setEvaluatorId(getEvaluatorName(currentRow.getCell(2).getStringCellValue()));
					ev.setFeedback(currentRow.getCell(3).getStringCellValue());
					ev.setGrade((int)currentRow.getCell(4).getNumericCellValue());
					ev.setEvaluationDate(new Date(currentRow.getCell(5).getStringCellValue()));
					evaluations.add(ev);
				}
			}
			return new ResponseEntity<>(evaluations,HttpStatus.OK);
			
		}catch(Exception e) {
			e.printStackTrace();
			return new ResponseEntity<>("Error: " +e.getMessage(),HttpStatus.NOT_FOUND);
		}
	}
	
	@PostMapping("/registerEvaluation")
	private ResponseEntity<Object> registerEvaluation(@RequestBody Evaluation evaluation){
		try {
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			int evId = worksheet.getLastRowNum() + 1;
			evaluation.setEvaluationId(evId);
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
			Sheet candidate = candidatesWorkBook.getSheet(sheetNameCandidates);
			Iterator<Row> data = candidate.iterator();
			data.next();
			while (data.hasNext()) {
				Row rc = data.next();
				int aux  = (int) rc.getCell(0).getNumericCellValue();
				if(aux == evaluation.getCandidateId()) {
					Cell c2 = rc.getCell(7);
					c2.setCellValue(evaluation.getStatus());
				}
			}
			
			
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
	
	private String getEvaluatorName(String id) {
		Sheet manager = candidatesWorkBook.getSheet(sheetNameEvaluators);
		Iterator<Row> data = manager.iterator();
		data.next();
		while(data.hasNext()) {
			Row row = data.next();
			if (row.getCell(0).getCellTypeEnum()==CellType.STRING) {
				String mani =row.getCell(0).getStringCellValue();
				if (mani.compareToIgnoreCase(id)==0) {
					return row.getCell(1).getStringCellValue();
				}
			}else {
				int mani = (int) row.getCell(0).getNumericCellValue();
				if(mani==Integer.parseInt(id)) {
					return row.getCell(1).getStringCellValue();
				}
			}
			
		}
		return id;
	}
}
