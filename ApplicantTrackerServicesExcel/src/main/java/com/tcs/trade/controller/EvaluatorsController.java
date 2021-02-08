package com.tcs.trade.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
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
@RequestMapping("/evaluators")
public class EvaluatorsController {

	
	Workbook candidatesWorkBook;
	FileInputStream excelFile;
	private String path= "/app/src/main/resources/files/CandidatesTracker.xlsx";
	private String sheetName = "Evaluators";
	
	@GetMapping("/getEvaluators")
	private ResponseEntity<Object> getEvaluators(){
		try {
			List<Evaluators> evaluators = new ArrayList<>();
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			data.next();
			while(data.hasNext()) {
				Row currentRow = data.next();
				Evaluators evaluator = new Evaluators();
				evaluator.setEvaluatorId((int)currentRow.getCell(0).getNumericCellValue());
				evaluator.setName(currentRow.getCell(1).getStringCellValue());
				evaluator.setMail(currentRow.getCell(2).getStringCellValue());
				evaluator.setPosition(currentRow.getCell(3).getStringCellValue());
				evaluators.add(evaluator);
			}
			return new ResponseEntity<>(evaluators,HttpStatus.OK);
			
		}catch(Exception e) {
			return new ResponseEntity<>("404 NOT Found",HttpStatus.NOT_FOUND);
		}
	}
	
	@GetMapping("/getEvaluatorById/{id}")
	private ResponseEntity<Object> getEvaluatorById(@PathVariable("id") int id){
		try {
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			data.next();
			while(data.hasNext()) {
				Row currentRow = data.next();
				if ((int)currentRow.getCell(0).getNumericCellValue() == id) {
					Evaluators ev = new Evaluators();
					ev.setEvaluatorId(id);
					ev.setName(currentRow.getCell(1).getStringCellValue());
					ev.setMail(currentRow.getCell(2).getStringCellValue());
					ev.setPosition(currentRow.getCell(3).getStringCellValue());
					
					return new ResponseEntity<>(ev,HttpStatus.OK);
				}
			}
			return new ResponseEntity<>("404 NOT Found",HttpStatus.NOT_FOUND);
			
		}catch(Exception e) {
			return new ResponseEntity<>("Error: " +e.getMessage(),HttpStatus.NOT_FOUND);
		}
	}
	
	@PostMapping("/registerEvaluator")
	private ResponseEntity<Object> registerEvaluator(@RequestBody Evaluators evaluator){
		try {
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			int evId = worksheet.getLastRowNum() + 1;
			Row row = worksheet.createRow(evId);
			Cell cell = row.createCell(0);
			cell.setCellValue(evId);
			cell = row.createCell(1);
			cell.setCellValue(evaluator.getName());
			cell = row.createCell(2);
			cell.setCellValue(evaluator.getMail());
			cell = row.createCell(3);
			cell.setCellValue(evaluator.getPosition());
			
			FileOutputStream output = new FileOutputStream(path);
			candidatesWorkBook.write(output);
			candidatesWorkBook.close();
			return new ResponseEntity<>(evaluator, HttpStatus.CREATED);
		}catch(Exception e) {
			return new ResponseEntity<>("ERROR: "+e.getMessage(),HttpStatus.INTERNAL_SERVER_ERROR);
		}
	}
}
