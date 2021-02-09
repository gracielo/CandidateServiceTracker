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
@RequestMapping("/managers")
public class ManagerController {

	
	Workbook candidatesWorkBook;
	FileInputStream excelFile;
	private String path= "/app/src/main/resources/files/CandidatesTracker.xlsx";
	//private String path = "C:\\Users\\ALEJANDROBARRETOJIME\\git\\CandidateServiceTracker\\ApplicantTrackerServicesExcel\\src\\main\\resources\\files\\CandidatesTracker.xlsx";
	private String sheetName = "Managers";
	
	@GetMapping("/getManagers")
	private ResponseEntity<Object> getManagers(){
		try {
			List<Managers> managers = new ArrayList<>();
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			data.next();
			while(data.hasNext()) {
				Row currentRow = data.next();
				Managers manager = new Managers();
				manager.setManagerId((int)currentRow.getCell(0).getNumericCellValue());
				manager.setName(currentRow.getCell(1).getStringCellValue());
				manager.setEmail(currentRow.getCell(2).getStringCellValue());
				manager.setProject(currentRow.getCell(3).getStringCellValue());
				managers.add(manager);
			}
			return new ResponseEntity<>(managers,HttpStatus.OK);
			
		}catch(Exception e) {
			return new ResponseEntity<>("Error: "+e.getMessage(),HttpStatus.NOT_FOUND);
		}
	}
	
	@GetMapping("/getmanagerById/{id}")
	private ResponseEntity<Object> getEvaluationsByCandidate(@PathVariable("id") int id){
		try {
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			data.next();
			while(data.hasNext()) {
				Row currentRow = data.next();
				if ((int)currentRow.getCell(0).getNumericCellValue() == id) {
					Managers ev = new Managers();
					ev.setManagerId(id);
					ev.setName(currentRow.getCell(1).getStringCellValue());
					ev.setEmail(currentRow.getCell(2).getStringCellValue());
					ev.setProject(currentRow.getCell(3).getStringCellValue());
					
					return new ResponseEntity<>(ev,HttpStatus.OK);
				}
			}
			return new ResponseEntity<>("404 NOT Found",HttpStatus.NOT_FOUND);
			
		}catch(Exception e) {
			return new ResponseEntity<>("Error: " +e.getMessage(),HttpStatus.NOT_FOUND);
		}
	}
	
	@PostMapping("/registermanager")
	private ResponseEntity<Object> registerEvaluation(@RequestBody Managers manager){
		try {
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			int evId = worksheet.getLastRowNum() + 1;
			manager.setManagerId(evId);
			Row row = worksheet.createRow(evId);
			Cell cell = row.createCell(0);
			cell.setCellValue(evId);
			cell = row.createCell(1);
			cell.setCellValue(manager.getName());
			cell = row.createCell(2);
			cell.setCellValue(manager.getEmail());
			cell = row.createCell(3);
			cell.setCellValue(manager.getProject());
			
			FileOutputStream output = new FileOutputStream(path);
			candidatesWorkBook.write(output);
			candidatesWorkBook.close();
			return new ResponseEntity<>(manager, HttpStatus.CREATED);
		}catch(Exception e) {
			return new ResponseEntity<>("ERROR: "+e.getMessage(),HttpStatus.INTERNAL_SERVER_ERROR);
		}
	}
}
