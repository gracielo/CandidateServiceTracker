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
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import com.tcs.trade.model.CandidateExcel;

@RestController
public class CandidateExcelController {

	Workbook candidatesWorkBook;
	FileInputStream excelFile; 
	ClassLoader classLoader;
	private static final String path = "file:/app/target/ApplicantTrackerServicesExcel-0.0.1-SNAPSHOT.jar!/BOOT-INF/classes!/files/CandidatesTracker.xlsx";//"files/CandidatesTracker.xlsx";
	private static final String sheetName = "Candidates";

	@PostMapping("/registerCandidate")
	public ResponseEntity<CandidateExcel> registerCandidate(@RequestBody CandidateExcel candidate) {
		try {
			if (candidate.getStatus() == null) {
				candidate.setStatus("IE Pending");
			}
			classLoader= getClass().getClassLoader();
			System.out.println(classLoader.getResource(path).getPath());
			excelFile = new FileInputStream(new File(classLoader.getResource(path).getFile()));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			int rowIndex = worksheet.getLastRowNum() + 1;
			Row row = worksheet.createRow(rowIndex);
			Cell cell = row.createCell(0);
			cell.setCellValue(candidate.getName());
			cell = row.createCell(1);
			cell.setCellValue(candidate.getEmail());
			cell = row.createCell(2);
			cell.setCellValue(candidate.getPhone());
			cell = row.createCell(3);
			cell.setCellValue(candidate.getProfile());
			cell = row.createCell(4);
			cell.setCellValue(candidate.getYearsOfExperience());
			cell = row.createCell(5);
			cell.setCellValue(candidate.getEnglishLevel());
			cell = row.createCell(6);
			cell.setCellValue(candidate.getStatus());
			cell = row.createCell(7);
			cell.setCellValue(this.getSystemDate());
			cell = row.createCell(8);
			cell.setCellValue(candidate.getGrade());
			cell = row.createCell(9);
			cell.setCellValue(candidate.getEvaluator());
			cell = row.createCell(10);
			cell.setCellValue(candidate.getFeedback());
			cell = row.createCell(11);
			cell.setCellValue(candidate.getSkills());
			cell = row.createCell(12);
			cell.setCellValue(candidate.getAging());

			FileOutputStream output = new FileOutputStream(classLoader.getSystemResource(path).getPath());
			candidatesWorkBook.write(output);
			candidatesWorkBook.close();
			return new ResponseEntity<>(candidate, HttpStatus.CREATED);
		} catch (Exception e) {
			e.printStackTrace();
			return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
		}
	}

	private String getSystemDate() {
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
		LocalDateTime now = LocalDateTime.now();
		return dtf.format(now);
	}

	@GetMapping("/getCandidateInfoById/{email}")
	public ResponseEntity<Object> getCandidateInfoById(@PathVariable("email") String email) {
		try {
			CandidateExcel can = new CandidateExcel();
			classLoader= getClass().getClassLoader();
			excelFile = new FileInputStream(new File(classLoader.getResource(path).getFile()));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			while (data.hasNext()) {
				Row currentRow = data.next();
				String aux = currentRow.getCell(1).getStringCellValue();
				if (aux.compareToIgnoreCase(email) == 0) {
					can.setName(currentRow.getCell(0).getStringCellValue());
					can.setEmail(currentRow.getCell(1).getStringCellValue());
					can.setPhone((int) currentRow.getCell(2).getNumericCellValue());
					can.setProfile(currentRow.getCell(3).getStringCellValue());
					can.setYearsOfExperience((int) currentRow.getCell(4).getNumericCellValue());
					can.setEnglishLevel(currentRow.getCell(5).getStringCellValue());
					can.setStatus(currentRow.getCell(6).getStringCellValue());
					can.setCreationDate(new Date(currentRow.getCell(7).getStringCellValue()));
					can.setGrade((int) currentRow.getCell(8).getNumericCellValue());
					can.setEvaluator(currentRow.getCell(9).getStringCellValue());
					can.setFeedback(currentRow.getCell(10).getStringCellValue());
					can.setSkills(currentRow.getCell(11).getStringCellValue());
					can.setAging((int) currentRow.getCell(12).getNumericCellValue());
				}
			}
			candidatesWorkBook.close();
			return new ResponseEntity<>(can, HttpStatus.OK);
		} catch (Exception e) {
			e.printStackTrace();
			return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
		}

	}

	@PutMapping("/updateCandidate/{email}")
	public ResponseEntity<Object> updateCandidate(@RequestBody CandidateExcel candidate,
			@PathVariable("email") String email) {
		try {
			classLoader= getClass().getClassLoader();
			excelFile = new FileInputStream(new File(classLoader.getResource(path).getFile()));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			while (data.hasNext()) {
				Row row = data.next();

				String aux = row.getCell(1).getStringCellValue();
				if (aux.compareToIgnoreCase(email) == 0) {
					Cell cell = row.getCell(0);
					cell.setCellValue(candidate.getName() != null ? candidate.getName() : cell.getStringCellValue());
					cell = row.getCell(1);
					cell.setCellValue(candidate.getEmail() != null ? candidate.getEmail() : cell.getStringCellValue());
					cell = row.getCell(2);
					cell.setCellValue(candidate.getPhone() != 0 ? candidate.getPhone() : cell.getNumericCellValue());
					cell = row.getCell(3);
					cell.setCellValue(
							candidate.getProfile() != null ? candidate.getProfile() : cell.getStringCellValue());
					cell = row.getCell(4);
					cell.setCellValue(candidate.getYearsOfExperience() != 0 ? candidate.getYearsOfExperience()
							: cell.getNumericCellValue());
					cell = row.getCell(5);
					cell.setCellValue(candidate.getEnglishLevel() != null ? candidate.getEnglishLevel()
							: cell.getStringCellValue());
					cell = row.getCell(6);
					cell.setCellValue(
							candidate.getStatus() != null ? candidate.getStatus() : cell.getStringCellValue());
					cell = row.getCell(8);
					cell.setCellValue(candidate.getGrade() != 0 ? candidate.getGrade() : cell.getNumericCellValue());
					cell = row.getCell(9);
					cell.setCellValue(
							candidate.getEvaluator() != null ? candidate.getEvaluator() : cell.getStringCellValue());
					cell = row.getCell(10);
					cell.setCellValue(candidate.getFeedback());
					cell = row.getCell(11);
					cell.setCellValue(candidate.getSkills());
					cell = row.getCell(12);
					cell.setCellValue(candidate.getAging() != 0 ? candidate.getAging()
							: cell.getNumericCellValue());
				}

			}
			FileOutputStream output = new FileOutputStream(classLoader.getSystemResource(path).getPath());
			candidatesWorkBook.write(output);
			candidatesWorkBook.close();
			return new ResponseEntity<>(candidate, HttpStatus.OK);
		} catch (Exception e) {
			e.printStackTrace();
			return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
		}
	}

	@GetMapping("/getCandidates")
	public ResponseEntity<Object> getCandidates() {
		try {
			classLoader= this.getClass().getClassLoader();
			File a = new File(classLoader.getResource(path).getFile());
			excelFile = new FileInputStream(a);
			CandidateExcel can;
			List<CandidateExcel> candidates = new ArrayList<>();
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			data.next();
			while (data.hasNext()) {
				Row currentRow = data.next();
				can = new CandidateExcel();
				can.setName(currentRow.getCell(0).getStringCellValue());
				can.setEmail(currentRow.getCell(1).getStringCellValue());
				can.setPhone((int)currentRow.getCell(2).getNumericCellValue());
				can.setProfile(currentRow.getCell(3).getStringCellValue());
				can.setYearsOfExperience((int) currentRow.getCell(4).getNumericCellValue());
				can.setEnglishLevel(currentRow.getCell(5).getStringCellValue());
				can.setStatus(currentRow.getCell(6).getStringCellValue());
				can.setCreationDate(new Date(currentRow.getCell(7).getStringCellValue()));
				can.setGrade((int) currentRow.getCell(8).getNumericCellValue());
				can.setEvaluator(currentRow.getCell(9).getStringCellValue());
				can.setFeedback(currentRow.getCell(10).getStringCellValue());
				can.setSkills(currentRow.getCell(11).getStringCellValue());
				can.setAging((int) currentRow.getCell(12).getNumericCellValue());
				candidates.add(can);
			}
			candidatesWorkBook.close();
			return new ResponseEntity<>(candidates.stream().filter(cand->!"inactive".equals(cand.getStatus())).collect(Collectors.toList()), HttpStatus.OK);
		} catch (

		Exception e) {
			e.printStackTrace();
			return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
		}

	}
	
	@GetMapping("createFile")
	private void createFile() {
		try {
		classLoader= this.getClass().getClassLoader();
		File archivo = new File(classLoader.getResource(path).getPath()+"/prueba.txt");
		}catch(Exception e) {
			e.printStackTrace();
		}
	}

}
