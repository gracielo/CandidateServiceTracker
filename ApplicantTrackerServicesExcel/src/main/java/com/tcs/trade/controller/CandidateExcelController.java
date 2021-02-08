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
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.DeleteMapping;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import com.tcs.trade.model.CandidateExcel;
import com.tcs.trade.model.Evaluation;

@RestController
public class CandidateExcelController {

	Workbook candidatesWorkBook;
	FileInputStream excelFile;
	ClassLoader classLoader;
	//@Value("${excel.path}")
	private String path= "/app/src/main/resources/files/CandidatesTracker.xlsx";
	//@Value("${excel.sheetName}")
	private String sheetName = "Candidates";
	//@Value("${excel.sheetNameEvaluations}")
	private String sheetNameEvaluations = "Evaluations";
	//@Value("${excel.sheetNameEvaluators}")
	private String sheetNameEvaluators= "Evaluators";
	//@Value("${excel.sheetNameManagers}")
	private String sheetNameManagers= "Managers";

	@PostMapping("/registerCandidate")
	public ResponseEntity<CandidateExcel> registerCandidate(@RequestBody CandidateExcel candidate) {
		try {
			if (candidate.getStatus() == null) {
				candidate.setStatus("ACTIVE");
			}
			classLoader = getClass().getClassLoader();
			// excelFile = new FileInputStream(new
			// File(classLoader.getResource(path).getFile()));
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			int rowIndex = worksheet.getLastRowNum() + 1;
			Row row = worksheet.createRow(rowIndex);
			Cell cell = row.createCell(0);
			cell.setCellValue(rowIndex);
			cell = row.createCell(1);
			cell.setCellValue(candidate.getName());
			cell = row.createCell(2);
			cell.setCellValue(candidate.getEmail());
			cell = row.createCell(3);
			cell.setCellValue(candidate.getPhone());
			cell = row.createCell(4);
			cell.setCellValue(candidate.getProfile());
			cell = row.createCell(5);
			cell.setCellValue(candidate.getYearsOfExperience());
			cell = row.createCell(6);
			cell.setCellValue(candidate.getEnglishLevel());
			cell = row.createCell(7);
			cell.setCellValue(candidate.getStatus());
			cell = row.createCell(8);
			cell.setCellValue(this.getSystemDate());
			cell = row.createCell(9);
			cell.setCellValue(candidate.getSkills());
			cell = row.createCell(10);
			cell.setCellValue(candidate.getAging());
			cell = row.createCell(11);
			cell.setCellValue(candidate.getManager());
			cell = row.createCell(12);
			cell.setCellValue(candidate.isRelocated());

			FileOutputStream output = new FileOutputStream(path);
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

	@GetMapping("/getCandidateInfoById/{id}")
	public ResponseEntity<Object> getCandidateInfoById(@PathVariable("id") String id) {
		try {
			CandidateExcel can = null;
			classLoader = getClass().getClassLoader();
			excelFile = new FileInputStream(new File(path)); // new FileInputStream(new
																// File(classLoader.getResource(path).getFile()));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			while (data.hasNext()) {
				Row currentRow = data.next();
				String aux = currentRow.getCell(0).getStringCellValue();
				if (aux.compareToIgnoreCase(id) == 0) {
					can = new CandidateExcel();
					can.setName(currentRow.getCell(1).getStringCellValue());
					can.setEmail(currentRow.getCell(2).getStringCellValue());
					can.setPhone((int) currentRow.getCell(3).getNumericCellValue());
					can.setProfile(currentRow.getCell(4).getStringCellValue());
					can.setYearsOfExperience((int) currentRow.getCell(5).getNumericCellValue());
					can.setEnglishLevel(currentRow.getCell(6).getStringCellValue());
					can.setStatus(currentRow.getCell(7).getStringCellValue());
					can.setCreationDate(new Date(currentRow.getCell(8).getStringCellValue()));
					can.setSkills(currentRow.getCell(9).getStringCellValue());
					can.setAging((int) currentRow.getCell(10).getNumericCellValue());
					can.setManager(currentRow.getCell(11).getStringCellValue());
					if (currentRow.getCell(12).getStringCellValue().compareTo("true") == 0) {
						can.setRelocated(true);
					} else {
						can.setRelocated(false);
					}

					candidatesWorkBook.close();
					return new ResponseEntity<>(can, HttpStatus.OK);
				}
			}
			return new ResponseEntity<>("Candidate NOT FOUNDED", HttpStatus.NOT_FOUND);
		} catch (Exception e) {
			e.printStackTrace();
			return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
		}
	}

	@PutMapping("/updateCandidate/{id}")
	public ResponseEntity<Object> updateCandidate(@RequestBody CandidateExcel candidate,
			@PathVariable("id") String id) {
		try {
			classLoader = getClass().getClassLoader();
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			while (data.hasNext()) {
				Row row = data.next();

				String aux = row.getCell(0).getStringCellValue();
				if (aux.compareToIgnoreCase(id) == 0) {
					Cell cell = row.getCell(1);
					cell.setCellValue(candidate.getName() != null ? candidate.getName() : cell.getStringCellValue());
					cell = row.getCell(2);
					cell.setCellValue(candidate.getEmail() != null ? candidate.getEmail() : cell.getStringCellValue());
					cell = row.getCell(3);
					cell.setCellValue(candidate.getPhone() != 0 ? candidate.getPhone() : cell.getNumericCellValue());
					cell = row.getCell(4);
					cell.setCellValue(
							candidate.getProfile() != null ? candidate.getProfile() : cell.getStringCellValue());
					cell = row.getCell(5);
					cell.setCellValue(candidate.getYearsOfExperience() != 0 ? candidate.getYearsOfExperience()
							: cell.getNumericCellValue());
					cell = row.getCell(6);
					cell.setCellValue(candidate.getEnglishLevel() != null ? candidate.getEnglishLevel()
							: cell.getStringCellValue());
					cell = row.getCell(7);
					cell.setCellValue(
							candidate.getStatus() != null ? candidate.getStatus() : cell.getStringCellValue());
					cell = row.getCell(9);
					cell.setCellValue(
							candidate.getSkills() != null ? candidate.getSkills() : cell.getStringCellValue());
					cell = row.getCell(10);
					cell.setCellValue(candidate.getAging() != 0 ? candidate.getAging() : cell.getNumericCellValue());
					cell = row.getCell(12);
					cell.setCellValue(
							candidate.getManager() != null ? candidate.getManager() : cell.getStringCellValue());
					cell = row.getCell(12);
					if (candidate.isRelocated()) {
						cell.setCellValue(candidate.isRelocated());
					}
				}

			}
			FileOutputStream output = new FileOutputStream(path);
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
			System.err.println(sheetName);
			classLoader = this.getClass().getClassLoader();
			File a = new File(path);// (classLoader.getResource(path).getFile());
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
				can.setName(currentRow.getCell(1).getStringCellValue());
				can.setEmail(currentRow.getCell(2).getStringCellValue());
				can.setPhone((int) currentRow.getCell(3).getNumericCellValue());
				can.setProfile(currentRow.getCell(4).getStringCellValue());
				can.setYearsOfExperience((int) currentRow.getCell(5).getNumericCellValue());
				can.setEnglishLevel(currentRow.getCell(6).getStringCellValue());
				can.setStatus(currentRow.getCell(7).getStringCellValue());
				can.setCreationDate(new Date(currentRow.getCell(8).getStringCellValue()));
				can.setSkills(currentRow.getCell(9).getStringCellValue());
				can.setAging((int) currentRow.getCell(10).getNumericCellValue());
				can.setManager(currentRow.getCell(11).getStringCellValue());
				if (currentRow.getCell(12).getStringCellValue().compareToIgnoreCase("true") == 0) {
					can.setRelocated(true);
				} else {
					can.setRelocated(false);
				}

				candidates.add(can);
			}
			candidatesWorkBook.close();
			return new ResponseEntity<>(candidates.stream().filter(cand -> !"inactive".equals(cand.getStatus()))
					.collect(Collectors.toList()), HttpStatus.OK);
		} catch (

		Exception e) {
			e.printStackTrace();
			return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
		}

	}

	@PostMapping("/registerEvaluation/{id}")
	public ResponseEntity<Object> registerEvaluation(@RequestBody Evaluation evaluation,
			@PathVariable("id") String id) {
		try {
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Sheet worksheetEv = candidatesWorkBook.getSheet(sheetNameEvaluations);
			Iterator<Row> data = worksheet.iterator();
			CandidateExcel existenteCandidate = null;
			int rowNumber = 0;
			while (data.hasNext()) {
				Row currentRow = data.next();
				String aux = currentRow.getCell(0).getStringCellValue();
				if (aux.compareToIgnoreCase(id) == 0) {
					rowNumber = currentRow.getRowNum();
					existenteCandidate = new CandidateExcel();
					existenteCandidate.setCandidateId((int) currentRow.getCell(0).getNumericCellValue());
					/*
					 * existenteCandidate.setName(currentRow.getCell(0).getStringCellValue());
					 * existenteCandidate.setEmail(currentRow.getCell(1).getStringCellValue());
					 * existenteCandidate.setPhone((int)
					 * currentRow.getCell(2).getNumericCellValue());
					 * existenteCandidate.setProfile(currentRow.getCell(3).getStringCellValue());
					 * existenteCandidate.setYearsOfExperience((int)
					 * currentRow.getCell(4).getNumericCellValue());
					 * existenteCandidate.setEnglishLevel(currentRow.getCell(5).getStringCellValue()
					 * ); existenteCandidate.setStatus(currentRow.getCell(6).getStringCellValue());
					 * existenteCandidate.setCreationDate(new
					 * Date(currentRow.getCell(7).getStringCellValue()));
					 * existenteCandidate.setGrade((int)
					 * currentRow.getCell(8).getNumericCellValue());
					 * existenteCandidate.setEvaluator(currentRow.getCell(9).getStringCellValue());
					 * existenteCandidate.setFeedback(currentRow.getCell(10).getStringCellValue());
					 * existenteCandidate.setSkills(currentRow.getCell(11).getStringCellValue());
					 * existenteCandidate.setAging((int)
					 * currentRow.getCell(12).getNumericCellValue());
					 */
				}
			}

			if (existenteCandidate != null) {
				int idEv = worksheetEv.getLastRowNum() + 1;
				Row row = worksheetEv.createRow(idEv);
				Cell cell = row.createCell(0);
				cell.setCellValue(idEv);
				cell = row.createCell(1);
				cell.setCellValue(id);
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
			} else {
				return new ResponseEntity<>("Candidate doesn't exists", HttpStatus.NOT_FOUND);
			}

			return new ResponseEntity<>("Evaluation registered", HttpStatus.CREATED);
		} catch (Exception e) {
			return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
		}
	}

	@DeleteMapping("/deleteCandidate/{id}")
	private ResponseEntity<Object> deleteCandidate(@PathVariable("id") String id) {
		try {
			classLoader = getClass().getClassLoader();
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			ArrayList<Integer> rowsNumbers = new ArrayList<>();
			while (data.hasNext()) {
				Row currentRow = data.next();
				String aux = currentRow.getCell(1).getStringCellValue();
				if (aux.compareToIgnoreCase(id) == 0) {
					rowsNumbers.add(currentRow.getRowNum());
				}
			}
			for (Integer integer : rowsNumbers) {
				worksheet.removeRow(worksheet.getRow(integer));
			}

			FileOutputStream output = new FileOutputStream(path);
			candidatesWorkBook.write(output);
			candidatesWorkBook.close();
			return new ResponseEntity<>("Candidate Eliminated", HttpStatus.OK);
		} catch (Exception e) {
			e.printStackTrace();
			return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
		}
	}

}
