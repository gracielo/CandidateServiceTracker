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
	//private String path = "C:\\Users\\ALEJANDROBARRETOJIME\\git\\CandidateServiceTracker\\ApplicantTrackerServicesExcel\\src\\main\\resources\\files\\CandidatesTracker.xlsx";
	//@Value("${excel.sheetName}")
	private String sheetName = "Candidates";
	//@Value("${excel.sheetNameEvaluations}")
	private String sheetNameEvaluations = "Evaluations";
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
			candidate.setCandidateId(rowIndex);
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
			
			Sheet manager = candidatesWorkBook.getSheet(sheetNameManagers);
			Iterator<Row> dataMan = manager.iterator();
			dataMan.next();
			while(dataMan.hasNext()) {
				Row rowMan = dataMan.next();
				if (rowMan.getCell(0).getCellTypeEnum()==CellType.STRING) {
					String mani =rowMan.getCell(0).getStringCellValue();
					if (mani.compareToIgnoreCase(candidate.getManager())==0) {
						candidate.setManager(rowMan.getCell(1).getStringCellValue());
					}
				}else {
					int mani = (int) rowMan.getCell(0).getNumericCellValue();
					if(mani==Integer.parseInt(candidate.getManager())) {
						candidate.setManager(rowMan.getCell(1).getStringCellValue());
					}
				}
				
			}

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
	public ResponseEntity<Object> getCandidateInfoById(@PathVariable("id") int id) {
		try {
			CandidateExcel can = null;
			classLoader = getClass().getClassLoader();
			excelFile = new FileInputStream(new File(path)); // new FileInputStream(new
																// File(classLoader.getResource(path).getFile()));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Sheet manager = candidatesWorkBook.getSheet(sheetNameManagers);
			Iterator<Row> data = worksheet.iterator();
			Iterator<Row> dataMan = manager.iterator();
			data.next();
			dataMan.next();
			while (data.hasNext()) {
				Row currentRow = data.next();
				int aux = (int) currentRow.getCell(0).getNumericCellValue();
				if (aux == id) {
					can = new CandidateExcel();
					can.setCandidateId(id);
					can.setName(currentRow.getCell(1).getStringCellValue());
					can.setEmail(currentRow.getCell(2).getStringCellValue());
					if (currentRow.getCell(3).getCellTypeEnum() == CellType.NUMERIC) {
						can.setPhone(currentRow.getCell(3).getNumericCellValue()+"");
					}else if(currentRow.getCell(3).getCellTypeEnum() == CellType.STRING) {
						can.setPhone(currentRow.getCell(3).getStringCellValue());
					}					
					can.setProfile(currentRow.getCell(4).getStringCellValue());
					can.setYearsOfExperience((int) currentRow.getCell(5).getNumericCellValue());
					can.setEnglishLevel(currentRow.getCell(6).getStringCellValue());
					can.setStatus(currentRow.getCell(7).getStringCellValue());
					can.setCreationDate(new Date(currentRow.getCell(8).getStringCellValue()));
					can.setSkills(currentRow.getCell(9).getStringCellValue());
					can.setAging((int) currentRow.getCell(10).getNumericCellValue());
					
					while(dataMan.hasNext()) {
						Row rowMan = dataMan.next();
						if (rowMan.getCell(0).getCellTypeEnum()==CellType.STRING) {
							String mani =rowMan.getCell(0).getStringCellValue();
							if (mani.compareToIgnoreCase(can.getManager())==0) {
								can.setManager(rowMan.getCell(1).getStringCellValue());
							}
						}else {
							int mani = (int) rowMan.getCell(0).getNumericCellValue();
							if(mani== Integer.parseInt(currentRow.getCell(11).getStringCellValue() )) {
								can.setManager(rowMan.getCell(1).getStringCellValue());
							}
						}
					}
					can.setRelocated(currentRow.getCell(12).getBooleanCellValue());

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
			@PathVariable("id") int id) {
		try {
			classLoader = getClass().getClassLoader();
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			data.next();
			while (data.hasNext()) {
				Row row = data.next();

				int aux = (int) row.getCell(0).getNumericCellValue();
				if (aux == id) {
					Cell cell = row.getCell(1);
					cell.setCellValue(candidate.getName() != null ? candidate.getName() : cell.getStringCellValue());
					cell = row.getCell(2);
					cell.setCellValue(candidate.getEmail() != null ? candidate.getEmail() : cell.getStringCellValue());
					cell = row.getCell(3);
					if (cell.getCellTypeEnum() == CellType.NUMERIC) {
						cell.setCellValue(candidate.getPhone() != null ? candidate.getPhone() : (cell.getNumericCellValue()+""));
					}else if(cell.getCellTypeEnum() == CellType.STRING) {
						cell.setCellValue(candidate.getPhone() != null ? candidate.getPhone() : cell.getStringCellValue());
					}
					
					
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
			
			Sheet manager = candidatesWorkBook.getSheet(sheetNameManagers);
			Iterator<Row> dataMan = manager.iterator();
			dataMan.next();
			while(dataMan.hasNext()) {
				Row rowMan = dataMan.next();
				if (rowMan.getCell(0).getCellTypeEnum()==CellType.STRING) {
					String mani =rowMan.getCell(0).getStringCellValue();
					if (mani.compareToIgnoreCase(candidate.getManager())==0) {
						candidate.setManager(rowMan.getCell(1).getStringCellValue());
					}
				}else {
					int mani = (int) rowMan.getCell(0).getNumericCellValue();
					if(mani==Integer.parseInt(candidate.getManager())) {
						candidate.setManager(rowMan.getCell(1).getStringCellValue());
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
			classLoader = this.getClass().getClassLoader();
			File a = new File(path);// (classLoader.getResource(path).getFile());
			excelFile = new FileInputStream(a);
			CandidateExcel can;
			List<CandidateExcel> candidates = new ArrayList<>();
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			Sheet manager = candidatesWorkBook.getSheet(sheetNameManagers);
			Iterator<Row> dataMan = manager.iterator();
			data.next();
			dataMan.next();
			while (data.hasNext()) {
				Row currentRow = data.next();
				can = new CandidateExcel();
				if (currentRow.getCell(0).getCellTypeEnum()==CellType.STRING) {
					can.setCandidateId(Integer.parseInt(currentRow.getCell(0).getStringCellValue()));
				}else if(currentRow.getCell(0).getCellTypeEnum()== CellType.NUMERIC) {
					can.setCandidateId((int)currentRow.getCell(0).getNumericCellValue());
				}
				
				can.setName(currentRow.getCell(1).getStringCellValue());
				can.setEmail(currentRow.getCell(2).getStringCellValue());
				if (currentRow.getCell(3).getCellTypeEnum() == CellType.NUMERIC) {
					can.setPhone(currentRow.getCell(3).getNumericCellValue()+"");
				}else if(currentRow.getCell(3).getCellTypeEnum() == CellType.STRING) {
					can.setPhone(currentRow.getCell(3).getStringCellValue());
				}
				can.setProfile(currentRow.getCell(4).getStringCellValue());
				can.setYearsOfExperience((int) currentRow.getCell(5).getNumericCellValue());
				can.setEnglishLevel(currentRow.getCell(6).getStringCellValue());
				can.setStatus(currentRow.getCell(7).getStringCellValue());
				can.setCreationDate(new Date(currentRow.getCell(8).getStringCellValue()));
				can.setSkills(currentRow.getCell(9).getStringCellValue());
				can.setAging((int) currentRow.getCell(10).getNumericCellValue());


				while(dataMan.hasNext()) {
					Row rowMan = dataMan.next();
					if (rowMan.getCell(0).getCellTypeEnum()==CellType.STRING) {
						String mani =rowMan.getCell(0).getStringCellValue();
						if (mani.compareToIgnoreCase(can.getManager())==0) {
							can.setManager(rowMan.getCell(1).getStringCellValue());
						}
					}else {
						int mani = (int) rowMan.getCell(0).getNumericCellValue();
						if(mani== Integer.parseInt(currentRow.getCell(11).getStringCellValue() )) {
							can.setManager(rowMan.getCell(1).getStringCellValue());
						}
					}
				}
				can.setRelocated(currentRow.getCell(12).getBooleanCellValue());

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

	@DeleteMapping("/deleteCandidate/{id}")
	private ResponseEntity<Object> deleteCandidate(@PathVariable("id") int id) {
		try {
			classLoader = getClass().getClassLoader();
			excelFile = new FileInputStream(new File(path));
			candidatesWorkBook = new XSSFWorkbook(excelFile);
			Sheet worksheet = candidatesWorkBook.getSheet(sheetName);
			Iterator<Row> data = worksheet.iterator();
			ArrayList<Integer> rowsNumbers = new ArrayList<>();
			data.next();
			int i=0;
			while (data.hasNext()) {
				Row currentRow = data.next();
				int aux = (int)currentRow.getCell(1).getNumericCellValue();
				if (aux== id) {
					worksheet.removeRow(worksheet.getRow(i));
				}
				i++;
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
