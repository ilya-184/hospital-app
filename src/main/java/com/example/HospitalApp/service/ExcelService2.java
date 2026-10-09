package com.example.HospitalApp.service;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
public class ExcelService2 {

    public ByteArrayInputStream processExcelFiles(List<MultipartFile> files) throws IOException {
        Workbook resultWorkbook = new XSSFWorkbook();

        List<Workbook> workbooks = new ArrayList<>();


        for (MultipartFile file : files) {
            workbooks.add(new XSSFWorkbook(file.getInputStream()));
        }


        int maxSheets = workbooks.stream().mapToInt(Workbook::getNumberOfSheets).max().orElse(0);
        for (int i = 0; i < maxSheets; i++) {
            Sheet resultSheet = resultWorkbook.createSheet("Sheet" + (i + 1));
            aggregateSheets(workbooks, resultSheet, i);
        }

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        resultWorkbook.write(outputStream);
        resultWorkbook.close();

        return new ByteArrayInputStream(outputStream.toByteArray());
    }


    private static void aggregateSheets(List<Workbook> workbooks, Sheet resultSheet, int sheetIndex) {
        List<FormulaEvaluator> evaluators = new ArrayList<>();
        for (Workbook workbook : workbooks) {
            evaluators.add(workbook.getCreationHelper().createFormulaEvaluator());
        }

        for (int i = 0; i < workbooks.size(); i++) {
            Workbook workbook = workbooks.get(i);
            FormulaEvaluator evaluator = evaluators.get(i);

            if (sheetIndex < workbook.getNumberOfSheets()) {

                Sheet sheet = workbook.getSheetAt(sheetIndex);

                for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                    Row row = sheet.getRow(rowNum);
                    Row resultRow = resultSheet.getRow(rowNum);
                    if (resultRow == null) {
                        resultRow = resultSheet.createRow(rowNum);
                    }
                    if (row != null) {
                        for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                            Cell cell = row.getCell(cellNum);
                            Cell resultCell = resultRow.getCell(cellNum);
                            if (resultCell == null) {
                                resultCell = resultRow.createCell(cellNum);
                            }
                            if (cell != null) {
                                if (cell.getCellType() == CellType.FORMULA) {
                                    evaluator.evaluateFormulaCell(cell);
                                }
                                aggregateCellValues(cell, resultCell, evaluator);
                            }
                        }
                    }
                }
            }
        }
    }

    private static void aggregateCellValues(Cell sourceCell, Cell resultCell, FormulaEvaluator evaluator) {
        CellValue cellValue = evaluator.evaluate(sourceCell);

        if (cellValue != null) {
            switch (cellValue.getCellType()) {
                case NUMERIC -> {
                    double newValue = cellValue.getNumberValue();
                    if (resultCell.getCellType() == CellType.NUMERIC) {
                        newValue += resultCell.getNumericCellValue();
                    }
                    resultCell.setCellValue(newValue);
                }
                case STRING -> resultCell.setCellValue(cellValue.getStringValue());
                case BOOLEAN -> resultCell.setCellValue(cellValue.getBooleanValue());
                default -> {
                }
            }
        }
    }



}