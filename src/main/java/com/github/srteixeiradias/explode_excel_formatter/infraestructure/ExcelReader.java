package com.github.srteixeiradias.explode_excel_formatter.infraestructure;


import com.github.srteixeiradias.explode_excel_formatter.domain.ExcelRow;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;



public class ExcelReader {
    // essa classe vai ser responsável por ler a planilha e extrair os dados sem células mescladas.

    public List<ExcelRow> readExcel(File inputFile) throws IOException {
        List<ExcelRow> rows = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            // Processa as células mescladas antes de ler os valores
            processMergedCells(sheet);

            // Ler cabeçalho
            Row headerRow = sheet.getRow(0);
            int columnCount = headerRow.getPhysicalNumberOfCells();

            // Ler os dados da planilha
            for (Row row : sheet) {
                List<String> cellValues = new ArrayList<>();
                for (int col = 0; col < columnCount; col++) {
                    Cell cell = row.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    cellValues.add(getCellValue(cell));
                }
                rows.add(new ExcelRow(cellValues));
            }
        }

        return rows;
    }

    //Preenche todas as células mescladas com o mesmo valor da célula original.
    private void processMergedCells(Sheet sheet) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            int firstRow = mergedRegion.getFirstRow();
            int lastRow = mergedRegion.getLastRow();
            int firstCol = mergedRegion.getFirstColumn();
            int lastCol = mergedRegion.getLastColumn();

            // Pega o valor da célula original
            Row row = sheet.getRow(firstRow);
            Cell firstCell = row.getCell(firstCol);
            String mergedValue = getCellValue(firstCell);

            // Preenche todas as células do intervalo mesclado com o mesmo valor
            for (int r = firstRow; r <= lastRow; r++) {
                Row currentRow = sheet.getRow(r);
                if (currentRow == null) {
                    currentRow = sheet.createRow(r);
                }
                for (int c = firstCol; c <= lastCol; c++) {
                    Cell currentCell = currentRow.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    currentCell.setCellValue(mergedValue);
                }
            }
        }
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> "";
        };
    } // retorna o valor da celula como string
}
