package com.github.srteixeiradias.explode_excel_formatter.application;

import com.github.srteixeiradias.explode_excel_formatter.domain.ExcelRow;
import com.github.srteixeiradias.explode_excel_formatter.infraestructure.ExcelReader;
import com.github.srteixeiradias.explode_excel_formatter.infraestructure.ExcelWriter;

import java.io.File;
import java.io.IOException;
import java.util.List;

public class ExcelProcessorService {
    //classe responsavel por processar e gerar planilha nova já formatada

    private final ExcelReader excelReader;
    private final ExcelWriter excelWriter;

    public ExcelProcessorService(ExcelReader excelReader, ExcelWriter excelWriter) {
        this.excelReader = excelReader;
        this.excelWriter = excelWriter;
    }

    public void process(File inputFile, File outputFile) throws IOException {
        // Lê os dados da planilha
        List<ExcelRow> rows = excelReader.readExcel(inputFile);

        // Escreve a nova planilha formatada
        excelWriter.writeExcel(outputFile, rows);
    }
}
