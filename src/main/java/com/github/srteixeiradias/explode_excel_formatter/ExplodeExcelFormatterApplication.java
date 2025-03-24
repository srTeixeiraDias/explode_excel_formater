package com.github.srteixeiradias.explode_excel_formatter;

import com.github.srteixeiradias.explode_excel_formatter.application.ExcelProcessorService;
import com.github.srteixeiradias.explode_excel_formatter.infraestructure.ExcelReader;
import com.github.srteixeiradias.explode_excel_formatter.infraestructure.ExcelWriter;

import java.io.File;
import java.util.Scanner;

public class ExplodeExcelFormatterApplication {

	public static void main(String[] args) {
		Scanner scanner = new Scanner(System.in);

		System.out.println("Digite o caminho do arquivo Excel de entrada:");
		String inputPath = scanner.nextLine();
		File inputFile = new File(inputPath);

		if (!inputFile.exists() || !inputFile.isFile()) {
			System.err.println("Arquivo de entrada não encontrado. Verifique o caminho e tente novamente.");
			return;
		}

		System.out.println("Digite o caminho onde deseja salvar o arquivo Excel processado:");
		String outputPath = scanner.nextLine();
		File outputFile = new File(outputPath);


		ExcelReader excelReader = new ExcelReader();
		ExcelWriter excelWriter = new ExcelWriter();
		ExcelProcessorService excelProcessorService = new ExcelProcessorService(excelReader, excelWriter);

		try {
			excelProcessorService.process(inputFile, outputFile);
			System.out.println("Processamento concluído com sucesso! Arquivo salvo em: " + outputPath);
		} catch (Exception e) {
			System.err.println("Erro ao processar o arquivo: " + e.getMessage());
		}
	}
}
