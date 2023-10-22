package com.alfaCentauri;

import org.apache.commons.csv.CSVFormat;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.extractor.XSSFExcelExtractor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.commons.csv.CSVPrinter;
import java.io.*;

public class Ejecutable2 {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        String ruta = "data/pruebas.xls";
        XSSFWorkbook input = null;
        try {
            input = new XSSFWorkbook(new File(ruta));
            CSVPrinter output = new CSVPrinter(new FileWriter("output.csv"), CSVFormat.DEFAULT);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        }

        String tsv = new XSSFExcelExtractor(input).getText();
        BufferedReader reader = new BufferedReader(new StringReader(tsv));
//        reader.lines().map(line -> line.split("\t").forEach(output::printRecord);
    }
}
