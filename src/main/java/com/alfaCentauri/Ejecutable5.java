package com.alfaCentauri;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.*;
import java.util.logging.Level;
import java.util.logging.Logger;

public class Ejecutable5 {
    public static void echoAsCSV(Sheet sheet) {
        Row row = null;
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            row = sheet.getRow(i);
            for (int j = 0; j < row.getLastCellNum(); j++) {
                System.out.print("\"" + row.getCell(j) + "\";");
            }
            System.out.println();
        }
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        String ruta = "data/pruebas.xls";
        File archivo = new File("input.xlsx");
        InputStream inputStream = null;
        InputStream result = null;
        try {
            inputStream = new FileInputStream(ruta);
            Workbook wb = WorkbookFactory.create(inputStream);
            TransformadorXLSToCSV transformadorXLSToCSV = new TransformadorXLSToCSV(wb);
            result = transformadorXLSToCSV.convertxlstoCSV_NotNull();
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                System.out.println(wb.getSheetAt(i).getSheetName());
                echoAsCSV(wb.getSheetAt(i));
            }
            if (inputStream != null ) {
                System.out.println("No Nulo");
            }
            else {
                System.out.println("Error: Nulo");
            }
            FileOutputStream out = new FileOutputStream(new File("data/output/procesado5.csv") );
            out.write(result.readAllBytes());
            out.close();
        } catch (IOException | InvalidFormatException ex) {
            Logger.getLogger(Ejecutable.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                inputStream.close();
            } catch (IOException ex) {
                Logger.getLogger(Ejecutable.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }

}
