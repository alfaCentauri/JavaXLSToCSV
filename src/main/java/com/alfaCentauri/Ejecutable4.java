package com.alfaCentauri;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;

public class Ejecutable4 {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {

        String ruta = "data/pruebas.xls";
        InputStream inputStream = null;
        InputStream outputStream = null;
        try{
            inputStream = new FileInputStream(ruta);
            Workbook wb = WorkbookFactory.create(inputStream);
            TransformadorXLSToCSV transformadorXLSToCSV = new TransformadorXLSToCSV(wb);
            outputStream = transformadorXLSToCSV.convertxlstoCSV_NotNull();
            FileOutputStream out = new FileOutputStream(new File("data/output/procesado4.csv") );
            out.write(outputStream.readAllBytes());
            out.close();
        } catch(IOException ex) {
            Logger.getLogger(Ejecutable.class.getName()).log(Level.SEVERE, null, ex);
        } catch (InvalidFormatException ex) {
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
