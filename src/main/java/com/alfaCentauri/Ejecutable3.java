package com.alfaCentauri;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.*;
import java.util.logging.Level;
import java.util.logging.Logger;

public class Ejecutable3 {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        TransformadorXLSToCSV transformadorXLSToCSV = new TransformadorXLSToCSV();
        String ruta = "data/pruebas.xls";
        InputStream inputStream = null;
        InputStream outputStream = null;
        try{
            inputStream = new FileInputStream(ruta);
            outputStream = transformadorXLSToCSV.convertxlstoCSV(inputStream);
            FileOutputStream out = new FileOutputStream(new File("data/output/procesado3.csv") );
            out.write(outputStream.readAllBytes());
            out.close();
        } catch( FileNotFoundException ex) {
            Logger.getLogger(Ejecutable.class.getName()).log(Level.SEVERE, null, ex);
        } catch( IOException ex) {
            Logger.getLogger(Ejecutable.class.getName()).log(Level.SEVERE, null, ex);
        } catch (InvalidFormatException e) {
            throw new RuntimeException(e);
        } finally {
            try {
                inputStream.close();
            } catch (IOException ex) {
                Logger.getLogger(Ejecutable.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
}
