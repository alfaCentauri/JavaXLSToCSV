package com.alfaCentauri;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.util.logging.Level;
import java.util.logging.Logger;

import static org.junit.jupiter.api.Assertions.*;

class TransformadorXLSToCSVTest {

    private TransformadorXLSToCSV transformadorXLSToCSV;

    private String ruta;

    private InputStream inputStream;

    private InputStream outputStream;

    @BeforeEach
    void setUp() {
    }

    @Test
    void convertxlstoCSVWithNull() {
        transformadorXLSToCSV = new TransformadorXLSToCSV();
        ruta = null;
        inputStream = null;
        outputStream = null;
        Throwable exception = assertThrows(
                NullPointerException.class,
                () -> {
                    //Basic
                    try{
                        inputStream = new FileInputStream(ruta);
                        outputStream = transformadorXLSToCSV.convertxlstoCSV(inputStream);
                        assertNotNull(outputStream, "Error on convertxlstoCSVWithNull");
                    } catch( FileNotFoundException ex) {
                        Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
                    } catch( IOException ex) {
                        Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (InvalidFormatException e) {
                        throw new RuntimeException(e);
                    } finally {
                        try {
                            inputStream.close();
                        } catch (IOException ex) {
                            Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
                        }
                    }
                },
                "Error: In test convertxlstoCSVWithNull."
        );
        assertEquals("Cannot invoke \"java.io.InputStream.close()\" because \"this.inputStream\" is null", exception.getMessage());
    }

    @Test
    void convertxlstoCSVWithFileNotFound() {
        transformadorXLSToCSV = new TransformadorXLSToCSV();
        ruta = "data/prueba.xls";
        inputStream = null;
        outputStream = null;
        //Basic with file no found
        Throwable exception = assertThrows(
                NullPointerException.class,
                () -> {
                    try{
                        inputStream = new FileInputStream(ruta);
                        outputStream = transformadorXLSToCSV.convertxlstoCSV(inputStream);
                        assertNotNull(outputStream, "Error on convertxlstoCSVWithNull");
                    } catch( FileNotFoundException ex) {
                        System.out.println("Falla del metodo. Prueba de FileNotFoundException");
                        Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
                    } catch( IOException ex) {
                        Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (InvalidFormatException e) {
                        throw new RuntimeException(e);
                    } finally {
                        try {
                            inputStream.close();
                        } catch (IOException ex) {
                            Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
                        }
                    }
                },
                "Error: In test convertxlstoCSVWithNull."
        );
        assertEquals("Cannot invoke \"java.io.InputStream.close()\" because \"this.inputStream\" is null", exception.getMessage());
        System.out.println("Prueba OK");
    }
}
