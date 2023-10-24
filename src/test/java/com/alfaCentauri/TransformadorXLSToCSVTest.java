package com.alfaCentauri;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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
                            System.out.println("Falla del metodo.\n" + ex.getMessage());
                            Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
                        }
                    }
                },
                "Error: In test convertxlstoCSVWithNull."
        );
        assertEquals("Cannot invoke \"java.io.InputStream.close()\" because \"this.inputStream\" is null", exception.getMessage());
        System.out.println("Prueba OK");
    }

    @Test
    void convertxlstoCSVWithFile_ShouldSuccess() {
        transformadorXLSToCSV = new TransformadorXLSToCSV();
        ruta = "data/pruebas.xls";
        inputStream = null;
        outputStream = null;
        //Basic with file exist
        try{
            inputStream = new FileInputStream(ruta);
            outputStream = transformadorXLSToCSV.convertxlstoCSV(inputStream);
            assertNotNull(outputStream, "Error on convertxlstoCSVWithNull");
        } catch( FileNotFoundException ex) {
            System.out.println("Falla del metodo. Prueba de FileNotFoundException\n" + ex.getMessage());
            Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch( IOException ex) {
            System.out.println("Falla del metodo. Prueba de IOException\n" + ex.getMessage());
            Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (InvalidFormatException ex) {
            System.out.println("Falla del metodo. Prueba de InvalidFormatException\n" + ex.getMessage());
        } finally {
            try {
                inputStream.close();
            } catch (IOException ex) {
                System.out.println("Falla del metodo.\n" + ex.getMessage());
                Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        //assertEquals("Cannot invoke \"java.io.InputStream.close()\" because \"this.inputStream\" is null", exception.getMessage());
        System.out.println("Prueba OK");
    }

    @Test
    void getLastNumberColumnWithFile_ShouldSuccess() {
        transformadorXLSToCSV = new TransformadorXLSToCSV();
        ruta = "data/pruebas.xls";
        inputStream = null;
        outputStream = null;
        //Basic with file exist
        try{
            inputStream = new FileInputStream(ruta);
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            int result = transformadorXLSToCSV.getLastNumberColumn(sheet);
            assertEquals(6, result);
        } catch( FileNotFoundException ex) {
            System.err.println("Falla del metodo. Prueba de FileNotFoundException");
            Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch( IOException ex) {
            System.err.println("Falla del metodo. Prueba de IOException");
            Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                inputStream.close();
            } catch (IOException ex) {
                Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        System.out.println("Prueba OK");
    }

    @Test
    void getColumnCount_ShouldSuccess() {
        transformadorXLSToCSV = new TransformadorXLSToCSV();
        ruta = "data/pruebas.xls";
        inputStream = null;
        outputStream = null;
        //Basic with file exist
        try{
            inputStream = new FileInputStream(ruta);
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            int result = transformadorXLSToCSV.getColumnCount(0);
            assertEquals(6, result);
        } catch( FileNotFoundException ex) {
            System.err.println("Falla del metodo. Prueba de FileNotFoundException");
            Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch( IOException ex) {
            System.err.println("Falla del metodo. Prueba de IOException");
            Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                inputStream.close();
            } catch (IOException ex) {
                Logger.getLogger(TransformadorXLSToCSVTest.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        System.out.println("Prueba OK");
    }
}
