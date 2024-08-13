package com.alfaCentauri;

import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class Ejecutable12 {
    /** Cuerpo del programa. **/
    public static void main(String[] args) {
        System.out.println("Ejemplo de ejecución de libreria Apache POI.\n");
        System.out.println("Leer multiples hojas de un libro de calculo.");
        String ruta = "data/pruebas.xls";
        try {
            FileInputStream file = new FileInputStream(new File(ruta));
            //Create Workbook instance holding reference to .xlsx file
            Workbook workbook = WorkbookFactory.create(file);
            //Get first/desired sheet from the workbook
            Sheet sheet = workbook.getSheetAt(0);
            int countSheets = workbook.getNumberOfSheets();
            System.out.println("Cantidad de hojas en el libro: " + countSheets);
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                    //Check the cell type and format accordingly
                    switch (evaluator.evaluateInCell(cell).getCellType())  {
                        case NUMERIC:
                            System.out.println(cell.getNumericCellValue() );
                            break;
                        case STRING:
                            System.out.println(cell.getStringCellValue() );
                            break;
                        case FORMULA:
                            String nombreHoja = workbook.getSheetName(0); 
                            System.out.println("El nombre de la hoja es " + nombreHoja + ". Contiene formula." );
                            break;
                    }
                }
                System.out.println("");
            }
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("Final de la prueba.");
    }

}
