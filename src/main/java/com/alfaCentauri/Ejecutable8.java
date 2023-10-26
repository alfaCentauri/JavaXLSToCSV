package com.alfaCentauri;

import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

public class Ejecutable8 {
    /** Cuerpo del programa. **/
    public static void main(String[] args) {
        System.out.println("Ejemplo de ejecución de libreria Apache POI.\n");
        String ruta = "data/pruebas.xls";
        try {
            FileInputStream file = new FileInputStream(new File(ruta));
            //Create Workbook instance holding reference to .xlsx file
            Workbook workbook = WorkbookFactory.create(file);
            //Get first/desired sheet from the workbook
            Sheet sheet = workbook.getSheetAt(0);
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
                            System.out.print(cell.getNumericCellValue() + "t");
                            break;
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "t");
                            break;
                        case FORMULA:
                            //Not again
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
