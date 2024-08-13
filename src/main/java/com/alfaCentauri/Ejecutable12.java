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
        System.out.println("Leer multiples hojas de un libro de cálculo.");
        String ruta = "data/pruebas.xlsx";
        try {
            FileInputStream file = new FileInputStream(new File(ruta));
            //Create Workbook instance holding reference to .xlsx file
            Workbook workbook = WorkbookFactory.create(file);
            //Get first/desired sheet from the workbook
            Sheet sheet = workbook.getSheetAt(0);
            int countSheets = workbook.getNumberOfSheets();
            String nombreHoja = workbook.getSheetName(0);
            int numberRow = 1;
            System.out.println("Cantidad de hojas en el libro: " + countSheets);
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                //Identificando la linea
                System.out.println("Hoja: " + nombreHoja);
                System.out.print("Fila #" + numberRow++);
                System.out.print(";");
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                    //Check the cell type and format accordingly
                    switch (evaluator.evaluateInCell(cell).getCellType())  {/* Se evaluan las formulas antes de seleccionar el case */
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + ";");
                            break;
                        case STRING:
                            System.out.print(cell.getStringCellValue() + ";");
                            break;
                        case BLANK:
                            System.out.print("Blank;");
                            break;
                        case BOOLEAN:
                            String result = String.valueOf( cell.getBooleanCellValue() );
                            System.out.print("Boolean: " + result + ";");
                            break;
                        case ERROR:
                            String messageError = String.valueOf( cell.getErrorCellValue() );
                            System.out.print("Error: " + messageError + ";");
                            break;
                        default:
                            System.out.print("El nombre de la hoja es " + nombreHoja + ". Contiene otros.;");
                            break;
                    }
                }
                System.out.println("");
            }
            file.close();
        } catch (Exception e) {
            System.err.println("Fallo en: " + e.getMessage());
            e.printStackTrace();
        }
        System.out.println("Final de la prueba.");
    }

}
