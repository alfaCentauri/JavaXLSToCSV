package com.alfaCentauri;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.DateUtil;
import java.nio.charset.StandardCharsets;
import java.util.Iterator;

import static org.apache.poi.ss.usermodel.CellType.FORMULA;

public class TransformadorXLSToCSV {

    private FileInputStream file;

    private String urlPath;

    private Workbook workbook;

    private Sheet sheet;

    private Row row;

    private Cell cell;

    /** Construct **/
    public TransformadorXLSToCSV() {
        workbook = null;
        sheet = null;
        row = null;
        urlPath = "pruebas.xlsx";
    }

    public TransformadorXLSToCSV(Workbook wb) {
        workbook = wb;
        sheet = workbook.getSheetAt(0);
    }
    /**
     *
     * @param inputStream
     * @return Regresa un inputStream con el contenido del archivo CSV.
     * @throws IOException
     * @throws InvalidFormatException
     **/
    public InputStream convertxlstoCSV(InputStream inputStream) throws IOException, InvalidFormatException {
        workbook = WorkbookFactory.create(inputStream);
        return  csvConverter(workbook.getSheetAt(0));
    }

    /**
     *
     * @param inputStream
     * @return Regresa un inputStream con el contenido del archivo CSV.
     * @throws IOException
     * @throws InvalidFormatException
     **/
    public InputStream convertxlstoCSV_NotNull(InputStream inputStream) throws IOException, InvalidFormatException {
        workbook = WorkbookFactory.create(inputStream);
        sheet = workbook.getSheetAt(0);
        return  csvConverter_notNull();
    }

    /**
     * @param sheet Type Sheet.
     * @return Return a inputStream.
     **/
    protected InputStream csvConverter(Sheet sheet) {
        String str = new String();
        for (int i = 0; i < sheet.getLastRowNum()+1; i++) {
            row = sheet.getRow(i);
            String rowString = new String();
            int maxColumna = 6;
            for (int j = 0; j < maxColumna && row != null; j++) {
                if(row.getCell(j)==null) {
                    rowString = rowString + Utility.BLANK_SPACE + Utility.COMMA;
                }
                else {
                    rowString = rowString + row.getCell(j)+ Utility.COMMA;
                }
            }
            str = str + rowString.substring(0,rowString.length()-1)+ Utility.NEXT_LINE_OPERATOR;
        }
        System.out.println(str);
        return new ByteArrayInputStream(str.getBytes(StandardCharsets.UTF_8));
    }

    /**
     * Get max columns.
     * @param sheet Type Sheet.
     * @return Return a integer with number of columns.
     **/
    protected int getLastNumberColumn(Sheet sheet) {
        int count = 0;
        Row currentRow = sheet.getRow(0);
        if ( currentRow != null ) {
            Iterator iteratorCell = currentRow.cellIterator();
            while (iteratorCell.hasNext()) {
                Cell valueCell = currentRow.getCell(count);
                if (valueCell != null && !valueCell.getStringCellValue().isBlank())
                    count++;
                else
                    break;
            }
        }
        return count;
    }

    /**
     * @return Return a inputStream.
     **/
    protected InputStream csvConverter_notNull() {
        String str = new String();
        if ( sheet != null ) {
            for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
                row = sheet.getRow(i);
                int maxColumnNumbers = this.getColumnCount(i);
                String rowString = new String();
                for (int j = 0; j < maxColumnNumbers && row != null; j++) {
                    rowString = getStringToRow(row, rowString, j);
                }
                if (rowString.length() > 0)
                    str = str + rowString.substring(0, rowString.length() - 1) + Utility.NEXT_LINE_OPERATOR;
                else
                    i = sheet.getLastRowNum()+1;
            }
        }
        System.out.println(str);
        return new ByteArrayInputStream(str.getBytes(StandardCharsets.UTF_8));
    }

    private String getStringToRow(Row row, String rowString, int j) {
        cell = row.getCell(0);
        if ( cell != null ) {
            if (row.getCell(j) == null) {//Refactorizar
                rowString = rowString + Utility.BLANK_SPACE + Utility.COMMA;
            } else {
                rowString = rowString + getDataFormatt(j) + Utility.COMMA;
            }
        }
        return rowString;
    }

    /**
     * @param j Type int.
     * @return Return a string with data of cell.
     **/
    protected String getDataFormatt(int j) {
        String result="";
        Cell currentCell = row.getCell(j);
        switch (currentCell.getCellType()) {
            case STRING:
                result = currentCell.getStringCellValue();
                break;
            case NUMERIC:
                //Estilos de la celda
                CellStyle cellStyle = currentCell.getCellStyle();
                var estiloCelda = cellStyle.getDataFormatString();
                if ( estiloCelda.equals("[$-40A]d\" de \"mmmm\" de \"yyyy\\ h:mm:ss")  ) {
                    var fecha = DateUtil.getJavaDate(currentCell.getNumericCellValue());
                    result = fecha.toString();
                } else {
                    DataFormatter formatoDatos = new DataFormatter();
                    result = formatoDatos.formatCellValue(currentCell);
                }
                break;
            case BLANK:
                System.out.println("Blank");
                break;
            case BOOLEAN:
                result = String.valueOf( currentCell.getBooleanCellValue() );
                System.out.println("Boolean");
                break;
            case FORMULA:
                var formula = currentCell.getCellFormula();
                System.out.println("Formula: " + formula );
                //
                try {
                    Double valueCell = currentCell.getNumericCellValue();
                    result = String.valueOf(valueCell);
                } catch (Exception error) {
                    System.err.println("Error la formula no puede ser resuelta.\n" + error.getMessage());
                }
                break;
            case ERROR:
                System.out.println("Error");
                break;
            default:
                System.out.println("Otros");
        }
        return result;
    }

    /**
     * Get column count.
     * @param idSheet Type int.
     * @return Return a integer more zero if exist.
     **/
    protected int getColumnCount( int idSheet ) {
        int maxColumnNumbers = 0;
        row = sheet.getRow(idSheet);
        if (row != null)
            maxColumnNumbers = row.getLastCellNum();

        return maxColumnNumbers;
    }

    /** Transformar el InputStream a una lista **/

    public int getColumnCount(String sheetName)
    {
        sheet = workbook.getSheet(sheetName);
        row = sheet.getRow(0);
        int colCount = row.getLastCellNum();
        return colCount;
    }

    public int getRowCount(String sheetName)
    {
        sheet = workbook.getSheet(sheetName);
        int rowCount = sheet.getLastRowNum()+1;
        return rowCount;
    }

    public void ExcelApiTest(String xlFilePath) throws Exception
    {
        this.urlPath = xlFilePath;
        file = new FileInputStream(xlFilePath);
        workbook = new XSSFWorkbook(file);
        file.close();
    }

    /**
     * Re-calculating all formulas in a Workbook
     *
     * @param fis Tyype FileInputStream.
     * @return Return a Workbook with values.
     * @throws IOException
     */
    protected Workbook reCalculatingAllFormulasInAWorkbook(FileInputStream fis) throws IOException {
        Workbook wb = new HSSFWorkbook(fis); //or new XSSFWorkbook("/somepath/test.xls")
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        for (Sheet sheet : wb) {
            for (Row r : sheet) {
                for (Cell c : r) {
                    if (c.getCellType() == FORMULA) {
                        evaluator.evaluateFormulaCell(c);
                    }
                }
            }
        }
        return workbook;
    }

}
