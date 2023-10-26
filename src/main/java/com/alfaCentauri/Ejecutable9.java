package com.alfaCentauri;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

public class Ejecutable9 {

    /** Cuerpo del programa. **/
    public static void main(String[] args) {
        System.out.println("Ejemplo de ejecución de libreria Apache POI.\n");
        String ruta = "data/pruebas9.xls";
        //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Data");
        basedOnValue(sheet);
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(ruta));
            workbook.write(out);
            out.close();
            System.out.println("Se guardó correctamente en el disco.");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    static void basedOnValue(Sheet sheet)
    {
        //Creating some random values
        sheet.createRow(0).createCell(0).setCellValue(84);
        sheet.createRow(1).createCell(0).setCellValue(74);
        sheet.createRow(2).createCell(0).setCellValue(50);
        sheet.createRow(3).createCell(0).setCellValue(51);
        sheet.createRow(4).createCell(0).setCellValue(49);
        sheet.createRow(5).createCell(0).setCellValue(41);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        //Condition 1: Cell Value Is   greater than  70   (Blue Fill)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GT, "70");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.BLUE.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        //Condition 2: Cell Value Is  less than      50   (Green Fill)
        ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.LT, "50");
        PatternFormatting fill2 = rule2.createPatternFormatting();
        fill2.setFillBackgroundColor(IndexedColors.GREEN.index);
        fill2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A1:A6")
        };

        sheetCF.addConditionalFormatting(regions, rule1, rule2);
    }
}
