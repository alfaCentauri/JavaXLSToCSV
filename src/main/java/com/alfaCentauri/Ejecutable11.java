package com.alfaCentauri;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

public class Ejecutable11 {

    /** Cuerpo del programa. **/
    public static void main(String[] args) {
        System.out.println("Ejemplo de ejecución de libreria Apache POI.\n");
        String ruta = "data/pruebas10.xls";
        //Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Data");
        //basedOnValue(sheet);
        formatDuplicates(sheet);
        try{//Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(ruta));
            workbook.write(out);
            out.close();
            System.out.println("Se guardó correctamente en el disco.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static void basedOnValue(Sheet sheet) {
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

    /**
     * Da formato a los duplicados.
     * @param sheet
     */
    static void formatDuplicates(Sheet sheet) {
        sheet.createRow(0).createCell(0).setCellValue("Code");
        sheet.createRow(1).createCell(0).setCellValue(4);
        sheet.createRow(2).createCell(0).setCellValue(3);
        sheet.createRow(3).createCell(0).setCellValue(6);
        sheet.createRow(4).createCell(0).setCellValue(3);
        sheet.createRow(5).createCell(0).setCellValue(5);
        sheet.createRow(6).createCell(0).setCellValue(8);
        sheet.createRow(7).createCell(0).setCellValue(0);
        sheet.createRow(8).createCell(0).setCellValue(2);
        sheet.createRow(9).createCell(0).setCellValue(8);
        sheet.createRow(10).createCell(0).setCellValue(6);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =A2=A1   (White Font)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("COUNTIF($A$2:$A$11,A2)>1");
        FontFormatting font = rule1.createFontFormatting();
        font.setFontStyle(false, true);
        font.setFontColorIndex(IndexedColors.BLUE.index);
        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A2:A11")
        };
        sheetCF.addConditionalFormatting(regions, rule1);
        sheet.getRow(2).createCell(1).setCellValue("<== Duplicates numbers in the column are highlighted.  " +
                "Condition: Formula Is =COUNTIF($A$2:$A$11,A2)>1   (Blue Font)");
    }

    /**
     * Alternate color rows in different colors.
     * @param sheet Type Sheet.
     */
    static void shadeAlt(Sheet sheet) {
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =A2=A1   (White Font)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("MOD(ROW(),2)");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A1:Z100")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.createRow(0).createCell(1).setCellValue("Shade Alternating Rows");
        sheet.createRow(1).createCell(1).setCellValue("Condition: Formula Is  =MOD(ROW(),2)   (Light Green Fill)");
    }
}
