package com.alfaCentauri;

public class ConvertirXLSToCSV {

    /**
     * @param args Type array String.
     **/
    public static void main(String []args){
        System.out.println("Ejemplo en Java para convertir un documento de Excel a CSV con la librería Apache POI.");
        String fileName = "C:/File raw.xlsx";
        File file = new File(fileName);
        FileInputStream fileInputStream;
        Workbook workbook = null;
        Sheet sheet;
        Iterator<Row> rowIterator;
        try {
            fileInputStream = new FileInputStream(file);
            String fileExtension = fileName.substring(fileName.indexOf("."));
            System.out.println(fileExtension);
            if(fileExtension.equals(".xls")){
                workbook  = new HSSFWorkbook(new POIFSFileSystem(fileInputStream));
            }
            else if(fileExtension.equals(".xlsx")){
                workbook  = new XSSFWorkbook(fileInputStream);
            }
            else {
                System.out.println("Wrong File Type");
            } 
            FormulaEvaluator evaluator workbook.getCreationHelper().createFormulaEvaluator();
            sheet = workbook.getSheetAt(0);
            rowIterator = sheet.iterator();
            while(rowIterator.hasNext()){
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()){
                Cell cell = cellIterator.next();
                //Check the cell type after evaluating formulae
               //If it is formula cell, it will be evaluated otherwise no change will happen
                switch (evaluator.evaluateInCell(cell).getCellType()){
                case Cell.CELL_TYPE_NUMERIC:
                System.out.print(cell.getNumericCellValue() + " ");
                break;
                case Cell.CELL_TYPE_STRING:
                System.out.print(cell.getStringCellValue() + " ");
                break;
                case Cell.CELL_TYPE_FORMULA:
                Not again
                break;
                case Cell.CELL_TYPE_BLANK:
                break;
            }
            }
             System.out.println("\n");
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e){
            e.printStackTrace();
        }
        System.out.println("Fin del programa.");
    }
}
