package ca.zac.ths;

import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;

public class StockCountToSheet {
  final static Integer SUM_COLUMN_INDEX = 0;
  final static Integer CATEGORY_NAME_COLUMN_INDEX = 1;
  final static Integer INSERTED_COLUMN_INDEX = 2;

  public static void write(
      List<Stock> stockList,
      String date,
      Workbook workbook,
      String sheetName) {

    workbook.setForceFormulaRecalculation(true);

    Sheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
      sheet = workbook.createSheet(sheetName);
      System.out.println("找不到Sheet: " + sheetName + ", 创建一个！");
    }

    prepareSheet(sheet, date, stockList.size());

    insertData(sheet, stockList);
  }

  private static void prepareSheet(Sheet sheet, String date, Integer stockListSize) {
    // Blank sheet
    CellStyle cellStyle = getStyle(sheet.getWorkbook(), "CENTER", "CENTER");

    if (sheet.getLastRowNum() == -1) {
      Row newRow = sheet.createRow(0);
      Cell newSumCell = newRow.createCell(SUM_COLUMN_INDEX);
      newSumCell.setCellValue("当日统计");
      newSumCell.setCellStyle(cellStyle);
      Cell newCatCell = newRow.createCell(CATEGORY_NAME_COLUMN_INDEX);
      newCatCell.setCellValue("类别");
      newCatCell.setCellStyle(cellStyle);
    }

    // Insert a blank column after the first column to the dataSheet, adding 3 to
    // solve outofbounds exception
    sheet.shiftColumns(2, sheet.getRow(0).getLastCellNum() + 3, 1);
    Cell newInsCell = sheet.getRow(0).createCell(2);
    newInsCell.setCellValue(date + " " + stockListSize);
    newInsCell.setCellStyle(cellStyle);
  }

  private static void insertData(Sheet sheet, List<Stock> stockList) {

    // Loop stockList
    System.out.println("Stock count: " + stockList.size());
    CellStyle cellStyle = getStyle(sheet.getWorkbook(), "CENTER", "CENTER");

    for (int i = 0; i < stockList.size(); i++) {
      System.out.println("Processing stock: " + stockList.get(i).getData().get(
          Stock.getHeaderIndex("名称")) + " (" + i + ")");
      // Loop reasons
      for (int r = 0; r < stockList.get(i).getReasons().size(); r++) {
        // System.out.println("Processing stock: " + i + ", reason: " +
        // stockList.get(i).getReasons().get(r));
        Boolean isNewCategory = true;
        // Loop excel sheet
        for (int j = 1; j <= sheet.getLastRowNum() + 1; j++) {
          // System.out.println("Checking row: " + j + ", category: " +
          // stockList.get(i).getReasons().get(r));
          // This is a blank row, create a new one
          if (isNewCategory && j == sheet.getLastRowNum() + 1) {
            Row newRow = sheet.createRow(j);
            // Write category name
            Cell newCatCell = newRow.createCell(CATEGORY_NAME_COLUMN_INDEX);
            newCatCell.setCellValue(stockList.get(i).getReasons().get(r));
            newCatCell.setCellStyle(cellStyle);
            // write fisrt column with as sum cell
            Cell newSumCell = newRow.createCell(SUM_COLUMN_INDEX);
            // newSumCell.setCellValue(0);
            newSumCell.setCellStyle(cellStyle);
            int currentRowNum = newRow.getRowNum() + 1;
            String colLetter = CellReference.convertNumToColString(newSumCell.getColumnIndex() + 1);
            String endColLetter = CellReference.convertNumToColString(newSumCell.getColumnIndex() + 20);
            newSumCell.setCellFormula("SUM(" + colLetter + currentRowNum + ":" +
                endColLetter + currentRowNum + ")");

            // Wirte the second row with 1
            Cell newInsCell = newRow.createCell(INSERTED_COLUMN_INDEX);
            newInsCell.setCellValue(1);
            newInsCell.setCellStyle(cellStyle);
            break;
          }

          // compare existing category
          Row currentRow = sheet.getRow(j);
          // blank row
          if (currentRow == null) {
            continue;
          }
          Cell cellCategory = currentRow.getCell(CATEGORY_NAME_COLUMN_INDEX);
          if (cellCategory == null) {
            continue;
          }
          // for custom category, split by "|"
          String[] categoryContent = cellCategory.getStringCellValue().trim().split("\\|");
          int k = 0;
          for (; k < categoryContent.length; k++) {
            // System.out.println("Checking category content: " + categoryContent[k].trim()
            // + ", reason: " +
            // stockList.get(i).getReasons().get(r));
            if (stockList.get(i).getReasons().get(r).contains(categoryContent[k].trim())) {
              break;
            }
          }
          if (k == categoryContent.length) {
            continue; // not found, continue to next row
          }
          // Found existing category
          // update number of second column
          // currentRow.getCell(1).setCellValue(currentRow.getCell(1).getNumericCellValue()
          // + 1);
          if (currentRow.getCell(INSERTED_COLUMN_INDEX) == null) {
            // currentRow.createCell(INSERTED_COLUMN_INDEX).setCellValue(0);
            Cell newInsCell = currentRow.createCell(INSERTED_COLUMN_INDEX);
            newInsCell.setCellStyle(cellStyle);
            newInsCell.setCellValue(1);
          } else if (currentRow.getCell(INSERTED_COLUMN_INDEX).getCellType() == CellType.NUMERIC) {
            currentRow.getCell(INSERTED_COLUMN_INDEX).setCellValue(currentRow.getCell(
                INSERTED_COLUMN_INDEX).getNumericCellValue() + 1);
          } else if (currentRow.getCell(INSERTED_COLUMN_INDEX).getCellType() == CellType.STRING) {
            // If the cell is a string, parse it to an integer and add 1
            currentRow.getCell(
                INSERTED_COLUMN_INDEX).setCellValue(
                    String.valueOf(Integer.parseInt(currentRow.getCell(
                        INSERTED_COLUMN_INDEX).getStringCellValue()) + 1));
          } else {
            // If the cell is not numeric or string, set it to 0
            // currentRow.getCell(INSERTED_COLUMN_INDEX).setCellValue(0);
          }
          isNewCategory = false;

        }

      }
    }

  }

  private static CellStyle getStyle(Workbook wokbook, String horizontal, String vertical) {
    CellStyle cellStyle = wokbook.createCellStyle();
    cellStyle.setAlignment(HorizontalAlignment.valueOf(horizontal));
    cellStyle.setVerticalAlignment(VerticalAlignment.valueOf(vertical));
    cellStyle.setWrapText(true);
    return cellStyle;
  }

}
