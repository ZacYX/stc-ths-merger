package ca.zac.ths;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.*;

public class StockListToSheet {

  public static void write(
      List<Stock> stockList,
      String date,
      Workbook workbook,
      String sheetName,
      String[] displayItems) {

    Sheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) {
      sheet = workbook.createSheet(sheetName);
      System.out.println("找不到Sheet: " + sheetName + ", 创建一个！");
    }

    prepareSheet(sheet, date, stockList.size());

    insertData(sheet, stockList, displayItems);
  }

  private static void prepareSheet(Sheet sheet, String date, Integer stockListSize) {
    // Blank sheet
    if (sheet.getLastRowNum() == -1) {
      Row newRow = sheet.createRow(0);
      newRow.createCell(0).setCellValue("类别");
      newRow.createCell(0).setCellValue("当日统计");
    }

    // Insert a blank column after the first column to the dataSheet, adding 3 to
    // solve outofbounds exception
    sheet.shiftColumns(2, sheet.getRow(0).getLastCellNum() + 3, 1);
    Cell newCell = sheet.getRow(0).createCell(2);
    newCell.setCellValue(date + " " + stockListSize);
    CellStyle cellStyle = getStyle(sheet.getWorkbook(), "CENTER", "CENTER");
    newCell.setCellStyle(cellStyle);
  }

  private static void insertData(Sheet sheet, List<Stock> stockList, String[] displayItems) {

    // Loop excel and set the second column cells 0
    System.out.println("Rows in excel: " + sheet.getLastRowNum());
    for (int j = 1; j <= sheet.getLastRowNum(); j++) {
      Row currentRow = sheet.getRow(j);
      if (currentRow == null) {
        continue;
      }
      if (currentRow.getCell(1) == null) {
        continue;
      }
      currentRow.getCell(1).setCellValue(0);
    }

    // Loop stockList
    System.out.println("Stock count: " + stockList.size());
    for (int i = 0; i < stockList.size(); i++) {
      System.out.println("Processing stock: " + stockList.get(i).getData().get(6) + " (" + i + ")");
      // Loop reasons
      for (int r = 0; r < stockList.get(i).getReasons().size(); r++) {
        // System.out.println("Processing stock: " + i + ", reason: " +
        // stockList.get(i).getReasons().get(r));
        Boolean isNewCategory = true;
        // Loop excel sheet
        for (int j = 0; j <= sheet.getLastRowNum() + 1; j++) {
          // System.out.println("Checking row: " + j + ", category: " +
          // stockList.get(i).getReasons().get(r));
          // This is a blank row, create a new one
          if (isNewCategory && j == sheet.getLastRowNum() + 1) {
            Row newRow = sheet.createRow(j);
            // Write category name
            newRow.createCell(0).setCellValue(stockList.get(i).getReasons().get(r));
            // Write formated stock info
            Cell newCell = newRow.createCell(2);
            newCell.setCellValue(formatDisplay("", stockList.get(i), displayItems));
            CellStyle cellStyle = getStyle(sheet.getWorkbook(), "LEFT", "TOP");
            newCell.setCellStyle(cellStyle);
            // Wirte the second row with 1
            newRow.createCell(1).setCellValue(1);
            break;
          }

          // compare existing category
          Row currentRow = sheet.getRow(j);
          // blank row
          if (currentRow == null) {
            continue;
          }
          Cell cellCategory = currentRow.getCell(0);
          if (cellCategory == null
              || !stockList.get(i).getReasons().get(r).contains(cellCategory.getStringCellValue().trim())) {
            continue;
          }
          // Found existing category
          Cell cellContent = currentRow.getCell(2);
          if (cellContent == null) {
            cellContent = currentRow.createCell(2);
          }
          String content = formatDisplay(cellContent.getStringCellValue(), stockList.get(i), displayItems);
          cellContent.setCellValue(content);
          CellStyle cellStyle = getStyle(sheet.getWorkbook(), "LEFT", "TOP");
          cellContent.setCellStyle(cellStyle);
          // update number of second column
          currentRow.getCell(1).setCellValue(currentRow.getCell(1).getNumericCellValue() + 1);
          isNewCategory = false;

        }

      }
    }

  }

  private static String formatDisplay(String originContent, Stock stock, String[] displayItems) {
    // 连续涨停天数+名称+首次涨停时间+涨停开板次数+最终涨停时间+涨停成交额+总金额+封单额+换手+流通市值
    SimpleDateFormat sdf = new SimpleDateFormat("hh:mm:ss");
    String outputString = "";
    // Loop displayItems
    for (int i = 0; i < displayItems.length; i++) {
      // Loop items in stock
      for (int j = 0; j < Stock.headers.size(); j++) {
        if (removeBracketContent(Stock.headers.get(j)).equals(displayItems[i])) {
          if (displayItems[i].equals("连续涨停天数")) {
            int days = (int) Double.parseDouble(stock.getData().get(j).toString());
            outputString += (days > 1) ? days : "  ";
          } else if (displayItems[i].equals("名称")) {
            // Avoid multiple stock in on cell
            // String stockName = stock.getData().get(j).toString().trim();
            String stockName = stock.getData().get(j).toString().replaceAll("[\\s\\u3000]+", "").trim();
            if (originContent.contains(stockName)) {
              return originContent;
            }
            outputString += stockName;
            outputString += (stockName.length() < 4) ? "      " : "  ";
          } else if (displayItems[i].equals("首次涨停时间") || displayItems[i].equals("最终涨停时间")) {
            Object dateObj = stock.getData().get(j);
            if (dateObj instanceof Date) {
              outputString += sdf.format((Date) dateObj) + "  ";
            } else {
              outputString += dateObj.toString() + "  ";
            }
          } else if (displayItems[i].equals("涨停开板次数")) {
            outputString += "K" + stock.getData().get(j).toString().split("\\.")[0] + "  ";
          } else if (displayItems[i].equals("涨停成交额")) {
            String amount = parseAmountOn10Per(stock.getData().get(j).toString()).toString();
            if (amount.length() < 4) {
              amount = amount + "  "; // 补齐长度
            }
            outputString += amount + "/";
          } else if (displayItems[i].equals("总金额")) {
            String amount = toYi((Double) stock.getData().get(j)).toString();
            if (amount.length() < 4) {
              amount = "  " + amount + "  "; // 补齐长度
            } else if (amount.length() < 5) {
              amount = "  " + amount; // 补齐长度
            }
            outputString += amount + "  ";
          } else if (displayItems[i].equals("封单额")) {
            Object data = stock.getData().get(j);
            if (data instanceof Number) {
              data = toYi((Double) data);
            }
            String dataString = data.toString();
            if (dataString.length() < 4) {
              data = data + "  "; // 补齐长度
            }
            outputString += "封" + data + "  ";
          } else if (displayItems[i].equals("流通市值")) {
            Object dataLiu = stock.getData().get(j);
            if (dataLiu instanceof Number) {
              dataLiu = toYi((Double) dataLiu);
            }
            outputString += "流" + dataLiu + "  ";
          } else if (displayItems[i].equals("换手")) {
            Object dataHuan = stock.getData().get(j);
            if (dataHuan instanceof Number) {
              dataHuan = Math.round((Double) dataHuan * 100.0); // 保留两位小数
            }
            String dataHuanString = dataHuan.toString();
            if (dataHuanString.length() < 2) {
              dataHuan = "  " + dataHuan; // 补齐长度
            }
            outputString += dataHuan + "%  ";
          } else {
            outputString += displayItems[i].substring(0, 1) + stock.getData().get(j) + "  ";
          }
          break;
        }
      }
    }
    return originContent + outputString + "\n";

  }

  private static Double toYi(Double value) {
    if (value == null) {
      return 0.0;
    }
    return Math.round(value / 100000000 * 100.0) / 100.0;
  }

  private static Double parseAmountOn10Per(String s) {
    double multiplieer = 1.0;
    String input = s;
    if (input.endsWith("万")) {
      input = input.replace("万", "");
      multiplieer = 10000;
    } else if (input.endsWith("亿")) {
      input = input.replace("亿", "");
      multiplieer = 100000000;
    }
    try {
      Double value = Double.parseDouble(input) * multiplieer;
      return Math.round(value / 100000000 * 100.0) / 100.0;
    } catch (NumberFormatException e) {
      System.out.println(s + " dose not contain a number!");
    }
    return 0.00;
  }

  private static CellStyle getStyle(Workbook wokbook, String horizontal, String vertical) {
    CellStyle cellStyle = wokbook.createCellStyle();
    cellStyle.setAlignment(HorizontalAlignment.valueOf(horizontal));
    cellStyle.setVerticalAlignment(VerticalAlignment.valueOf(vertical));
    cellStyle.setWrapText(true);
    return cellStyle;
  }

  private static String removeBracketContent(String text) {
    return text.replaceAll("\\[.*?\\]", "");
  }
}
