package ca.zac.ths;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class SheetToStockList {

  public static List<Stock> get(String filePath, int sheetIndex) {
    List<Stock> stockList = new ArrayList<>();

    // try (FileInputStream fis = new FileInputStream(filePath);
    // Workbook workbook = new XSSFWorkbook(fis)) {
    try (
        FileInputStream fis = new FileInputStream(filePath);
        Workbook workbook = (filePath.endsWith(".xls")) ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);) {
      Sheet sheet = workbook.getSheetAt(sheetIndex);
      if (sheet == null) {
        System.err.println("找不到Sheet: " + sheetIndex);
        return stockList; // 返回空列表
      }

      Iterator<Row> rowIterator = sheet.iterator();
      if (!rowIterator.hasNext()) {
        return stockList; // 空sheet
      }

      // 读取表头
      Row headerRow = rowIterator.next();
      List<String> headers = new ArrayList<>();
      for (Cell cell : headerRow) {
        headers.add(cell.getStringCellValue().trim());
      }
      Stock.headers = headers;
      System.out.println("Headers: " + Stock.headers);

      // 查找涨停原因列的索引
      // 如果没有涨停原因列，直接返回空列表
      int dailyLimitReasonHeaderIndex = Stock.getHeaderIndex("涨停原因类别");
      if (dailyLimitReasonHeaderIndex == -1) {
        System.err.println("未找到涨停原因列: ");
        return stockList; // 如果没有找到涨停原因列，结束处理
      }
      System.out.println("Daily Limit Reason Header Index: " + dailyLimitReasonHeaderIndex);

      int stockNameHeaderIndex = Stock.getHeaderIndex("名称");
      if (stockNameHeaderIndex == -1) {
        System.err.println("未找到股票名称列");
        return stockList; // 如果没有找到股票名称列，结束处理
      }
      System.out.println("Stock Name Header Index: " + stockNameHeaderIndex);

      int continualDaysHeaderIndex = Stock.getHeaderIndex("连续涨停天数");
      if (continualDaysHeaderIndex == -1) {
        System.err.println("未找到连续涨停天数列");
        return stockList; // 如果没有找到连续涨停天数列，结束处理
      }
      System.out.println("Continual Days Header Index: " + continualDaysHeaderIndex);

      int firstTimeToLimitHeaderIndex = Stock.getHeaderIndex("首次涨停时间");
      if (firstTimeToLimitHeaderIndex == -1) {
        System.err.println("未找到首次涨停时间列");
        return stockList; // 如果没有找到首次涨停时间列，结束处理
      }
      System.out.println("First Time To Limit Header Index: " + firstTimeToLimitHeaderIndex);

      int increaseRateHeaderIndex = Stock.getHeaderIndex("涨幅");
      if (increaseRateHeaderIndex == -1) {
        System.err.println("未找到涨幅列");
        return stockList; // 如果没有找到涨幅列，结束处理
      }
      System.out.println("Increase Rate Header Index: " + increaseRateHeaderIndex);
      System.out.println("开始读取数据行...");

      // 读取数据行
      while (rowIterator.hasNext()) {
        Row row = rowIterator.next();
        List<Object> data = new ArrayList<>();
        List<String> reasons = new ArrayList<>();
        if (row.getLastCellNum() < 1) {
          continue; // 跳过空行
        }
        Object dailyLimitReason = getCellValue(row.getCell(dailyLimitReasonHeaderIndex));
        if (dailyLimitReason == null
            || dailyLimitReason.toString().isEmpty()
            || dailyLimitReason.toString().equals("--")) {
          continue; // 跳过没有涨停原因的行
        }
        Object continualDays = getCellValue(row.getCell(continualDaysHeaderIndex));
        if (continualDays == null
            || continualDays.toString().isEmpty()
            || continualDays.toString().equals("--")
            || continualDays.toString().equals("0")) {
          continue; // 跳过没有连续涨停天数的行
        }
        Object increaseRate = getCellValue(row.getCell(increaseRateHeaderIndex));
        if (increaseRate == null
            || increaseRate.toString().isEmpty()
            || increaseRate.toString().equals("--")
            || Double.parseDouble(increaseRate.toString()) < 0.09) {
          continue; // 跳过没有涨幅的行
        }

        for (int i = 0; i < Stock.headers.size(); i++) {
          Cell cell = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
          Object cellValue = getCellValue(cell);
          data.add(cellValue);
          if (i == dailyLimitReasonHeaderIndex) {
            reasons.addAll(Optional.ofNullable(cellValue).map(String::valueOf)
                .map(v -> Arrays.asList(v.split("\\+"))).orElse(Collections.emptyList()));
          }
        }
        System.out.println(data.get(stockNameHeaderIndex) + " " + reasons);
        Stock stock = new Stock(data, reasons);
        stockList.add(stock);
      }

      StockSorter.sortByDataColumn(stockList, firstTimeToLimitHeaderIndex, true);

    } catch (IOException e) {
      System.err.println("读取Excel失败: " + e.getMessage());
      e.printStackTrace();
    }

    return stockList;
  }

  private static Object getCellValue(Cell cell) {
    if (cell == null)
      return null;

    switch (cell.getCellType()) {
      case STRING:
        return cell.getStringCellValue().trim();
      case NUMERIC:
        if (DateUtil.isCellDateFormatted(cell)) {
          return cell.getDateCellValue();
        } else {
          return cell.getNumericCellValue();
        }
      case BOOLEAN:
        return cell.getBooleanCellValue();
      case FORMULA:
        return cell.getCellFormula(); // 或 cell.getStringCellValue()
      case BLANK:
        return null;
      default:
        return null;
    }
  }

}
