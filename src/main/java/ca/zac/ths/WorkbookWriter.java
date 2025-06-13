package ca.zac.ths;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;

public class WorkbookWriter {

  /**
   * 将传入的 workbook 内容写入新的 Excel 文件
   *
   * @param workbook   来源 workbook
   * @param targetPath 原始文件路径，如 C:\data\marketinfo.xlsx
   */
  public static void write(Workbook workbook, String targetPath) {
    // 1. 获取当前日期
    String today = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss").format(new Date());

    // 2. 构造新的文件名
    String originalFileName = Paths.get(targetPath).getFileName().toString(); // marketinfo.xlsx
    String baseName = originalFileName.contains(".")
        ? originalFileName.substring(0, originalFileName.lastIndexOf('.'))
        : originalFileName;
    String newFileName = baseName + "--" + today + ".xlsx";

    // 3. 获取目录路径并构建完整路径
    String dirPath = Paths.get(targetPath).getParent().toString();
    String outputPath = Paths.get(dirPath, newFileName).toString();

    try (
        FileOutputStream fos = new FileOutputStream(outputPath)) {
      workbook.write(fos);
    } catch (Exception e) {
      System.err.println("Write excel file failed: " + e.getMessage());
    }
  }

}
