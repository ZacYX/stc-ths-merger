/**
 * 获取目录下所有符合要求的文件路径
 */
package ca.zac.ths;

import java.io.File;
import java.io.FilenameFilter;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelFileCollector {
  public static List<ExcelFileInfo> getFiles(String fullPathWithPrefix) {
    // 1. 分离目录和前缀
    File fullPathFile = new File(fullPathWithPrefix);
    String parentDir = fullPathFile.getParent(); // 提取目录路径
    String prefix = fullPathFile.getName(); // 提取前缀（如 u）

    System.out.println("parentDir = " + parentDir);
    System.out.println("prefix = " + prefix);

    List<ExcelFileInfo> fileInfoList = new ArrayList<>();

    if (parentDir == null) {
      System.out.println("目录路径无效");
      return fileInfoList;
    }

    File dir = new File(parentDir);

    // 2. 构造正则：匹配前缀+4位数字+.xlsx
    // String regex = "^" + Pattern.quote(prefix) + "(\\d{4})\\.xlsx$"; //4位数字
    // String regex = "^" + Pattern.quote(prefix) + "([A-Za-z0-9]+)\\.xlsx$";
    // //字母或数字
    // String regex = "^" + Pattern.quote(prefix) + "([^/\\\\:*?\"<>|]+)\\.xlsx$";
    String regex = "(?i)^" + Pattern.quote(prefix) + "([^/\\\\:*?\"<>|]+)\\.(?:xls|xlsx)$";
    // // 字符，数字或字母
    final Pattern pattern = Pattern.compile(regex);

    // 3. 过滤匹配的文件
    File[] files = dir.listFiles(new FilenameFilter() {
      @Override
      public boolean accept(File dir, String name) {
        return pattern.matcher(name).matches();
      }
    });
    System.out.println("files.length = " + (files == null ? 0 : files.length));
    // 如果目录不存在或没有符合条件的文件，files 可能为 null

    // 4. 输出结果
    if (files != null) {
      for (File file : files) {
        Matcher matcher = pattern.matcher(file.getName());
        if (matcher.matches()) {
          String datePart = matcher.group(1); // 提取 u 后面的4位数字
          String fullPath = file.getAbsolutePath();
          ExcelFileInfo info = new ExcelFileInfo(fullPath, datePart);
          fileInfoList.add(info);

          System.out.println("文件名：" + fullPath);
          System.out.println("提取的日期部分：" + datePart);
        }
      }
    } else {
      System.out.println("目录不存在或没有符合条件的文件");
    }
    Collections.sort(fileInfoList, (a, b) -> a.getDatePart().compareTo(b.getDatePart()));
    System.out.println("找到 " + fileInfoList.size() + " 个符合条件的文件。");
    for (ExcelFileInfo info : fileInfoList) {
      System.out.println(info);
    }
    return fileInfoList;
  }
}
