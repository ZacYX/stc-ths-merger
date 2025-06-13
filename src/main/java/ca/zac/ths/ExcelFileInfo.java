package ca.zac.ths;

public class ExcelFileInfo {
  private String filePath;
  private String datePart;

  public ExcelFileInfo(String filePath, String datePart) {
    this.filePath = filePath;
    this.datePart = datePart;
  }

  public String getFilePath() {
    return filePath;
  }

  public String getDatePart() {
    return datePart;
  }

  @Override
  public String toString() {
    return "路径: " + filePath + ", 日期: " + datePart;
  }
}
