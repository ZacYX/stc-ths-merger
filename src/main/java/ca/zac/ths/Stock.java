package ca.zac.ths;

import java.util.List;

public class Stock {
  public static final String DAILY_LIMIT_REASON_HEADER = "涨停原因类别";

  static List<String> headers;

  private List<Object> data;
  private List<String> reasons;

  public Stock(List<Object> data, List<String> reasons) {
    this.data = data;
    this.reasons = reasons;
  }

  public static int getHeaderIndex(String headerName) {
    if (headers == null || headers.isEmpty()) {
      return -1;
    }
    for (int i = 0; i < headers.size(); i++) {
      if (removeBracketContent(headers.get(i)).equals(headerName)) {
        return i;
      }
    }
    return -1;
  }

  public List<Object> getData() {
    return this.data;
  }

  public void setData(List<Object> data) {
    this.data = data;
  }

  public List<String> getReasons() {
    return this.reasons;
  }

  public void setReasons(List<String> reasons) {
    this.reasons = reasons;
  }

  private static String removeBracketContent(String text) {
    return text.replaceAll("\\[.*?\\]", "");
  }

}
