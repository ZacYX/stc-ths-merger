package ca.zac.ths;

import java.util.*;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class StockSorter {

  public static void sortByDataColumn(List<Stock> stocks, int index, boolean ascending) {
    stocks.sort((s1, s2) -> {
      Comparable c1 = toComparable(s1.getData().get(index));
      Comparable c2 = toComparable(s2.getData().get(index));

      if (!c1.getClass().equals(c2.getClass())) {
        // throw new RuntimeException("不同类型不能比较: " + c1.getClass() + " vs " +
        // c2.getClass());
        System.out.println("不同类型不能比较: " + c1.getClass() + " vs " + c2.getClass());
        return 0; // 不同类型直接返回0，表示不排序
      }

      if (c1 == null && c2 == null)
        return 0;
      if (c1 == null)
        return ascending ? -1 : 1;
      if (c2 == null)
        return ascending ? 1 : -1;

      return ascending ? c1.compareTo(c2) : c2.compareTo(c1);
    });
  }

  private static Comparable toComparable(Object o) {
    if (o == null)
      return null;

    // 尝试将 String 或 Date 转换成 Date
    if (o instanceof Date) {
      return (Date) o;
    }

    if (o instanceof String) {
      String str = (String) o;

      // 尝试解析为日期
      String[] formats = {
          "yyyy-MM-dd HH:mm:ss",
          "yyyy-MM-dd",
          "MM/dd/yyyy",
          "yyyyMMdd"
      };
      for (String format : formats) {
        try {
          return new SimpleDateFormat(format).parse(str);
        } catch (ParseException ignored) {
        }
      }

      // 如果不是日期，再尝试数字
      try {
        return Double.parseDouble(str);
      } catch (NumberFormatException ignored) {
      }

      // 最后保底用字符串
      return str;
    }

    // 其他类型强制转为字符串
    return o.toString();
  }
}
