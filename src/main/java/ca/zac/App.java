package ca.zac;

import ca.zac.ths.ExcelFileInfo;
import ca.zac.ths.SheetToStockList;
import ca.zac.ths.StockListToSheet;
import ca.zac.ths.WorkbookWriter;

import java.io.FileInputStream;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.ss.usermodel.*;

import ca.zac.ths.ExcelFileCollector;

public class App {
    public static void main(String[] args) {

        System.out.println("Hello THS Merger!");
        if (args.length != 4) {
            System.out.println("Args must be 4 ");
            return;
        }
        String updaterPath = args[0];
        String marketInfoPath = args[1];
        String outputPath = args[2];
        String[] displayItems = args[3].split("\\+");

        System.out.println("Updater Path: " + updaterPath);
        System.out.println("Market Info Path: " + marketInfoPath);
        System.out.println("Output Path: " + outputPath);
        System.out.println("Display Items: " + String.join(", ", displayItems));

        // step 1: get all updater files
        List<ExcelFileInfo> updaters = ExcelFileCollector.getFiles(updaterPath);
        if (updaters.isEmpty()) {
            System.out.println("No updater files found in the specified path.");
            return;
        }
        System.out.println("Found " + updaters.size() + " updater files.");
        // step 2: get all marketinfo files
        List<ExcelFileInfo> marketInfos = ExcelFileCollector.getFiles(marketInfoPath);
        if (marketInfos.isEmpty()) {
            System.out.println("No market info files found in the specified path.");
            return;
        }
        System.out.println("Found " + marketInfos.size() + " market info files.");
        if (marketInfos.size() > 1) {
            System.out.println("Warning: More than one market info file found. Using the first one.");
        }

        try (
                // step 2: open marketinfo file and workbook
                FileInputStream fisMarketInfo = new FileInputStream(marketInfos.get(0).getFilePath());
                Workbook marketInfoWorkbook = new XSSFWorkbook(fisMarketInfo);) {

            // step 3: write update info in each file to marketinfo workbook
            for (ExcelFileInfo updater : updaters) {
                StockListToSheet.write(
                        SheetToStockList.get(updater.getFilePath(), 0),
                        updater.getDatePart(),
                        marketInfoWorkbook,
                        "概念全因",
                        displayItems);

            }

            // step 4: write final result to file
            WorkbookWriter.write(marketInfoWorkbook, outputPath);
            fisMarketInfo.close();

        } catch (Exception e) {
            System.err.println("Write final result failed: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
