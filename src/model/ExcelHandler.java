package model;


import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatter;
import java.util.Date;

/**
 * Created by chenhui on 7/29/16.
 * This class is used for handling the Excel
 */
public class ExcelHandler {
    public final static String path = "/home/chenhui/myworkplace/xixiExeler/testFiles/";

    //parse the excel file and grep the needed entries and create the new excel file
    public static void parseFile() throws IOException {
        //read the excel file
        FileInputStream fileInputStream = new FileInputStream(path + "591_591_2.xls");
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
        HSSFSheet worksheet = workbook.getSheet("Sheet1");
        int lastRow = worksheet.getLastRowNum();
        HSSFRow row1 = worksheet.getRow(lastRow);
        HSSFCell dateCell = row1.getCell(1);
        String submittedDate = dateCell.getStringCellValue();
        System.out.println(submittedDate);

        //create the excel file
        Workbook wb = new HSSFWorkbook();  // or new XSSFWorkbook();
        Sheet sheet1 = wb.createSheet("Sheet1");
        Row row = sheet1.createRow(0);
        row.createCell(0).setCellValue("a");
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(true);

        FileOutputStream fileOut = new FileOutputStream(path + "workbook.xls");
        wb.write(fileOut);
        fileOut.close();
    }


}
