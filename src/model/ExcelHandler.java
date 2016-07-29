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
import java.util.Date;

/**
 * Created by chenhui on 7/29/16.
 * This class is used for handling the Excel
 */
public class ExcelHandler {

    //parse the excel file and grep the needed entries and create the new excel file
    public static void parseFile() throws IOException {
        FileInputStream fileInputStream = new FileInputStream("/home/chenhui/myworkplace/xixiExeler/testFiles/591_591_2.xls");
        HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
        HSSFSheet worksheet = workbook.getSheet("Sheet1");
        HSSFRow row1 = worksheet.getRow(0);
        HSSFCell cellA1 = row1.getCell(0);
        String a1Val = cellA1.getStringCellValue();
        HSSFCell cellB1 = row1.getCell(1);
        String b1Val = cellB1.getStringCellValue();

        System.out.println("A1: " + a1Val);
        System.out.println("B1: " + b1Val);
    }


}
