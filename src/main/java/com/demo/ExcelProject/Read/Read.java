package com.demo.ExcelProject.Read;
//https://www.youtube.com/watch?v=xabbFBBn6T8
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class Read {
    public static void main(String[] args) throws IOException {
        FileInputStream file = new FileInputStream("C:\\Users\\User\\IdeaProjects\\ExcelRead\\readFile.xlsx");
        Workbook workbook = new XSSFWorkbook(file);

        Sheet sheet = workbook.getSheetAt(0);

        System.out.println("Страница " + sheet.getSheetName());

        DataFormatter dataFormatter = new DataFormatter();

        for (Row row : sheet) {
            for (Cell cell : row) {
                    String data = dataFormatter.formatCellValue(cell);

                System.out.print(data + "\t");
            }
            System.out.println();
        }
    }
}
