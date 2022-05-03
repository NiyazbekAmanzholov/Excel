package com.demo.ExcelProject.Write;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class Write {
    public static void main(String[] args) throws IOException {

//        Давайте создадим метод, который записывает список лиц на лист под названием “Лица".
//        Сначала мы создадим и оформим строку заголовка, содержащую ячейки “Имя” и “Возраст”.:
        XSSFWorkbook workbook = new XSSFWorkbook();

        Sheet sheet = workbook.createSheet("Persons");//Название страницы
        sheet.setColumnWidth(0, 6000);//Ширина первого столбца
        sheet.setColumnWidth(1, 4000);//Ширина второго

        Row header = sheet.createRow(0);//Строка

        Cell headerCell = header.createCell(0);//столбец
        headerCell.setCellValue("Name");

        headerCell = header.createCell(1);
        headerCell.setCellValue("Age");

        Row row = sheet.createRow(2);
        Cell cell = row.createCell(0);
        cell.setCellValue("John Smith");

        cell = row.createCell(1);
        cell.setCellValue(20);

//      Наконец, давайте запишем содержимое в “writeFile.xlsx ” файл в текущем каталоге и закройте книгу:
        FileOutputStream outputStream = new FileOutputStream("C:\\Users\\User\\IdeaProjects\\ExcelRead\\writeFile.xlsx");
        workbook.write(outputStream);

        System.out.println("файл writeFile.xlsx создан.");
        workbook.close();
    }
}