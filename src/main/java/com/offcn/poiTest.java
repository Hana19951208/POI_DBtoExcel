package com.offcn;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

/**
 * @author Hana
 * @create 2018/12/13-18:57
 */
public class poiTest {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        FileInputStream fis = new FileInputStream(new File("D:\\feiq\\Recv Files\\第五阶段最新课表-17天.xlsx"));
        Workbook workbook = WorkbookFactory.create(fis);

        Sheet sheet = workbook.getSheet("sheet1");
        int rows = sheet.getPhysicalNumberOfRows();

        for (int i = 0; i < rows; i++) {

            Row row = sheet.getRow(i);
            int cells = row.getPhysicalNumberOfCells();
            for (int j = 0; j < cells; j++) {
                Cell cell = row.getCell(j);
                if(cell.getCellTypeEnum()== CellType.STRING)
                    System.out.print(cell.getStringCellValue()+"\t");
                if(cell.getCellTypeEnum()== CellType.NUMERIC)
                    System.out.print(cell.getNumericCellValue()+"\t");
            }
            System.out.println();

        }


//        HSSFRow row = sheet.getRow(0);
//        HSSFCell cell = row.getCell(1);
//        System.out.println(cell.getStringCellValue());
//        fis.close();


    }
}
