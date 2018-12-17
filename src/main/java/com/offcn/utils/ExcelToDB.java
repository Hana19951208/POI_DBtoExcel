package com.offcn.utils;

import com.offcn.bean.Student;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;

public class ExcelToDB {
//
//	@Autowired
//	private StudentMapper mapper;

	@Test
	public void test() {
//		读取的文件：
		File ff = new File("E:\\11.xlsx");
		Workbook workbook=null;
		try {
			 workbook = WorkbookFactory.create(new FileInputStream(ff));
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		Sheet sheet = workbook.getSheet("sheet1");
		int rows = sheet.getPhysicalNumberOfRows();
		for (int i = 0; i < rows; i++) {
			Row row = sheet.getRow(i);
			int cells = row.getPhysicalNumberOfCells();
			Student stu = new Student();
			StringBuilder sb = new StringBuilder();
			
			for (int j = 0; j < cells; j++) {
				Cell cell = row.getCell(j);
				
				
				if(cell.getCellTypeEnum()==CellType.STRING) {
					sb.append(cell.getStringCellValue()+",");
				}
				if(cell.getCellTypeEnum()==CellType.NUMERIC){
					DecimalFormat df = new DecimalFormat("####");
					sb.append(df.format(cell.getNumericCellValue())+",");
				}
				
			}
			String[] str = sb.toString().split(",");
			
				
				stu.setId(Integer.valueOf(str[0]));
				stu.setName(str[1]);
				stu.setScore(Integer.valueOf(str[2]));
//		向数据库插入数据：
//				mapper.insert(stu);

			System.out.println(stu);
		}
	}
}
