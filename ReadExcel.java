package com.ysxy.tool;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class ReadExcel {
	public static void main(String[] args) {
		// 首先需要下载jxl.jar包，这个在搜索引擎一搜就有
		ReadExcel readExcel = new ReadExcel();
		// 注意jxl.jar包只能读取xls格式的文件，不能读取xlsx的文件，需要把xlsx另存为xls格式的文件
		File file = new File("C:\\Users\\Mystery\\Desktop\\历届美术留校作品汇清单.xls");
		readExcel.readExcel(file);
	}

	// 去读Excel的方法readExcel，该方法的入口参数为一个File对象
	public void readExcel(File file) {
		try {
			// 创建输入流，读取Excel
			InputStream is = new FileInputStream(file.getAbsolutePath());
			// jxl提供的Workbook类
			Workbook wb = Workbook.getWorkbook(is);
			// Excel的页签数量
			int sheet_size = wb.getNumberOfSheets();
			String cellinfo = "";
			for (int index = 0; index < sheet_size; index++) {
				// 每个页签创建一个Sheet对象
				Sheet sheet = wb.getSheet(index);
				// sheet.getRows()返回该页的总行数
				for (int i = 0; i < sheet.getRows(); i++) {
					// sheet.getColumns()返回该页的总列数
					for (int j = 0; j < sheet.getColumns(); j++) {
						if (j == sheet.getColumns() - 1) {
							cellinfo += "\"" + sheet.getCell(j, i).getContents() + "\" " + "\n";
						} else {
							cellinfo += "\"" + sheet.getCell(j, i).getContents() + "\", ";
						}
					}
					// 按行输出Excel的内容
					System.out.println(cellinfo);
				}
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
