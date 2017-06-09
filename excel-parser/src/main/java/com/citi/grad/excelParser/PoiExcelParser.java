package com.citi.grad.excelParser;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.citi.grad.pojo.Column;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

public class PoiExcelParser {

	private static Map<Integer, String> indexMap = new HashMap<Integer, String>();
	private static int count = 0;
	public static void main(String[] args) {
		try {
			FileInputStream file = new FileInputStream(new File("src/main/resources/WriteSheet.xlsx"));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();

			indexColumnFields(rowIterator.next());
            List<Column> columnList = new ArrayList<Column>();
            
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				columnList.add(assignColumn(row));
			}
			System.out.println("Size of column list is "+columnList.size());
			
			System.out.println("Row 1 is "+columnList.toString());
			
			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@SuppressWarnings("deprecation")
	private static Column assignColumn(Row row) {
		Column column = new Column();
		for (int i = 0; i < count; i++) {
			Cell cell = row.getCell(i);
			if (cell.getCellType() != 3) {
				if ((cell.getCellType() == 0) && (!indexMap.get(i).equalsIgnoreCase( ColumnFieldConstants.VERSION)))
					cell.setCellType(1);
				switch (indexMap.get(i)) {
				case ColumnFieldConstants.ID:
					column.id = row.getCell(i).getStringCellValue();
					break;
				case ColumnFieldConstants.AUTHOR:
					column.author = row.getCell(i).getStringCellValue();
					break;
				case ColumnFieldConstants.FROMCOLUMNS:
					column.fromColumns = new HashSet<String>(Arrays.asList(cell.getStringCellValue().trim().split(",")));
					break;
				case ColumnFieldConstants.NAME:
					column.name = cell.getStringCellValue();
					break;
				case ColumnFieldConstants.STATUS:
					column.status = cell.getStringCellValue();
					break;
				case ColumnFieldConstants.TABLEID:
					column.tableId = cell.getStringCellValue();
					break;
				case ColumnFieldConstants.VERSION:
					column.version = (int) cell.getNumericCellValue();
					break;
				default:
					column.metadata.put(indexMap.get(i),cell.getStringCellValue());
					break;
				}
			}
		}
		return column;
	}

	@SuppressWarnings("deprecation")
	private static void indexColumnFields(Row row) {
		Iterator<Cell> cellIterator = row.cellIterator();
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			if (cell.getCellType() != 1)
				cell.setCellType(1);
			indexMap.put(count++, cell.getStringCellValue().toLowerCase());
		}
	}

}
