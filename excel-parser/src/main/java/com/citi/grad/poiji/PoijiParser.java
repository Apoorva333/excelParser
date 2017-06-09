package com.citi.grad.poiji;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.List;

import com.citi.grad.pojo.Column;
import com.poiji.internal.Poiji;

public class PoijiParser {

	public static void main(String args[]) throws FileNotFoundException{
		List<Column> columns = Poiji.fromExcel(new File("src/main/resources/WriteSheet.xlsx"), Column.class);
		System.out.println("columns size "+columns.size());
		//columns.get(1).fromColumns.getClass()
		
		
		System.out.println(columns.get(1).author);
	}
}
