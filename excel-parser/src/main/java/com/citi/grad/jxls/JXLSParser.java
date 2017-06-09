package com.citi.grad.jxls;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.jxls.reader.ReaderBuilder;
import org.jxls.reader.XLSReadStatus;
import org.jxls.reader.XLSReader;

import com.citi.grad.pojo.Column;






public class JXLSParser {

	public static void main(String args[]) throws Exception {
		
		List<Column> perscolumnsons =parseExcelFileToBeans(new File("src/main/resources/WriteSheet.xlsx"),
                new File("src/main/resources/columnConfig.xml"));
		
		System.out.println(perscolumnsons.get(0).toString());
	}

	public  static<T> List<T> parseExcelFileToBeans(final File xlsFile, final File jxlsConfigFile)	throws Exception {

		final XLSReader xlsReader = ReaderBuilder.buildFromXML(jxlsConfigFile);
		final List<T> result = new ArrayList<>();
		final Map<String, Object> beans = new HashMap<>();
		beans.put("result", result);
		try (InputStream inputStream = new BufferedInputStream(new FileInputStream(xlsFile))) {
			xlsReader.read(inputStream, beans);
		}
		return result;
		
		/*InputStream inputXML = new BufferedInputStream(getClass().getResourceAsStream(jxlsConfigFile));
	    XLSReader mainReader = ReaderBuilder.buildFromXML( inputXML );
	    InputStream inputXLS = new BufferedInputStream(getClass().getResource(xlsFile));
	    Department department = new Department();
	    Department hrDepartment = new Department();
	    List departments = new ArrayList();
	    Map beans = new HashMap();
	    beans.put("department", department);
	    beans.put("hrDepartment", hrDepartment);
	    beans.put("departments", departments);
	    XLSReadStatus readStatus = mainReader.read( inputXLS, beans);*/
		
		
	}
}
