package com.jbaysolutions.excelparser;

import lombok.extern.java.Log;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.*;
import java.net.URL;

/**
 * Created by IntelliJ IDEA.
 * User: Gus - gustavo.santos@jbaysolutions.com - http://gmsa.github.io/
 * Date: 04-03-2015
 * Time: 11:38
 *
 * A sample class to parse Netsuite Excel Spreadsheet XML files. Full details in http://blog.jbaysolutions.com/2015/03/04/parsing-excel-spreadsheet-xml
 *
 */
@Log
public class ExcelXmlConverterNs {

    /**
     *
     * Downloads an Excel Spreadsheet XML and converts it to OOXML file format
     *
     * @throws Exception
     */
    public static File getAndConvertFile(File file) throws Exception {

        String fileContent = IOUtils.toString(new FileInputStream(file));

        fileContent = fileContent.replaceAll("\r&#13;&#10;","");

        SAXParserFactory parserFactor = SAXParserFactory.newInstance();
        SAXParser parser = parserFactor.newSAXParser();
        SAXHandler handler = new SAXHandler();

        ByteArrayInputStream bis = new ByteArrayInputStream(fileContent.getBytes());

        parser.parse(bis, handler);

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();

        //Converts all rows to POI rows
        int rowCount = 0;
        for (XmlRow subsRow : handler.xmlRowList) {
            Row row = sheet.createRow(rowCount);
            int cellCount = 0;
            for (String cellValue : subsRow.cellList) {
                Cell cell = row.createCell(cellCount);
                cell.setCellValue(cellValue);
                cellCount++;
            }
            rowCount++;
        }

        String fileOutPath = file.getAbsolutePath().replace(".xls",".xlsx");
        FileOutputStream fout = new FileOutputStream(fileOutPath);
        workbook.write(fout);
        workbook.close();
        fout.close();

        if (file.exists()) {
            System.out.println("delete file-> " + file.getAbsolutePath());
            if (!file.delete()) {
                System.out.println("file '" + file.getAbsolutePath() + "' was not deleted!");
            }
        }
        //System.out.println(result);
        System.out.println("getAndConvertFile finished, processed " + rowCount + " lines!");
        return new File(fileOutPath);
    }

}

