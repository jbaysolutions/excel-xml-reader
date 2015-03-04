package com.jbaysolutions.excelparser;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.IOUtils;

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
 * A sample class to parse Excel Spreadsheet XML files. Full details in http://blog.jbaysolutions.com/2015/03/04/parsing-excel-spreadsheet-xml
 *
 */
public class ExcelXmlReader {
    public static void main(String[] args) throws Exception {
        getAndParseFile();
    }


    /**
     *
     * Downloads an Excel Spreadsheet XML and reads the substance values (name, ADI and ARfD)
     *
     * @throws Exception
     */
    private static void getAndParseFile() throws Exception {
        System.out.println("getAndParseFile");
        File file = File.createTempFile("substances", "tmp");


        String excelFileUrl = "http://ec.europa.eu/sanco_pesticides/public/?event=activesubstance.exportList";
        URL url = new URL(excelFileUrl);
        System.out.println("downloading file from " + excelFileUrl + " ...");
        FileUtils.copyURLToFile(url, file);
        System.out.println("downloading finished, parsing...");

        removeLineFromFile(file.getAbsolutePath(), 1, 2);

        String fileContent = IOUtils.toString(new FileInputStream(file));
        fileContent = fileContent.replaceAll("&ECCO", "&#38;ECCO");
        fileContent = "<?xml version=\"1.0\"?>\n" +
                "<!DOCTYPE some_name [ \n" +
                "<!ENTITY nbsp \"&#160;\"> \n" +
                "<!ENTITY acute \"&#180;\"> \n" +
                "]>" + fileContent;


        SAXParserFactory parserFactor = SAXParserFactory.newInstance();
        SAXParser parser = parserFactor.newSAXParser();
        SAXHandler handler = new SAXHandler();

        ByteArrayInputStream bis = new ByteArrayInputStream(fileContent.getBytes());

        parser.parse(bis, handler);

        //Parsing the substance list obtained from XML
        int count = 0;
        for (XmlRow subsRow : handler.xmlRowList) {
            if (subsRow.cellList.size() > 19) {
                String substance = subsRow.cellList.get(0);
                if (substance.equals("Substance")) {
                    continue;
                }

                count++;
                String adi = subsRow.cellList.get(15); // column P
                String arfd = subsRow.cellList.get(18); // column S

                System.out.println("Substance name='" + substance + "', ADI='" + adi + "', ARfD='" + arfd + "'");
            }
        }

        if (file.exists()) {
            System.out.println("delete file-> " + file.getAbsolutePath());
            if (!file.delete()) {
                System.out.println("file '" + file.getAbsolutePath() + "' was not deleted!");
            }
        }
        System.out.println("getAndParseFile finished, processed " + count + " substances!");
    }


    /**
     * Removes the supplied line numbers from a file
     *
     * @param file the path to the file
     * @param lineNumbersToRemove a list of line numbers to remove
     */
    private static void removeLineFromFile(String file, Integer... lineNumbersToRemove) {

        BufferedReader br = null;
        PrintWriter pw = null;
        try {

            File inFile = new File(file);

            if (!inFile.isFile()) {
                System.out.println("file '" + file + "' is not an existing file");
                return;
            }

            //Construct the new file that will later be renamed to the original filename.
            File tempFile = new File(inFile.getAbsolutePath() + ".tmp");

            br = new BufferedReader(new FileReader(file));
            pw = new PrintWriter(new FileWriter(tempFile));

            String line = null;

            //Read from the original file and write to the new
            //unless content matches data to be removed.
            int lineCount = 1;
            while ((line = br.readLine()) != null) {
                boolean remove = false;
                for (Integer lineToRemove : lineNumbersToRemove) {
                    if (lineToRemove == lineCount)
                        remove = true;
                }
                //if (!line.trim().equals(lineToRemove)) {
                if (!remove) {
                    pw.println(line);
                    pw.flush();
                }
                lineCount++;
            }
            pw.close();
            br.close();

            //Delete the original file
            if (!inFile.delete()) {
                System.out.println("Could not delete original file");
                return;
            }

            //Rename the new file to the filename the original file had.
            if (!tempFile.renameTo(inFile))
                System.out.println("Could not rename temp file");

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}

