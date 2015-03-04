package com.jbaysolutions.excelparser;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.ArrayList;
import java.util.List;

/**
 * The Handler for SAX Events.
 *
 * User: Gus - gustavo.santos@jbaysolutions.com - http://gmsa.github.io/
 *
 */
class SAXHandler extends DefaultHandler {

    List<XmlRow> xmlRowList = new ArrayList<>();
    XmlRow xmlRow = null;
    String content = null;

    @Override
    //Triggered when the start of tag is found.
    public void startElement(String uri, String localName, String qName, Attributes attributes)
            throws SAXException {
        switch (qName) {
            //Create a new Row object when the start tag is found
            case "Row":
                xmlRow = new XmlRow();
                break;
        }
    }

    @Override
    public void endElement(String uri, String localName,
                           String qName) throws SAXException {
        switch (qName) {
            case "Row":
                xmlRowList.add(xmlRow);
                break;
            case "Data":
                xmlRow.cellList.add(content);
                break;
        }
    }

    @Override
    public void characters(char[] ch, int start, int length)
            throws SAXException {
        content = String.copyValueOf(ch, start, length).trim();
    }
}
