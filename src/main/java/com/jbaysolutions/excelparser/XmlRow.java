package com.jbaysolutions.excelparser;

import java.util.ArrayList;

/**
 * Created by IntelliJ IDEA.
 * User: Gus - gustavo.santos@jbaysolutions.com - http://gmsa.github.io/
 * Date: 04-03-2015
 * Time: 16:03
 */
class XmlRow {
    ArrayList<String> cellList = new ArrayList<>();

    @Override
    public String toString() {
        return cellList.toString();
    }
}
