package com.oy.excel.fast.parser.row;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Administrator on 2014-12-13.
 */
public class XRow {
    private int rowIndex;// 0-base
    private boolean notExist;
    private List<String> cellValus;
    public XRow(int rowIndex){
         this.rowIndex = rowIndex;
         this.cellValus = new ArrayList<String>();
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public List<String> getCellValus() {
        return cellValus;
    }

    public void setCellValus(List<String> cellValus) {
        this.cellValus = cellValus;
    }
    public void setCellValue(String value){
        this.cellValus.add(value);
    }

    public boolean isNotExist() {
        return notExist;
    }

    public void setNotExist(boolean notExist) {
        this.notExist = notExist;
    }

    @Override
    public String toString(){
        System.out.print("第" + rowIndex + "行: ");
        for(String s : this.cellValus){
            System.out.print(s + " ");
        }
        System.out.println("");
        return "";
    }
}
