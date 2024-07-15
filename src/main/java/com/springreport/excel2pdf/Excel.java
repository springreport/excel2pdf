package com.springreport.excel2pdf;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

/**
 * Created by cary on 6/15/17.
 */
public class Excel {

    protected Workbook wb;
    protected Sheet sheet;

    public Excel(InputStream is,int sheetIndex) {
        try {
            this.wb = WorkbookFactory.create(is);
          //强制计算
            try {
            	wb.getCreationHelper().createFormulaEvaluator().evaluateAll();
			} catch (Exception e) {
				e.printStackTrace();
			}
            this.sheet = wb.getSheetAt(sheetIndex);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public Sheet getSheet() {
        return sheet;
    }

    public Workbook getWorkbook(){
        return wb;
    }
}
