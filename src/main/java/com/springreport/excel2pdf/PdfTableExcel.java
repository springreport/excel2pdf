package com.springreport.excel2pdf;

import com.itextpdf.text.Font;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.awt.font.FontRenderContext;
import java.awt.geom.AffineTransform;
import java.io.IOException;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Created by cary on 6/15/17.
 */
public class PdfTableExcel {
    protected ExcelObject excelObject;
    protected Excel excel;
    protected boolean setting = false;

    /**
     * <p>Description: Constructor</p>
     *
     * @param excelObject
     */
    public PdfTableExcel(ExcelObject excelObject) {
        this.excelObject = excelObject;
        this.excel = excelObject.getExcel();
    }
    
    /**
     * <p>Description: 获取转换过的Excel内容Table</p>
     *
     * @return PdfPTable
     * @throws BadElementException
     * @throws MalformedURLException
     * @throws IOException
     */
    public List<PdfPTable> getTable() throws BadElementException, MalformedURLException, IOException {
        Sheet sheet = this.excel.getSheet();
        return toParseContent(sheet);
    }
    
    public ResMobileInfos getMobileInfos() throws BadElementException, MalformedURLException, IOException {
    	 Sheet sheet = this.excel.getSheet();
    	 ResMobileInfos result = toParseCells(sheet);
    	 return result;
    }
    
    protected List<PdfPTable> toParseContent(Sheet sheet) throws BadElementException, MalformedURLException, IOException {
    	List<PdfPTable> result = new ArrayList<>();
    	int rows = 0;
    	if(this.excelObject.getEndx().intValue() != 0)
    	{
    		rows = this.excelObject.getEndx().intValue()+1;
    	}else {
    		rows = sheet.getPhysicalNumberOfRows();
    	}
        List<PdfPCell> cells = new ArrayList<PdfPCell>();
        Set<String> mergeCells = new HashSet<>();
        JSONObject colhidden = this.excelObject.getColhidden();
        JSONObject rowhidden = this.excelObject.getRowhidden();
        Set<String> rowhiddenkeys = rowhidden.keySet();
        Map<String, String> hiddenCellsReplace = new HashMap<>();
        Map<String, Cell> hiddenCells = new HashMap<>();
        Map<String, Integer> rowSpan = new HashMap<>();
        Map<String, Integer> colSpan = new HashMap<>();
        Map<String, CellStyle> splitMergeCells = new HashMap<>();//分割的合并单元格
        Map<String, Integer> splitMergeCellsRowSpan = new HashMap<>();
        Map<String, Integer> splitMergeCellsColSpan = new HashMap<>();
        int starty = this.excelObject.getStarty();
        int endy = this.excelObject.getStarty();
        PdfContentByte canvas = this.excelObject.getWriter().getDirectContent();
        JSONArray pageDivider = null;
        Map<Integer, Integer> pageRows = new HashMap<>();
        int fixedHeader = this.excelObject.getPrintSettings().getFixedHeader().intValue();
        int fixedHeaderStart = 0;
        if(this.excelObject.getPrintSettings().getFixedHeaderStart() != null) {
        	fixedHeaderStart = this.excelObject.getPrintSettings().getFixedHeaderStart().intValue();
        }
        int fixedHeaderEnd = 0;
        if(this.excelObject.getPrintSettings().getFixedHeaderEnd() != null) {
        	fixedHeaderEnd = this.excelObject.getPrintSettings().getFixedHeaderEnd().intValue();
        }
        int headerRows = 0;
        if(fixedHeader == 1) {
        	headerRows = (fixedHeaderEnd - this.excelObject.getStartx())-(fixedHeaderStart - this.excelObject.getStartx())+1;
        	if(headerRows < 1) {
        		headerRows = 1;
        	}
        }
        if(this.excelObject.getPrintSettings().getHorizontalPage().intValue() == 1)
        {//横向分页
        	pageDivider = this.excelObject.getPrintSettings().getPageDivider();
        	if(pageDivider!=null && pageDivider.size() > 0) {
        		int last = pageDivider.getIntValue(pageDivider.size() -1);
            	if(last > this.excelObject.getEndy())
            	{
            		pageDivider.set(pageDivider.size() -1, this.excelObject.getEndy());
            	}else if(last < this.excelObject.getEndy()) {
            		pageDivider.add(this.excelObject.getEndy());
            	}
        	}else {
        		pageDivider = new JSONArray();
            	pageDivider.add(this.excelObject.getEndy());
        	}
        }else {
        	pageDivider = new JSONArray();
        	pageDivider.add(this.excelObject.getEndy());
        }
        float height = 0;
        Map<Integer, Float> rowHeightsMap = new HashMap<>();//记录行高
        for (int t = 0; t < pageDivider.size(); t++) {
        	 float[] widths = null;
             float mw = 0;
             Map<PdfPCell, XxbtDto> xxbtCells = new HashMap<>();
			if(pageDivider.getIntValue(t)<starty)
			{
				continue;
			}else {
				cells = new ArrayList<PdfPCell>();
				endy = pageDivider.getIntValue(t)+1;
				
				for (int i = this.excelObject.getStartx(); i < rows; i++) {
		        	if(rowhidden.get(String.valueOf(i)) != null)
		        	{//隐藏行
 		        		Row row = sheet.getRow(i);
		                if (row == null) {
		                    continue;
		                }
 		                for (int j = starty; j < endy; j++) {
		                 	 Cell cell = row.getCell(j);
		                      if (cell == null) {
		                         cell = row.createCell(j);
		                     }
		                     CellRangeAddress range = getColspanRowspanByExcel(row.getRowNum(), cell.getColumnIndex());
		                     if(range != null)
		                     {
		                    	 Set<String> merge = new HashSet<>();
		                    	 int rowspan = range.getLastRow() - range.getFirstRow() + 1;
		                         int colspan = range.getLastColumn() - range.getFirstColumn() + 1;
		                    	 this.getMergeCells(i, j, rowspan, colspan, merge);
		                    	 this.getMergeCells(i, j, rowspan, colspan, mergeCells);
		                    	 if(colhidden.containsKey(String.valueOf(j)))
		                    	 {
		                     		 for(String key:merge)
		                    		 {
		                    			 if(!key.split("_")[1].equals(String.valueOf(j)) && !rowhidden.containsKey(key.split("_")[0]) && !colhidden.containsKey(key.split("_")[1]))
		                    			 {
		                    				 hiddenCellsReplace.put(key,i+"_"+j);
		                    				 hiddenCells.put(i+"_"+j, cell);
		                    				 colSpan.put(i+"_"+j, colspan);
		                    				 int hiddenRows = 0;
		                    				 for(String rowKey:rowhiddenkeys)
		                    				 {
		                    					 int x = Integer.parseInt(rowKey);
		                    					 if(x >= range.getFirstRow() && x <= range.getLastRow())
		                    					 {
		                    						 hiddenRows ++ ;
		                    					 }
		                    				 }
		                    				 rowSpan.put(i+"_"+j, rowspan - hiddenRows);
		                    				 break;
		                    			 }
		                    		 }
		                    	 }else {
		                    		 for(String key:merge)
		                    		 {
		                    			 if(!key.split("_")[0].equals(String.valueOf(i)) && !rowhidden.containsKey(key.split("_")[0]))
		                    			 {
		                    				 hiddenCellsReplace.put(key,i+"_"+j);
		                    				 hiddenCells.put(i+"_"+j, cell);
		                    				 colSpan.put(i+"_"+j, colspan);
		                    				 int hiddenRows = 0;
		                    				 for(String rowKey:rowhiddenkeys)
		                    				 {
		                    					 int x = Integer.parseInt(rowKey);
		                    					 if(x >= range.getFirstRow() && x <= range.getLastRow())
		                    					 {
		                    						 hiddenRows ++ ;
		                    					 }
		                    				 }
		                    				 rowSpan.put(i+"_"+j, rowspan - hiddenRows);
		                    				 break;
		                    			 }
		                    		 }
		                    	 }
		                     }
		                }
		        		continue;
		        	}
		            Row row = sheet.getRow(i);
		            if (row == null) {
		                continue;
		            }
		            float[] cws = new float[endy-starty];
		            float rowHeight = 0;
		            for (int j = starty; j < endy; j++) {
		                Cell cell = row.getCell(j);
		                if (cell == null) {
		                    cell = row.createCell(j);
		                }
		                float cw = getPOIColumnWidth(cell);
 		                if(colhidden.get(String.valueOf(j)) != null)
		            	{
		                	cws[cell.getColumnIndex()-starty] = 0;
		            	}else {
		            		cws[cell.getColumnIndex()-starty] = cw;
		            	}
		                int rowspan = 1;
		                int colspan = 1;
		                Cell hiddenCell = null;
		                PdfPCell pdfpCell = new PdfPCell();
		                if(mergeCells.contains(i+"_"+j))
		                {
		                	if(hiddenCellsReplace.containsKey(i+"_"+j))
		                	{
		                		hiddenCell = hiddenCells.get(hiddenCellsReplace.get(i+"_"+j));
		                		colspan = colSpan.get(hiddenCellsReplace.get(i+"_"+j));
		                		rowspan = rowSpan.get(hiddenCellsReplace.get(i+"_"+j));
		                	}else {
		                		if(splitMergeCells.containsKey(i+"_"+j)) {
		                			addBorderByExcel(pdfpCell, splitMergeCells.get(i+"_"+j));
		                			pdfpCell.setRowspan(splitMergeCellsRowSpan.get(i+"_"+j));
		                			pdfpCell.setColspan(splitMergeCellsColSpan.get(i+"_"+j));
		                			pdfpCell.setFixedHeight(this.getPixelHeight(splitMergeCellsRowSpan.get(i+"_"+j),row.getRowNum(),sheet,rowhidden,rowHeightsMap,t,cws,(XSSFCellStyle) cell.getCellStyle(),j,starty,pdfpCell,splitMergeCellsColSpan.get(i+"_"+j)));
		                			cells.add(pdfpCell);
		                		}
		                		continue;
		                	}
		                }
		                pdfpCell.setVerticalAlignment(hiddenCell==null?getVAlignByExcel(cell.getCellStyle().getVerticalAlignment().getCode()):getVAlignByExcel(hiddenCell.getCellStyle().getVerticalAlignment().getCode()));
		                pdfpCell.setHorizontalAlignment(hiddenCell==null?getHAlignByExcel(cell.getCellStyle().getAlignment().getCode()):getHAlignByExcel(hiddenCell.getCellStyle().getAlignment().getCode()));

		                pdfpCell.setPhrase(getPhrase(hiddenCell==null?cell:hiddenCell));
		                XSSFColor background = hiddenCell==null?(XSSFColor)cell.getCellStyle().getFillForegroundColorColor():(XSSFColor)hiddenCell.getCellStyle().getFillForegroundColorColor();
		                if(background != null)
		                {
		                	List<Integer> rgb = this.getColor(background);
		                	pdfpCell.setBackgroundColor(new BaseColor(rgb.get(0),rgb.get(1),rgb.get(2)));
		                }
		                addBorderByExcel(pdfpCell, cell.getCellStyle());
		                // 执行此方法在poi导出为 Workbook 是 SXSSFWorkbook的类型时，此方法会导致转换cell 为""
		                //cell.setCellType(CellType.STRING);
		                CellRangeAddress range = getColspanRowspanByExcel(row.getRowNum(), cell.getColumnIndex());
		                
		                if (range != null && hiddenCell == null) {
		                    rowspan = range.getLastRow() - range.getFirstRow() + 1;
		                    colspan = range.getLastColumn() - range.getFirstColumn() + 1;
		                    int colspan2 = range.getLastColumn() - range.getFirstColumn() + 1;
		                    if(range.getLastColumn()<=pageDivider.getIntValue(t))
		                    {
		                    	colspan = range.getLastColumn() - range.getFirstColumn() + 1;
		                    }else {
		                    	colspan = pageDivider.getIntValue(t) - range.getFirstColumn() + 1;
		                    }
		                    this.getMergeCells(i, j, rowspan, colspan2, mergeCells);
		                    for (int k = 1; k < colspan2; k++) {
		                    	int y = j + k;
		                    	Cell cell2 = row.getCell(y);
		                    	if (cell2 == null) {
		                    		cell2 = row.createCell(y);
				                }
		                    	if(y >=endy)
		                    	{
		                    		for (int l = t; l < pageDivider.size(); l++) {
										if(y == pageDivider.getIntValue(l)+1)
										{
											int pageColspan = 1;
											if(colspan2-k < pageDivider.getIntValue(l+1)-pageDivider.getIntValue(l))
											{
												pageColspan = colspan2-k;
											}else {
												pageColspan = pageDivider.getIntValue(l+1)-pageDivider.getIntValue(l);
											}
											splitMergeCells.put(i+"_"+y, cell.getCellStyle());
				                    		int rowSpan3 = rowspan;
				                    		if(rowspan > 1) {
				    		                	for (int m = 1; m < rowspan; m++) {
				    		                		if(rowhidden.containsKey(String.valueOf(row.getRowNum()+m))) {
				    		                			rowSpan3 = rowSpan3 - 1;
				    		                		}
				    		                	}
				    		                }
				                    		splitMergeCellsRowSpan.put(i+"_"+y, rowSpan3);
				                    		splitMergeCellsColSpan.put(i+"_"+y, pageColspan);
										}
									}
		                    		
		                    		continue;
		                    	}
		                    	float cw2 = 0;
		                    	if(colhidden.containsKey(String.valueOf(y)))
		                    	{
		                    		cw2 = 0;
		                    	}else {
		                    		cw2 = getPOIColumnWidth(cell2);
		                    	}
		                    	cws[cell2.getColumnIndex()-starty] = cw2;
							}
		                }
		                int rowSpan2 = rowspan;
		                if(rowspan > 1) {
		                	for (int k = 1; k < rowspan; k++) {
		                		if(rowhidden.containsKey(String.valueOf(row.getRowNum()+k))) {
		                			rowSpan2 = rowSpan2 - 1;
		                		}
		                	}
		                }
		                pdfpCell.setColspan(colspan);
		                pdfpCell.setRowspan(rowSpan2);
//		            	if (sheet.getDefaultRowHeightInPoints() != row.getHeightInPoints()) {
//		            		pdfpCell.setFixedHeight(this.getPixelHeight(rowspan,row.getRowNum(),sheet));
//			            }else {
//			            	pdfpCell.setFixedHeight(this.getPixelHeight(rowspan,row.getRowNum(),sheet));
//			            }
		                pdfpCell.setFixedHeight(this.getPixelHeight(rowspan,row.getRowNum(),sheet,rowhidden,rowHeightsMap,t,cws,(XSSFCellStyle) cell.getCellStyle(),j,starty,pdfpCell,colspan));
		                addImageByPOICell(pdfpCell, cell, cw);
		                if(j == starty)
		                {
		                	rowHeight = pdfpCell.getFixedHeight();
		                }
		            	if(this.excelObject.getXxbtScreenshot() != null && this.excelObject.getXxbtScreenshot().containsKey(i+"_"+j))
		                {
		            		XxbtDto xxbtDto = new XxbtDto();
		            		xxbtDto.setRowNo(row.getRowNum());
		            		xxbtDto.setRowSpan(rowspan);
		            		xxbtDto.setColumnNo(j);
		            		xxbtDto.setStartx(this.excelObject.getStartx());
		            		xxbtDto.setStarty(starty);
		            		xxbtDto.setColSpan(colspan);
		            		xxbtDto.setCellValue(this.excelObject.getXxbtScreenshot().getString(i+"_"+j));
		            		Font font = pdfpCell.getPhrase().getFont();
		            		xxbtDto.setFont(font);
		            		xxbtCells.put(pdfpCell, xxbtDto);
		            		pdfpCell.setPhrase(null);
		                }
		            	pdfpCell.setNoWrap(false);
		                cells.add(pdfpCell);
		                if(hiddenCell == null)
		                {
		                	j += colspan - 1;
		                }
		            }

		            float rw = 0;
		            for (int j = 0; j < cws.length; j++) {
		                rw += cws[j];
		            }
		            if (rw > mw || mw == 0) {
		                widths = cws;
		                mw = rw;
		            }
 		            if(t == 0)
	                {
	                	height = height + rowHeight;
	                	if(height > this.excelObject.getTableHeight())
	                	{
	                		pageRows.put(i-1, i-1);
	                		height = rowHeight;
	                	}else if(height == this.excelObject.getTableHeight()) {
	                		pageRows.put(i, i);
	                		height = 0;
	                	}
	                }else {
	                	if(pageRows.containsKey(i))
	                	{
	                		PdfPTable table = new PdfPTable(widths);
//	                		if(fixedHeader == 1) {
//	        					table.setHeaderRows(headerRows);
//	        				}
	                		table.setTotalWidth(excelObject.getTableWidth());
	            		    table.setWidthPercentage(100);
	            		    for (PdfPCell pdfpCell : cells) {
	            		    	pdfpCell.setNoWrap(false);
	            		        table.addCell(pdfpCell);
	            		        if(xxbtCells.containsKey(pdfpCell)) {
	        			        	drawLine(xxbtCells.get(pdfpCell),table.getAbsoluteWidths(),rowHeightsMap,canvas);
	        			        }
	            		    }
	            		    try {
	            		    	drawPicture(excelObject.getImageInfos(), widths, rowHeightsMap, canvas, starty);
							} catch (Exception e) {
								e.printStackTrace();
							}
	            		    table.setKeepTogether(true);
	            		    table.getDefaultCell().setBorder(PdfPCell.NO_BORDER);
	            		    result.add(table);
	            		    cells = new ArrayList<PdfPCell>();
	                	}else if(i == this.excelObject.getEndx()) {
	                		PdfPTable table = new PdfPTable(widths);
//	                		if(fixedHeader == 1) {
//	        					table.setHeaderRows(headerRows);
//	        				}
	                		table.setTotalWidth(excelObject.getTableWidth());
	            		    table.setWidthPercentage(100);
	            		    for (PdfPCell pdfpCell : cells) {
	            		    	pdfpCell.setNoWrap(false);
	            		        table.addCell(pdfpCell);
	            		        if(xxbtCells.containsKey(pdfpCell)) {
	        			        	drawLine(xxbtCells.get(pdfpCell),table.getAbsoluteWidths(),rowHeightsMap,canvas);
	        			        }
	            		    }
	            		    try {
	            		    	drawPicture(excelObject.getImageInfos(), widths, rowHeightsMap, canvas, starty);
							} catch (Exception e) {
								e.printStackTrace();
							}
	            		    table.setKeepTogether(true);
	            		    table.getDefaultCell().setBorder(PdfPCell.NO_BORDER);
	            		    result.add(table);
	                	}
	                }
		        }
			}
			if(t == 0)
			{
				PdfPTable table = new PdfPTable(widths);
				if(fixedHeader == 1) {
					table.setHeaderRows(headerRows);
				}
				table.setTotalWidth(excelObject.getTableWidth());
			    table.setWidthPercentage(100);
			    for (PdfPCell pdfpCell : cells) {
			    	pdfpCell.setNoWrap(false);
			        table.addCell(pdfpCell);
			        if(xxbtCells.containsKey(pdfpCell)) {
			        	drawLine(xxbtCells.get(pdfpCell),table.getAbsoluteWidths(),rowHeightsMap,canvas);
			        }
			    }
			    try {
    		    	drawPicture(excelObject.getImageInfos(), widths, rowHeightsMap, canvas, starty);
				} catch (Exception e) {
					e.printStackTrace();
				}
			    table.setKeepTogether(true);
			    table.getDefaultCell().setBorder(PdfPCell.NO_BORDER);
			    result.add(table);	
			}
			starty = endy;
		}

        return result;
    }
    
    private void drawLine(XxbtDto xxbtDto,float[] widths, Map<Integer, Float> rowHeightsMap,PdfContentByte canvas) {
    	if(xxbtDto.getCellValue() != null && xxbtDto.getCellValue() != "") {
    		String[] titles = xxbtDto.getCellValue().split("\\|");
    		if(titles.length > 1) {
    			float left = 0;
    	    	float top = 0;
    	    	float right = 0;
    	    	float bottom = 0;
    	    	float width = 0;
    	    	float cellHeight = 0;
    	    	float cellWidth = 0;
    	    	for (int i = xxbtDto.getRowNo(); i < xxbtDto.getRowNo()+xxbtDto.getRowSpan(); i++) {
    	    		cellHeight = cellHeight + rowHeightsMap.get(i);
    	    	}
    	    	for (int i = 0; i < xxbtDto.getColSpan(); i++) {
    	    		cellWidth = cellWidth + widths[xxbtDto.getColumnNo()-xxbtDto.getStarty()+i];
    	    	}
    	    	for (int i = 0; i < (xxbtDto.getColumnNo()-xxbtDto.getStarty()); i++) {
    	    		width = width + widths[i];
    			}
    	    	left = this.excelObject.getDocument().leftMargin() + width;
    	    	width = 0;
    	    	for (int i = 0; i < (xxbtDto.getColumnNo()+xxbtDto.getColSpan()-xxbtDto.getStarty()); i++) {
    	    		width = width + widths[i];
    			}
    	    	right = this.excelObject.getDocument().leftMargin() + width;
    	    	float height = 0;
    	    	for (int i = xxbtDto.getStartx(); i < xxbtDto.getRowNo(); i++) {
    	    		height = height + rowHeightsMap.get(i);
    			}
    	    	top = this.excelObject.getDocument().getPageSize().getTop()- (this.excelObject.getDocument().topMargin()+height);
    	    	height = 0;
    	    	for (int i = xxbtDto.getStartx(); i < xxbtDto.getRowNo()+xxbtDto.getRowSpan(); i++) {
    	    		height = height + rowHeightsMap.get(i);
    			}
    	    	bottom = this.excelObject.getDocument().getPageSize().getTop()- (this.excelObject.getDocument().topMargin()+height);
    			int lineCount = titles.length - 1;
    			BaseFont baseFont = new Font(Resource.BASE_FONT_CHINESE , 10, Font.NORMAL).getBaseFont();
    			if(lineCount == 1) {
        	    	canvas.saveState();
        	    	canvas.moveTo(left, top);
        	    	canvas.lineTo(right,bottom);
        	    	canvas.setFontAndSize(baseFont, xxbtDto.getFont().getSize());
        	    	canvas.setTextMatrix(left+5,bottom+5);
        	    	canvas.showText(titles[0]);
        	    	canvas.setTextMatrix(left+cellWidth/2,top-15);
        	    	canvas.showText(titles[1]);
        	    	canvas.stroke();
        	    	canvas.restoreState();
    			}else if(lineCount == 2) {
    				canvas.saveState();
    				canvas.setFontAndSize(baseFont, xxbtDto.getFont().getSize());
    				canvas.moveTo(left, top);
    				canvas.lineTo(left+cellWidth/2,bottom);
    				canvas.moveTo(left, top);
    				canvas.lineTo(right,top-cellHeight/2);
    				canvas.setTextMatrix(left+5,bottom+5);
        	    	canvas.showText(titles[0]);
        	    	canvas.setTextMatrix(left+cellWidth/2,top-cellHeight*2/3);
        	    	canvas.showText(titles[1]);
        	    	canvas.setTextMatrix(left+cellWidth/2,top-15);
        	    	canvas.showText(titles[2]);
    				canvas.stroke();
        	    	canvas.restoreState();
    			}else if(lineCount == 3) {
    				canvas.saveState();
    				canvas.setFontAndSize(baseFont, xxbtDto.getFont().getSize());
    				canvas.moveTo(left, top);
    				canvas.lineTo(left+cellWidth/2,bottom);
    				canvas.moveTo(left, top);
    				canvas.lineTo(right,bottom);
    				canvas.moveTo(left, top);
    				canvas.lineTo(right,top-cellHeight/2);
    				canvas.setTextMatrix(left+5,bottom+5);
        	    	canvas.showText(titles[0]);
    				canvas.stroke();
    				canvas.setTextMatrix(left+cellWidth/2,bottom+5);
        	    	canvas.showText(titles[1]);
        	    	canvas.setTextMatrix(left+cellWidth*2/3,top-cellHeight*3/5);
        	    	canvas.showText(titles[2]);
        	    	canvas.setTextMatrix(left+cellWidth/2,top-15);
        	    	canvas.showText(titles[3]);
        	    	canvas.restoreState();
    			}
    			
    		}
    	}
    	
    }
    
    private void getMergeCells(int r,int c,int rowSpan,int colSpan,Set<String> mergeCells){
    	for (int i = 0; i < rowSpan; i++) {
			for (int j = 0; j < colSpan; j++) {
				if(i == 0 && j==0)
				{
				}else {
					mergeCells.add((r+i)+"_"+(c+j));
				}
			}
		}
    }


    protected Phrase getPhrase(Cell cell) {
    	if(isRichTextString(cell))
    	{
    		return getRichTextPhrase(cell);
    	}else {
    		return new Phrase(String.valueOf(getCellValue(cell)), getFontByExcel((XSSFCellStyle) cell.getCellStyle()));
    	}
//        if (this.setting || this.excelObject.getAnchorName() == null) {
//        	if(isRichTextString(cell))
//        	{
//        		return getRichTextPhrase(cell);
//        	}else {
//        		return new Phrase(String.valueOf(getCellValue(cell)), getFontByExcel((XSSFCellStyle) cell.getCellStyle()));
//        	}
//        }
//        Anchor anchor = new Anchor(String.valueOf(getCellValue(cell)), getFontByExcel((XSSFCellStyle) cell.getCellStyle()));
//        anchor.setName(this.excelObject.getAnchorName());
//        this.setting = true;
//        return anchor;
    }

    protected boolean isRichTextString(Cell cell)
    {
    	if (cell.getCellType() == CellType.STRING)
    	{
    		XSSFRichTextString xSSFRichTextString = (XSSFRichTextString) cell.getRichStringCellValue();
    		if(xSSFRichTextString.numFormattingRuns() == 0)
    		{
    			return false;
    		}else {
    			return true;
    		}
    	}else {
    		return false;
    	}
    }
    
    protected Phrase getRichTextPhrase(Cell cell)
    {
    	Phrase phrase = new Phrase();
    	XSSFRichTextString xSSFRichTextString = (XSSFRichTextString) cell.getRichStringCellValue();
    	String cellVal = xSSFRichTextString.getString();
    	int size = xSSFRichTextString.numFormattingRuns();
    	int subLength = 0;
    	for (int t = 0; t < size; t++) {
    		XSSFFont font = xSSFRichTextString.getFontOfFormattingRun(t);
    		short typeOffset = font.getTypeOffset();
    		String value = cellVal.substring(subLength, subLength + xSSFRichTextString.getLengthOfFormattingRun(t));
    		if(org.apache.poi.ss.usermodel.Font.SS_SUB == typeOffset)
    		{//下标
    			Chunk sub = new Chunk(value,getFontByExcel(font,true));
    			sub.setTextRise(-font.getFontHeightInPoints()*2/5);
    			phrase.add(sub);
    		}else if(org.apache.poi.ss.usermodel.Font.SS_SUPER == typeOffset){
    			//上标
    			Chunk sup = new Chunk(value,getFontByExcel(font,true));
    			sup.setTextRise(font.getFontHeightInPoints()/2);
    			phrase.add(sup);
    		}else {
    			Chunk chunk = new Chunk(value, getFontByExcel(font,false));
    			phrase.add(chunk);
    		}
    		subLength = subLength + xSSFRichTextString.getLengthOfFormattingRun(t);
    	}
    	return phrase;
    }

    /**
     * 获取单元格值
     *
     * @return 单元格值
     */
    public Object getCellValue(Cell cell) {
        Object val = "";
        try {
            if (cell != null) {
                if (cell.getCellType() == CellType.NUMERIC) {
                    val = cell.getNumericCellValue();
                    DataFormatter formatter = new DataFormatter();
                    final CellStyle cellStyle = cell.getCellStyle();
                    val = formatter.formatRawCellContents(cell.getNumericCellValue(), cellStyle.getDataFormat(), cellStyle.getDataFormatString());
                }else if(cell.getCellType() == CellType.FORMULA) {
                	CellType resultType = cell.getCachedFormulaResultType();
                	if(resultType == CellType.NUMERIC) {
                		val = cell.getNumericCellValue();
                		val = LuckysheetUtil.formatValue(cell.getCellStyle().getDataFormatString(), val);
                	}else if(resultType == CellType.STRING) {
                		val = cell.getStringCellValue();
                	}else if(resultType == CellType.BOOLEAN) {
                		val = cell.getBooleanCellValue();
                	}
                }else if (cell.getCellType() == CellType.STRING) {
                    val = cell.getStringCellValue();
                } else if (cell.getCellType() == CellType.BOOLEAN) {
                    val = cell.getBooleanCellValue();
                } else if (cell.getCellType() == CellType.ERROR) {
                    val = cell.getErrorCellValue();
                }

            }
        } catch (Exception e) {
            return val;
        }
        return val;
    }


    protected void addImageByPOICell(PdfPCell pdfpCell, Cell cell, float cellWidth) throws BadElementException, MalformedURLException, IOException {
        String key = cell.getRowIndex() + "_" + cell.getColumnIndex();
        if(this.excelObject.getBackImages() != null & this.excelObject.getBackImages().containsKey(key)) {
        	return;
        }
    	POIImage poiImage = new POIImage().getCellImage(cell);

    	byte[] bytes = poiImage.getBytes();
        if (bytes != null) {
//           double cw = cellWidth;
//           double ch = pdfpCell.getFixedHeight();
//
//           double iw = poiImage.getDimension().getWidth();
//           double ih = poiImage.getDimension().getHeight();
//
//           double scale = cw / ch;
//
//           double nw = iw * scale;
//           double nh = ih - (iw - nw);
//
//        	POIUtil.scale(bytes , nw  , nh);
//        	byte[] barCodeByte = BarCodeUtil.generateBarcodeImage(cell.getStringCellValue(), 360, 80);
        	Image image = null;
        	XSSFFont font = (XSSFFont) this.excel.wb.getFontAt(cell.getCellStyle().getFontIndex());
        	String ff = font.getFontName();//字体名称
        	if(ff != null && ff.contains("barCode128")) {
        		int rowspan = 1;
        		int colspan = 1;
        		CellRangeAddress range = getColspanRowspanByExcel(cell.getRowIndex(), cell.getColumnIndex());
        		if(range != null) {
        			rowspan = range.getLastRow() - range.getFirstRow() + 1;
        			colspan = range.getLastColumn() - range.getFirstColumn() + 1;
        		}
        		int width = (int)(cellWidth*1.333*colspan);
        		if(width > this.excelObject.getTableWidth()) {
        			width =  (int) this.excelObject.getTableWidth();
        		}
        		int height = (int)(pdfpCell.getFixedHeight()*1);
        		byte[] barCodeByte = BarCodeUtil.generateBarcodeImage(cell.getStringCellValue(), width, (int)pdfpCell.getFixedHeight());
        		image = Image.getInstance(barCodeByte);
        	}else if(ff != null && ff.contains("qrCode")) {
//        		int rowspan = 1;
//        		int colspan = 1;
//        		CellRangeAddress range = getColspanRowspanByExcel(cell.getRowIndex(), cell.getColumnIndex());
//        		if(range != null) {
//        			rowspan = range.getLastRow() - range.getFirstRow() + 1;
//        			colspan = range.getLastColumn() - range.getFirstColumn() + 1;
//        		}
//        		byte[] qrCodeByte = QRCodeUtil.generateQRCodeImage(cell.getStringCellValue(), (int)(cellWidth*1.333*colspan), (int)(pdfpCell.getFixedHeight()/7.5*rowspan));
        		image = Image.getInstance(bytes);
        	}else {
        		image = Image.getInstance(bytes);
        	}
            pdfpCell.setImage(image);
        }
    }
    
    protected float getPixelHeight(int rowSpan,int rowNum,Sheet sheet,JSONObject rowhidden,Map<Integer, Float> rowHeightsMap,int page,float[] cws,XSSFCellStyle style,int colNum,int starty,PdfPCell cell,int colSpan) {
    	java.awt.Font f = new java.awt.Font("STSongStd-Light", style.getFont().getBold()?Font.BOLD:Font.NORMAL, style.getFont().getFontHeightInPoints());
    	Row row = null;
    	float pixel = 0;
    	for (int i = 0; i < rowSpan; i++) {
    		if(rowhidden.get(String.valueOf(rowNum+i)) != null)
    		{
    			rowHeightsMap.put(rowNum+i, 0f);
    			continue;
    		}else {
    			float poiHeight = 0;
    			row = sheet.getRow(rowNum+i);
//    			if(rowHeightsMap.containsKey(rowNum+i))
//    			{
//    				poiHeight = rowHeightsMap.get(rowNum+i);
//    				int ls = 0;
//    				if(this.excelObject.getWrapText().containsKey(rowNum+"_"+colNum+"_ls")) {
//						ls = this.excelObject.getWrapText().get(rowNum+"_"+colNum+"_ls");
//						cell.setLeading(ls, 1f);
//					}
//    			}else {
//    				if(this.excelObject.getWrapText()!=null && this.excelObject.getWrapText().containsKey((rowNum+i)+"_"+colNum)) {
////    					FontMetrics fm = sun.font.FontDesignMetrics.getMetrics(f);
////    					int chartWidth = fm.charWidth('国');
////    					int width = chartWidth * this.excelObject.getWrapText().get(rowNum+i);
////    					poiHeight = (float) (fm.getHeight() * width/(this.excelObject.getTableWidth()/cws.length));
//    					int ls = 0;
//    					if(this.excelObject.getWrapText().containsKey(rowNum+"_"+colNum+"_ls")) {
//    						ls = this.excelObject.getWrapText().get(rowNum+"_"+colNum+"_ls");
//    						cell.setLeading(ls, 1f);
//    					}
//    					FontRenderContext frc = new FontRenderContext(new AffineTransform(), true, true);
//    					java.awt.Font font = new java.awt.Font("微软雅黑", style.getFont().getBold()?Font.BOLD:Font.NORMAL, style.getFont().getFontHeightInPoints());
//    					String wordContent = "国";
//    					java.awt.Rectangle rec = font.getStringBounds(wordContent, frc).getBounds();
//    					int chartWidth = rec.width;
//    					int width = chartWidth * this.excelObject.getWrapText().get((rowNum+i)+"_"+colNum);
//    					poiHeight = (float) ((rec.height+ls) * width/(this.excelObject.getTableWidth()/cws.length));
//    					if(ls > 0) {
//    						poiHeight = poiHeight + rec.height+ls;
//    					}
//    					if(poiHeight > this.excelObject.getTableHeight()) {
//    						poiHeight = this.excelObject.getTableHeight();
//    					}else if(poiHeight < row.getHeightInPoints()) {
//    						poiHeight = sheet.getDefaultRowHeightInPoints();
//    					}
//    				}else {
//                		if(row == null) {
//                			poiHeight = sheet.getDefaultRowHeightInPoints();
//                		}else {
//                			poiHeight = row.getHeightInPoints();
//                		}
//    				}
////    				if(rowSpan > 30) {
////    					poiHeight = poiHeight / 2;
////    				}
//            		if(page == 0) {
//            			if(rowHeightsMap.containsKey(rowNum+i)) {
//            				if(poiHeight > rowHeightsMap.get(rowNum+i)) {
//            					rowHeightsMap.put(rowNum+i, poiHeight);
//            				}
//            			}else {
//            				rowHeightsMap.put(rowNum+i, poiHeight);
//            			}
//            		}
//    			}
    			if(this.excelObject.getWrapText()!=null && this.excelObject.getWrapText().containsKey((rowNum+i)+"_"+colNum)) {
					int ls = 0;
					if(this.excelObject.getWrapText().containsKey(rowNum+"_"+colNum+"_ls")) {
						ls = this.excelObject.getWrapText().get(rowNum+"_"+colNum+"_ls");
						cell.setLeading(ls, 1f);
					}
					FontRenderContext frc = new FontRenderContext(new AffineTransform(), true, true);
					java.awt.Font font = new java.awt.Font("微软雅黑", style.getFont().getBold()?Font.BOLD:Font.NORMAL, style.getFont().getFontHeightInPoints());
					String wordContent = "国";
					java.awt.Rectangle rec = font.getStringBounds(wordContent, frc).getBounds();
					int chartWidth = rec.width;
					int width = chartWidth * this.excelObject.getWrapText().get((rowNum+i)+"_"+colNum);
					float rows = width/(this.excelObject.getTableWidth()/cws.length*colSpan);
					poiHeight = (float) ((rec.height+ls) * rows);
//					if(ls > 0) {
//						poiHeight = poiHeight + rec.height+ls;
//					}
					if(poiHeight > this.excelObject.getTableHeight()) {
						poiHeight = this.excelObject.getTableHeight();
					}else if(poiHeight < row.getHeightInPoints()) {
						poiHeight = row.getHeightInPoints();
					}
				}else {
            		if(row == null) {
            			poiHeight = sheet.getDefaultRowHeightInPoints();
            		}else {
            			poiHeight = row.getHeightInPoints();
            		}
				}
        		if(page == 0) {
        			if(rowHeightsMap.containsKey(rowNum+i)) {
        				if(poiHeight > rowHeightsMap.get(rowNum+i)) {
        					rowHeightsMap.put(rowNum+i, poiHeight);
        				}
        			}else {
        				rowHeightsMap.put(rowNum+i, poiHeight);
        			}
        		}
        		pixel = pixel + poiHeight;
    		}
		}
        return pixel;
    }

    /**
     * <p>Description: 此处获取Excel的列宽像素(无法精确实现,期待有能力的朋友进行改善此处)</p>
     *
     * @param cell
     * @return 像素宽
     */
    protected float getPOIColumnWidth(Cell cell) {
    	return excel.getSheet().getColumnWidthInPixels(cell.getColumnIndex());
//        int poiCWidth = excel.getSheet().getColumnWidth(cell.getColumnIndex());
//        // com.itextpdf.text.pdf.PdfPTable.calculateWidths,此方法已经等比例转换了。不知道为什么还需要转换
//        // int colWidthpoi = poiCWidth;
//        // int widthPixel = 0;
//        // if (colWidthpoi >= 416) {
//        //     widthPixel = (int) (((colWidthpoi - 416.0) / 256.0) * 8.0 + 13.0 + 0.5);
//        // } else {
//        //     widthPixel = (int) (colWidthpoi / 416.0 * 13.0 + 0.5);
//        // }
//        return poiCWidth;
    }

    protected CellRangeAddress getColspanRowspanByExcel(int rowIndex, int colIndex) {
        CellRangeAddress result = null;
        Sheet sheet = excel.getSheet();
        int num = sheet.getNumMergedRegions();
        for (int i = 0; i < num; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.getFirstColumn() == colIndex && range.getFirstRow() == rowIndex) {
                result = range;
            }
        }
        return result;
    }

    protected boolean isUsed(int colIndex, int rowIndex) {
        boolean result = false;
        Sheet sheet = excel.getSheet();
        int num = sheet.getNumMergedRegions();
        for (int i = 0; i < num; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            if (firstRow < rowIndex && lastRow >= rowIndex) {
                if (firstColumn <= colIndex && lastColumn >= colIndex) {
                    result = true;
                }
            }
        }
        return result;
    }

    protected Font getFontByExcel(XSSFCellStyle style) {
    	short fontSize = 8;
    	String fontName = "";
    	try {
    		fontSize = style.getFont().getFontHeightInPoints();
    		fontName = style.getFont().getFontName();
		} catch (Exception e) {
			e.printStackTrace();
		}
    	XSSFColor color = (XSSFColor)style.getFont().getXSSFColor();
    	BaseColor baseColor = null;
        if(color != null)
        {
        	List<Integer> rgb = this.getColor(color);
        	baseColor = new BaseColor(rgb.get(0),rgb.get(1),rgb.get(2));
        }
        Font result = new Font(Resource.getFont(fontName), fontSize, Font.NORMAL,baseColor==null?BaseColor.BLACK:baseColor);
        org.apache.poi.ss.usermodel.Font font = style.getFont();

        if (font.getBold()) {
            result.setStyle(Font.BOLD);
        }

        FontUnderline underline = FontUnderline.valueOf(font.getUnderline());
        if (underline == FontUnderline.SINGLE) {
            String ulString = Font.FontStyle.UNDERLINE.getValue();
            result.setStyle(ulString);
        }
        return result;
    }
    
    protected Font getFontByExcel(XSSFFont font,boolean isSubSup) {
    	short fontSize = 8;
    	try {
    		fontSize = font.getFontHeightInPoints();
		} catch (Exception e) {
		}
    	if(isSubSup)
    	{
    		fontSize = (short) (fontSize*2/3);
    	}
    	XSSFColor color = (XSSFColor)font.getXSSFColor();
    	BaseColor baseColor = null;
        if(color != null)
        {
        	List<Integer> rgb = this.getColor(color);
        	baseColor = new BaseColor(rgb.get(0),rgb.get(1),rgb.get(2));
        }
        Font result = new Font(Resource.BASE_FONT_CHINESE, fontSize, Font.NORMAL,baseColor==null?BaseColor.BLACK:baseColor);
        Workbook wb = excel.getWorkbook();

//        int index = style.getFontIndexAsInt();
//        org.apache.poi.ss.usermodel.Font font = wb.getFontAt(index);

        if (font.getBold()) {
            result.setStyle(Font.BOLD);
        }

        FontUnderline underline = FontUnderline.valueOf(font.getUnderline());
        if (underline == FontUnderline.SINGLE) {
            String ulString = Font.FontStyle.UNDERLINE.getValue();
            result.setStyle(ulString);
        }
        return result;
    }

    protected void addBorderByExcel(PdfPCell cell, CellStyle style) {
        Workbook wb = excel.getWorkbook();
      //获取单元格边框信息
    	BorderStyle topBorderStyle = style.getBorderTop();
    	BorderStyle bottomBorderStyle = style.getBorderBottom();
    	BorderStyle leftBorderStyle = style.getBorderLeft();
    	BorderStyle rightBorderStyle = style.getBorderRight();
    	if(BorderStyle.NONE.getCode() == topBorderStyle.getCode())
    	{
    		cell.disableBorderSide(1);
    	}else {
    		cell.setBorderColorTop(new BaseColor(POIUtil.getBorderRBG(wb, style.getTopBorderColor())));
    	}
    	if(BorderStyle.NONE.getCode() == bottomBorderStyle.getCode())
    	{
    		cell.disableBorderSide(2);
    	}else {
    		cell.setBorderColorBottom(new BaseColor(POIUtil.getBorderRBG(wb, style.getBottomBorderColor())));
    	}
    	if(BorderStyle.NONE.getCode() == leftBorderStyle.getCode())
    	{
    		cell.disableBorderSide(4);
    	}else {
    		cell.setBorderColorLeft(new BaseColor(POIUtil.getBorderRBG(wb, style.getLeftBorderColor())));
    	}
    	if(BorderStyle.NONE.getCode() == rightBorderStyle.getCode())
    	{
    		cell.disableBorderSide(8);
    	}else {
    		cell.setBorderColorRight(new BaseColor(POIUtil.getBorderRBG(wb, style.getRightBorderColor())));
    	}
    }

    protected int getVAlignByExcel(short align) {
        int result = Element.ALIGN_BOTTOM;
        if (align == VerticalAlignment.BOTTOM.getCode()) {
            result = Element.ALIGN_BOTTOM;
        }
        if (align == VerticalAlignment.CENTER.getCode()) {
            result = Element.ALIGN_MIDDLE;
        }
        if (align == VerticalAlignment.TOP.getCode()) {
            result = Element.ALIGN_TOP;
        }
        return result;
    }

    protected int getHAlignByExcel(short align) {
        int result = 0;
        if (align == HorizontalAlignment.LEFT.getCode()) {
            result = Element.ALIGN_LEFT;
        }
        if (align == HorizontalAlignment.RIGHT.getCode()) {
            result = Element.ALIGN_RIGHT;
        }
//        if (align == HorizontalAlignment.JUSTIFY.getCode()) {
//            result = Element.ALIGN_JUSTIFIED;
//        }
        if (align == HorizontalAlignment.CENTER.getCode()) {
            result = Element.ALIGN_CENTER;
        }
        return result;
    }
    protected ResMobileInfos toParseCells(Sheet sheet) throws BadElementException, MalformedURLException, IOException{
    	ResMobileInfos result = new ResMobileInfos();
    	List<List<TableCell>> list = new ArrayList<>();
    	int rows = 0;
    	if(this.excelObject.getEndx().intValue() != 0)
    	{
    		rows = this.excelObject.getEndx().intValue()+1;
    	}else {
    		rows = sheet.getPhysicalNumberOfRows();
    	}
    	Set<String> mergeCells = new HashSet<>();
    	JSONObject colhidden = this.excelObject.getColhidden();
        JSONObject rowhidden = this.excelObject.getRowhidden();
        Set<String> colhiddenkeys = colhidden.keySet();
        Set<String> rowhiddenkeys = rowhidden.keySet();
        Map<String, String> hiddenCellsReplace = new HashMap<>();
        Map<String, Cell> hiddenCells = new HashMap<>();
        Map<String, Integer> rowSpan = new HashMap<>();
        Map<String, Integer> colSpan = new HashMap<>();
        Set<Integer> emptyRows = new HashSet<>();
    	for (int i = this.excelObject.getStartx(); i < rows; i++) {
     		List<TableCell> rowCells = new ArrayList<>();
     		if(rowhidden.get(String.valueOf(i)) != null)
        	{//隐藏行
        		Row row = sheet.getRow(i);
                if (row == null) {
                    continue;
                }
                int columns = 0;
                if(this.excelObject.getEndy().intValue() != 0)
                {
                	columns = this.excelObject.getEndy()+1;
                }else {
                	columns = row.getLastCellNum();
                }
                for (int j = this.excelObject.getStarty(); j < columns; j++) {
                	 Cell cell = row.getCell(j);
                     if (cell == null) {
                         cell = row.createCell(j);
                     }
                     CellRangeAddress range = getColspanRowspanByExcel(row.getRowNum(), cell.getColumnIndex());
                     if(range != null)
                     {
                    	 Set<String> merge = new HashSet<>();
                    	 int rowspan = range.getLastRow() - range.getFirstRow() + 1;
                         int colspan = range.getLastColumn() - range.getFirstColumn() + 1;
                    	 this.getMergeCells(i, j, rowspan, colspan, merge);
                    	 this.getMergeCells(i, j, rowspan, colspan, mergeCells);
                    	 if(colhidden.containsKey(String.valueOf(j)))
                    	 {
                    		 for(String key:merge)
                    		 {
                    			 if(!key.split("_")[1].equals(String.valueOf(j)) && !rowhidden.containsKey(key.split("_")[0]) && !colhidden.containsKey(key.split("_")[1]))
                    			 {
                    				 hiddenCellsReplace.put(key,i+"_"+j);
                    				 hiddenCells.put(i+"_"+j, cell);
                    				 int hiddenCols = 0;
                    				 for(String colKey:colhiddenkeys)
                    				 {
                    					 int y = Integer.parseInt(colKey);
                    					 if(y >= range.getFirstColumn() && y <= range.getLastColumn())
                    					 {
                    						 hiddenCols ++ ;
                    					 }
                    				 }
                    				 colSpan.put(i+"_"+j, colspan-hiddenCols);
                    				 int hiddenRows = 0;
                    				 for(String rowKey:rowhiddenkeys)
                    				 {
                    					 int x = Integer.parseInt(rowKey);
                    					 if(x >= range.getFirstRow() && x <= range.getLastRow())
                    					 {
                    						 hiddenRows ++ ;
                    					 }
                    				 }
                    				 rowSpan.put(i+"_"+j, rowspan - hiddenRows);
                    				 break;
                    			 }
                    		 }
                    	 }else {
                    		 for(String key:merge)
                    		 {
                    			 if(!key.split("_")[0].equals(String.valueOf(i)) && !rowhidden.containsKey(key.split("_")[0]))
                    			 {
                    				 hiddenCellsReplace.put(key,i+"_"+j);
                    				 hiddenCells.put(i+"_"+j, cell);
                    				 int hiddenCols = 0;
                    				 for(String colKey:colhiddenkeys)
                    				 {
                    					 int y = Integer.parseInt(colKey);
                    					 if(y >= range.getFirstColumn() && y <= range.getLastColumn())
                    					 {
                    						 hiddenCols ++ ;
                    					 }
                    				 }
                    				 colSpan.put(i+"_"+j, colspan-hiddenCols);
                    				 int hiddenRows = 0;
                    				 for(String rowKey:rowhiddenkeys)
                    				 {
                    					 int x = Integer.parseInt(rowKey);
                    					 if(x >= range.getFirstRow() && x <= range.getLastRow())
                    					 {
                    						 hiddenRows ++ ;
                    					 }
                    				 }
                    				 rowSpan.put(i+"_"+j, rowspan - hiddenRows);
                    				 break;
                    			 }
                    		 }
                    	 }
                     }
                }
        		continue;
        	}
    		Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            int columns = 0;
            columns = row.getLastCellNum();
            boolean isRowCellValueEmpty = true;//行单元格的值是否全是空
            for (int j = this.excelObject.getStarty(); j < columns; j++) {
            	XSSFCell cell = (XSSFCell) row.getCell(j);
                if (cell == null) {
                    cell = (XSSFCell) row.createCell(j);
                }
                int rowspan = 1;
                int colspan = 1;
                XSSFCell hiddenCell = null;
                if(mergeCells.contains(i+"_"+j))
                {
                	if(hiddenCellsReplace.containsKey(i+"_"+j))
                	{
                		hiddenCell = (XSSFCell) hiddenCells.get(hiddenCellsReplace.get(i+"_"+j));
                		colspan = colSpan.get(hiddenCellsReplace.get(i+"_"+j));
                		rowspan = rowSpan.get(hiddenCellsReplace.get(i+"_"+j));
                		mergeCells.remove(i+"_"+j);
                	}else {
                		isRowCellValueEmpty = false;
                    	continue;
                	}
                }
//                PdfPCell pdfpCell = new PdfPCell();
//                pdfpCell.setBackgroundColor(new BaseColor(POIUtil.getRGB(
//                        cell.getCellStyle().getFillForegroundColorColor())));
//                pdfpCell.setVerticalAlignment(getVAlignByExcel(cell.getCellStyle().getVerticalAlignment().getCode()));
//                pdfpCell.setHorizontalAlignment(getHAlignByExcel(cell.getCellStyle().getAlignment().getCode()));
//                pdfpCell.setPhrase(getPhrase(cell));
                TableCell tableCell = new TableCell();
                tableCell.setVerticalAlignment(hiddenCell==null?getVAlignByExcel(cell.getCellStyle().getVerticalAlignment().getCode()):getVAlignByExcel(hiddenCell.getCellStyle().getVerticalAlignment().getCode()));
                tableCell.setHorizontalAlignment(hiddenCell==null?getHAlignByExcel(cell.getCellStyle().getAlignment().getCode()):getHAlignByExcel(hiddenCell.getCellStyle().getAlignment().getCode()));
                XSSFColor color = hiddenCell==null?(XSSFColor)cell.getCellStyle().getFillForegroundColorColor():(XSSFColor)hiddenCell.getCellStyle().getFillForegroundColorColor();
                if(color != null)
                {//背景颜色
                	tableCell.setBackgroundColor("#"+color.getARGBHex().substring(2));
                }
                try {//字体颜色
                	tableCell.setFontColor("#"+(hiddenCell==null?cell.getCellStyle().getFont().getXSSFColor().getARGBHex().substring(2):hiddenCell.getCellStyle().getFont().getXSSFColor().getARGBHex().substring(2)));
				} catch (Exception e) {
				}
                try {//是否加粗
                	tableCell.setBold(hiddenCell==null?cell.getCellStyle().getFont().getBold():hiddenCell.getCellStyle().getFont().getBold());
				} catch (Exception e) {
				}
                try {//字体大小
                	tableCell.setFontSize(String.valueOf(hiddenCell==null?cell.getCellStyle().getFont().getFontHeightInPoints():hiddenCell.getCellStyle().getFont().getFontHeightInPoints()));
				} catch (Exception e) {
				}
                try {//是否斜体
                	tableCell.setItalic(hiddenCell==null?cell.getCellStyle().getFont().getItalic():hiddenCell.getCellStyle().getFont().getItalic());
				} catch (Exception e) {
				}
                try {//是否下划线斜体
                	tableCell.setUnderLine(hiddenCell==null?cell.getCellStyle().getFont().getUnderline():hiddenCell.getCellStyle().getFont().getUnderline());
				} catch (Exception e) {
				}
                CellRangeAddress range = getColspanRowspanByExcel(row.getRowNum(), cell.getColumnIndex());
                
                if (range != null && hiddenCell == null) {
                    rowspan = range.getLastRow() - range.getFirstRow() + 1;
                    colspan = range.getLastColumn() - range.getFirstColumn() + 1;
                    this.getMergeCells(i, j, rowspan, colspan, mergeCells);
                }
                tableCell.setColSpan(colspan);
                tableCell.setRowSpan(rowspan);
                if(colspan > 1 || rowspan > 1)
                {
                	tableCell.setIsMerge(1);
                }
                rowCells.add(tableCell);

                Object cellValue = getCellValue(hiddenCell==null?cell:hiddenCell);
                if (cellValue != null && String.valueOf(cellValue).trim().length() > 0)
                {
                	isRowCellValueEmpty = false;
                }
                tableCell.setCellValue(cellValue != null?String.valueOf(cellValue):null);
                tableCell.setX(i);
                tableCell.setY(j);
            }
            if(!isRowCellValueEmpty)
            {
            	list.add(rowCells);
            }else {
            	emptyRows.add(i);
            }
    	}
    	result.setTableCells(list);
    	result.setMergeCells(mergeCells);
    	result.setEmptyRows(emptyRows);
    	return result;
    }
    
    private List<Integer> getColor(XSSFColor xssfColor) {
    	List<Integer> result = new ArrayList<>();
    	int red = 0;
        int green = 0;
        int blue = 0;
        try {
        	byte[] rgb = xssfColor.getRGB();
            if (rgb != null) {
                red = (rgb[0] < 0) ? (rgb[0] + 256) : rgb[0];
                green = (rgb[1] < 0) ? (rgb[1] + 256) : rgb[1];
                blue = (rgb[2] < 0) ? (rgb[2] + 256) : rgb[2];
            }
		} catch (Exception e) {
		}
        result.add(red);
        result.add(green);
        result.add(blue);
        return result;
    }
    
    private void drawPicture(Map<String, Map<String, Object>> imageInfos,float[] widths, Map<Integer, Float> rowHeightsMap,PdfContentByte canvas,int starty) throws Exception{
    	float totalWidth = 0;
    	for (int i = 0; i < widths.length; i++) {
    		totalWidth = totalWidth + widths[i];
		}
    	float ratio = this.excelObject.getTableWidth()/totalWidth;
    	for(String mapKey : imageInfos.keySet()){
    		Map<String, Object> value = imageInfos.get(mapKey);
    		float imageHeight = 0;
    		float imageWidth = 0;
    		float left = 0;
    		float top = 0;
    		int col1 = (int)value.get("col1");
    		int row1 = (int)value.get("row1");
    		int col2 = (int)value.get("col2");
    		int row2 = (int)value.get("row2");
    		float dy1Percent = (float)value.get("dy1Percent");
    		float dy2Percent = (float)value.get("dy2Percent");
    		float dx1Percent = (float)value.get("dx1Percent");
    		float dx2Percent = (float)value.get("dx2Percent");
    		float dx1 = 0;//第一个单元格距离单元格左边框的横向距离
    		float dx2 = 0;//最后一个单元格距离单元格左边框的横向距离
    		float dy1 = 0;//第一个单元格距离单元格上边框的纵向距离
    		float dy2 = 0;//最后一个单元格距离单元格上边框的纵向距离
    		if(col2<starty)
    		{
    			continue;
    		}else if(col1 < starty && col2 >starty) {
    			continue;
    		}
    		if((col2-col1+1)>widths.length) {
    			col2 = col1 + widths.length-1;
    		}
    		Image background = Image.getInstance((byte[])value.get("pictureBytes"));
    		for (int i = col1; i <= col2; i++) {
    			if(i == col1) {
    				dx1 = widths[i-starty]*ratio * dx1Percent;
    				imageWidth = imageWidth + widths[i-starty]*ratio - dx1;
    			}else if(i == col2) {
    				dx2 = widths[i-starty]*ratio * dx2Percent;
    				imageWidth = imageWidth + dx2;
    			}else {
    				imageWidth = imageWidth + widths[i-starty]*ratio;
    			}
			}
    		
    		for (int i = row1; i <= row2; i++) {
    			if(i == row1) {
    				dy1 = rowHeightsMap.get(i) * dy1Percent;
    				imageHeight = imageHeight + rowHeightsMap.get(i) - dy1;
    			}else if(i == row2) {
    				dy2 = rowHeightsMap.get(i) * dy2Percent;
    				imageHeight = imageHeight + dy2;
    			}else {
    				imageHeight = imageHeight + rowHeightsMap.get(i);
    			}
			}
    		for (int i = 0; i <= (col1-starty); i++) {
				if(i == col1-starty) {
					left = left + dx1;
				}else {
					left = left + widths[i];
				}
			}
    		for (int i = this.excelObject.getStartx(); i <= row1; i++) {
    			if(i == row1) {
    				top = top + dy1;
    			}else {
    				top = top + rowHeightsMap.get(i);
    			}
    		}
    		background.setAbsolutePosition(left+this.excelObject.getDocument().leftMargin(), this.excelObject.getDocument().getPageSize().getTop()-this.excelObject.getDocument().topMargin()-imageHeight-top);
    		background.scaleAbsolute(imageWidth, imageHeight);
    		canvas.addImage(background);
    	}
    }

}