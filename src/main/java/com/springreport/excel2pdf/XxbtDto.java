package com.springreport.excel2pdf;

import com.itextpdf.text.Font;

public class XxbtDto {

	 
	/**  
	 * @Fields rowNo : 横坐标
	 * @author caiyang
	 * @date 2024-04-29 06:19:59 
	 */  
	private int rowNo;
	
	/**  
	 * @Fields columnIndex : 列坐标
	 * @author caiyang
	 * @date 2024-04-29 05:37:05 
	 */  
	private int columnNo;
	
	/**  
	 * @Fields rowSpan : 合并行数
	 * @author caiyang
	 * @date 2024-04-29 06:20:34 
	 */  
	private int rowSpan = 1;
	
	/**  
	 * @Fields colSpan : 合并列数
	 * @author caiyang
	 * @date 2024-04-29 06:21:06 
	 */  
	private int colSpan = 1;
	
	/**  
	 * @Fields starty : 本页起始列
	 * @author caiyang
	 * @date 2024-04-29 08:42:48 
	 */   
	private int starty;
	
	/**  
	 * @Fields startx : 本页起始行
	 * @author caiyang
	 * @date 2024-04-29 09:49:54 
	 */  
	private int startx;
	
	/**  
	 * @Fields cellValue : 单元格值
	 * @author caiyang
	 * @date 2024-04-29 08:51:43 
	 */  
	private String cellValue;
	
	private Font font;

	public int getRowNo() {
		return rowNo;
	}

	public void setRowNo(int rowNo) {
		this.rowNo = rowNo;
	}


	public int getRowSpan() {
		return rowSpan;
	}

	public void setRowSpan(int rowSpan) {
		this.rowSpan = rowSpan;
	}

	public int getColSpan() {
		return colSpan;
	}

	public void setColSpan(int colSpan) {
		this.colSpan = colSpan;
	}

	public int getColumnNo() {
		return columnNo;
	}

	public void setColumnNo(int columnNo) {
		this.columnNo = columnNo;
	}

	public int getStarty() {
		return starty;
	}

	public void setStarty(int starty) {
		this.starty = starty;
	}

	public String getCellValue() {
		return cellValue;
	}

	public void setCellValue(String cellValue) {
		this.cellValue = cellValue;
	}

	public int getStartx() {
		return startx;
	}

	public void setStartx(int startx) {
		this.startx = startx;
	}

	public Font getFont() {
		return font;
	}

	public void setFont(Font font) {
		this.font = font;
	}
}
