package com.springreport.excel2pdf;

public class TableCell {
	
	/**  
	 * @Fields cellValue : 单元格值
	 * @author caiyang
	 * @date 2023-07-10 02:27:33 
	 */  
	private String cellValue;

	/**  
	 * @Fields rowSpan : 行合并
	 * @author caiyang
	 * @date 2023-07-10 10:14:17 
	 */  
	private int rowSpan = 1;
	
	/**  
	 * @Fields colSpan : 列合并
	 * @author caiyang
	 * @date 2023-07-10 10:14:07 
	 */  
	private int colSpan = 1;
	
	/**  
	 * @Fields x : 横坐标
	 * @author caiyang
	 * @date 2023-07-10 02:40:16 
	 */  
	private int x;
	
	/**  
	 * @Fields y : 纵坐标
	 * @author caiyang
	 * @date 2023-07-10 02:40:10 
	 */  
	private int y;
	
	/**  
	 * @Fields isMerge : 是否是合并单元格
	 * @author caiyang
	 * @date 2023-07-10 03:29:23 
	 */  
	private Integer isMerge = 2;
	
	/**  
	 * @Fields verticalAlignment : 垂直对齐方式 6下方对齐 5中间对齐 4 顶端对齐
	 * @author caiyang
	 * @date 2023-07-12 09:26:14 
	 */  
	private Integer verticalAlignment;
	
	/**  
	 * @Fields horizontalAlignment : 水平对齐方式 0左对齐 1居中对齐 2右对齐
	 * @author caiyang
	 * @date 2023-07-12 09:31:11 
	 */  
	private Integer horizontalAlignment;
	
	/**  
	 * @Fields backgroundColor : 背景颜色
	 * @author caiyang
	 * @date 2023-07-12 10:00:04 
	 */  
	private String backgroundColor;
	
	/**  
	 * @Fields fontColor : 字体颜色
	 * @author caiyang
	 * @date 2023-07-12 10:15:39 
	 */  
	private String fontColor;
	
	/**  
	 * @Fields bold : 是否加粗
	 * @author caiyang
	 * @date 2023-07-12 10:22:06 
	 */  
	private boolean bold = false;
	
	/**  
	 * @Fields fontSize : 字体大小
	 * @author caiyang
	 * @date 2023-07-12 10:24:07 
	 */  
	private String fontSize;
	
	/**  
	 * @Fields italic : 是否斜体
	 * @author caiyang
	 * @date 2023-07-12 10:25:52 
	 */  
	private boolean italic = false;
	
	/**  
	 * @Fields underLine : 是否有下划线 1是 2否
	 * @author caiyang
	 * @date 2023-07-24 08:51:18 
	 */  
	private byte underLine;

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

	public String getCellValue() {
		return cellValue;
	}

	public void setCellValue(String cellValue) {
		this.cellValue = cellValue;
	}

	public int getX() {
		return x;
	}

	public void setX(int x) {
		this.x = x;
	}

	public int getY() {
		return y;
	}

	public void setY(int y) {
		this.y = y;
	}

	public Integer getIsMerge() {
		return isMerge;
	}

	public void setIsMerge(Integer isMerge) {
		this.isMerge = isMerge;
	}

	public Integer getVerticalAlignment() {
		return verticalAlignment;
	}

	public void setVerticalAlignment(Integer verticalAlignment) {
		this.verticalAlignment = verticalAlignment;
	}

	public Integer getHorizontalAlignment() {
		return horizontalAlignment;
	}

	public void setHorizontalAlignment(Integer horizontalAlignment) {
		this.horizontalAlignment = horizontalAlignment;
	}

	public String getBackgroundColor() {
		return backgroundColor;
	}

	public void setBackgroundColor(String backgroundColor) {
		this.backgroundColor = backgroundColor;
	}

	public String getFontColor() {
		return fontColor;
	}

	public void setFontColor(String fontColor) {
		this.fontColor = fontColor;
	}

	public boolean isBold() {
		return bold;
	}

	public void setBold(boolean bold) {
		this.bold = bold;
	}

	public String getFontSize() {
		return fontSize;
	}

	public void setFontSize(String fontSize) {
		this.fontSize = fontSize;
	}

	public boolean isItalic() {
		return italic;
	}

	public void setItalic(boolean italic) {
		this.italic = italic;
	}

	public byte getUnderLine() {
		return underLine;
	}

	public void setUnderLine(byte underLine) {
		this.underLine = underLine;
	}
}
