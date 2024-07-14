package com.springreport.excel2pdf;

import java.io.InputStream;
import java.util.Map;

import com.alibaba.fastjson.JSONObject;
import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfWriter;

/**
 * Created by cary on 6/15/17.
 */
public class ExcelObject {
    /**
     * 锚名称
     */
    private String anchorName;
    /**
     * Excel Stream
     */
    private InputStream inputStream;
    /**
     * POI Excel
     */
    private Excel excel;
    
    /**  
     * @Fields startx : 起始横坐标
     * @author caiyang
     * @date 2023-04-04 02:38:28 
     */  
    private Integer startx = 0;
    
    /**  
     * @Fields starty : 起始纵坐标
     * @author caiyang
     * @date 2023-04-04 02:38:48 
     */  
    private Integer starty = 0;
    
    /**  
     * @Fields endx : 结束横坐标
     * @author caiyang
     * @date 2023-04-27 07:18:07 
     */  
    private Integer endx = 0;
    
    /**  
     * @Fields inty : 结束纵坐标
     * @author caiyang
     * @date 2023-04-27 07:36:02 
     */  
    private Integer endy = 0;
    
    /**  
     * @Fields pageType : 横向纵向 1纵向，2横向
     * @author caiyang
     * @date 2023-04-06 10:51:40 
     */  
    private Integer pageType = 1;
    
    /**  
	 * @Fields colhidden : 隐藏列
	 * @author caiyang
	 * @date 2023-04-27 05:16:08 
	 */  
	private JSONObject colhidden;
	
	/**  
	 * @Fields rowhidden : 隐藏行
	 * @author caiyang
	 * @date 2023-04-27 05:16:32 
	 */  
	private JSONObject rowhidden;
	
	private int sheetIndex = 0;
    
	/**  
	 * @Fields xxbtScreenshot : 斜线表头截图
	 * @author caiyang
	 * @date 2023-11-29 12:20:42 
	 */  
	private JSONObject xxbtScreenshot;
	
	/**  
	 * @Fields printSettings : pdf打印设置
	 * @author caiyang
	 * @date 2023-12-08 09:20:53 
	 */  
	private PrintSettingsDto printSettings;
	
	private float tableHeight;
	
	private float tableWidth;
	
	private PdfWriter writer;
	
	private Document document;
	
	/**  
	 * @Fields imageInfos : 计算位置后的图片信息
	 * @author caiyang
	 * @date 2024-07-13 09:54:54 
	 */  
	private Map<String, Map<String, Object>> imageInfos;
	
	/**  
	 * @Fields backImages : 背景图片，通过插入图片添加的图片
	 * @author caiyang
	 * @date 2024-07-13 07:55:23 
	 */  
	private Map<String, String> backImages;
	
    
    public ExcelObject(String anchorName , InputStream inputStream,Integer startx,Integer starty,Integer endx,Integer endy,Integer pageType,JSONObject colhidden,JSONObject rowhidden,JSONObject xxbtScreenshot,PrintSettingsDto printSettings,Map<String, Map<String, Object>> imageInfos,Map<String, String> backImages){
        this.anchorName = anchorName;
        this.inputStream = inputStream;
        this.excel = new Excel(this.inputStream,this.sheetIndex);
        this.startx = startx;
        this.starty = starty;
        this.pageType = pageType;
        this.endx = endx;
        this.endy = endy;
        this.rowhidden = rowhidden;
        this.colhidden = colhidden;
        this.xxbtScreenshot = xxbtScreenshot;
        this.printSettings = printSettings;
        this.imageInfos = imageInfos;
        this.backImages = backImages;
    }
    
    public ExcelObject(String anchorName , InputStream inputStream,Integer startx,Integer starty,Integer endx,Integer endy,Integer pageType,JSONObject colhidden,JSONObject rowhidden,Integer sheetIndex,JSONObject xxbtScreenshot,PrintSettingsDto printSettings,Map<String, Map<String, Object>> imageInfos,Map<String, String> backImages){
        this.anchorName = anchorName;
        this.inputStream = inputStream;
        this.sheetIndex = sheetIndex;
        this.excel = new Excel(this.inputStream,this.sheetIndex);
        this.startx = startx;
        this.starty = starty;
        this.pageType = pageType;
        this.endx = endx;
        this.endy = endy;
        this.rowhidden = rowhidden;
        this.colhidden = colhidden;
        this.xxbtScreenshot = xxbtScreenshot;
        this.printSettings = printSettings;
        this.imageInfos = imageInfos;
        this.backImages = backImages;
    }
    public String getAnchorName() {
        return anchorName;
    }
    public void setAnchorName(String anchorName) {
        this.anchorName = anchorName;
    }
    public InputStream getInputStream() {
        return this.inputStream;
    }
    public void setInputStream(InputStream inputStream) {
        this.inputStream = inputStream;
    }
    Excel getExcel() {
        return excel;
    }

	public Integer getStartx() {
		return startx;
	}

	public void setStartx(Integer startx) {
		this.startx = startx;
	}

	public Integer getStarty() {
		return starty;
	}

	public void setStarty(Integer starty) {
		this.starty = starty;
	}

	public Integer getPageType() {
		return pageType;
	}

	public void setPageType(Integer pageType) {
		this.pageType = pageType;
	}

	public Integer getEndx() {
		return endx;
	}

	public void setEndx(Integer endx) {
		this.endx = endx;
	}

	public Integer getEndy() {
		return endy;
	}

	public void setEndy(Integer endy) {
		this.endy = endy;
	}

	public JSONObject getColhidden() {
		return colhidden;
	}

	public void setColhidden(JSONObject colhidden) {
		this.colhidden = colhidden;
	}

	public JSONObject getRowhidden() {
		return rowhidden;
	}

	public void setRowhidden(JSONObject rowhidden) {
		this.rowhidden = rowhidden;
	}

	public int getSheetIndex() {
		return sheetIndex;
	}

	public void setSheetIndex(int sheetIndex) {
		this.sheetIndex = sheetIndex;
	}

	public JSONObject getXxbtScreenshot() {
		return xxbtScreenshot;
	}

	public void setXxbtScreenshot(JSONObject xxbtScreenshot) {
		this.xxbtScreenshot = xxbtScreenshot;
	}

	public void setExcel(Excel excel) {
		this.excel = excel;
	}

	public PrintSettingsDto getPrintSettings() {
		return printSettings;
	}

	public void setPrintSettings(PrintSettingsDto printSettings) {
		this.printSettings = printSettings;
	}

	public float getTableHeight() {
		return tableHeight;
	}

	public void setTableHeight(float tableHeight) {
		this.tableHeight = tableHeight;
	}

	public float getTableWidth() {
		return tableWidth;
	}

	public void setTableWidth(float tableWidth) {
		this.tableWidth = tableWidth;
	}

	public PdfWriter getWriter() {
		return writer;
	}

	public void setWriter(PdfWriter writer) {
		this.writer = writer;
	}

	public Document getDocument() {
		return document;
	}

	public void setDocument(Document document) {
		this.document = document;
	}

	public Map<String, Map<String, Object>> getImageInfos() {
		return imageInfos;
	}

	public void setImageInfos(Map<String, Map<String, Object>> imageInfos) {
		this.imageInfos = imageInfos;
	}

	public Map<String, String> getBackImages() {
		return backImages;
	}

	public void setBackImages(Map<String, String> backImages) {
		this.backImages = backImages;
	}

//	public Map<String, Integer> getMaxCoordinate() {
//		return maxCoordinate;
//	}
//
//	public void setMaxCoordinate(Map<String, Integer> maxCoordinate) {
//		this.maxCoordinate = maxCoordinate;
//	}
}