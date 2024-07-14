package com.springreport.excel2pdf;

import com.alibaba.fastjson.JSONArray;

/**  
 * @ClassName: PrintSettingsDto
 * @Description: pdf打印设置
 * @author caiyang
 * @date 2023-12-08 09:19:40 
*/ 
public class PrintSettingsDto {

	 /** page_type - 纸张类型 1 A3 2 A4 3 A5 4 A6 5 B2 6B3 7B4 8 B5 9LETTER 10 LEGAL */
    private Integer pageType = 2;

    /** page_layout - 纸张布局 1纵向 2横向 */
    private Integer pageLayout = 1;

    /** page_header_show - 页眉是否显示 1是 2否 */
    private Integer pageHeaderShow = 2;

    /** page_header_content - 页眉显示内容 */
    private String pageHeaderContent;

    /** page_header_position - 页眉显示位置 1左 2中 3右 */
    private Integer pageHeaderPosition;

    /** water_mark_show - 水印是否显示 1是 2否 */
    private Integer waterMarkShow = 2;

    /** water_mark_type - 水印类型 1文本 2图片 */
    private Integer waterMarkType;

    /** water_mark_content - 文本水印内容 */
    private String waterMarkContent;

    /** water_mark_img - 图片水印url */
    private String waterMarkImg;

    /** page_show - 页码是否显示 1是 2否 */
    private Integer pageShow = 2;

    /** page_position - 页码显示位置 1左 2中 3右 */
    private Integer pagePosition;
    
    /** water_mark_opacity - 水印透明度 大于0小于1的值 */
    private Float waterMarkOpacity;
    
    /**  
     * @Fields author : 导出人
     * @author caiyang
     * @date 2023-12-09 08:09:25 
     */  
    private String author;
    
    /**  
     * @Fields title : 标题（报表模板名称）
     * @author caiyang
     * @date 2023-12-09 08:09:50 
     */  
    private String title;
    
    /**  
     * @Fields subject : 主题（sheet页名称）
     * @author caiyang
     * @date 2023-12-09 08:10:37 
     */  
    private String subject;
    
    /**  
     * @Fields keyWords : 关键字
     * @author caiyang
     * @date 2023-12-09 08:11:17 
     */  
    private String keyWords = "SpringReport";
    
    /**  
     * @Fields horizontalPage : 是否水平分页 1是 2否 默认2
     * @author caiyang
     * @date 2024-02-26 05:02:05 
     */  
    private Integer horizontalPage = 2;
    
    /**  
	 * @Fields pageDivider : 分页线配置
	 * @author caiyang
	 * @date 2024-02-26 04:45:47 
	 */  
	private JSONArray pageDivider;

	public Integer getPageType() {
		return pageType;
	}

	public void setPageType(Integer pageType) {
		this.pageType = pageType;
	}

	public Integer getPageLayout() {
		return pageLayout;
	}

	public void setPageLayout(Integer pageLayout) {
		this.pageLayout = pageLayout;
	}

	public Integer getPageHeaderShow() {
		return pageHeaderShow;
	}

	public void setPageHeaderShow(Integer pageHeaderShow) {
		this.pageHeaderShow = pageHeaderShow;
	}

	public String getPageHeaderContent() {
		return pageHeaderContent;
	}

	public void setPageHeaderContent(String pageHeaderContent) {
		this.pageHeaderContent = pageHeaderContent;
	}

	public Integer getPageHeaderPosition() {
		return pageHeaderPosition;
	}

	public void setPageHeaderPosition(Integer pageHeaderPosition) {
		this.pageHeaderPosition = pageHeaderPosition;
	}

	public Integer getWaterMarkShow() {
		return waterMarkShow;
	}

	public void setWaterMarkShow(Integer waterMarkShow) {
		this.waterMarkShow = waterMarkShow;
	}

	public Integer getWaterMarkType() {
		return waterMarkType;
	}

	public void setWaterMarkType(Integer waterMarkType) {
		this.waterMarkType = waterMarkType;
	}

	public String getWaterMarkContent() {
		return waterMarkContent;
	}

	public void setWaterMarkContent(String waterMarkContent) {
		this.waterMarkContent = waterMarkContent;
	}

	public String getWaterMarkImg() {
		return waterMarkImg;
	}

	public void setWaterMarkImg(String waterMarkImg) {
		this.waterMarkImg = waterMarkImg;
	}

	public Integer getPageShow() {
		return pageShow;
	}

	public void setPageShow(Integer pageShow) {
		this.pageShow = pageShow;
	}

	public Integer getPagePosition() {
		return pagePosition;
	}

	public void setPagePosition(Integer pagePosition) {
		this.pagePosition = pagePosition;
	}

	public String getAuthor() {
		return author;
	}

	public void setAuthor(String author) {
		this.author = author;
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public String getSubject() {
		return subject;
	}

	public void setSubject(String subject) {
		this.subject = subject;
	}

	public String getKeyWords() {
		return keyWords;
	}

	public void setKeyWords(String keyWords) {
		this.keyWords = keyWords;
	}

	public Float getWaterMarkOpacity() {
		return waterMarkOpacity;
	}

	public void setWaterMarkOpacity(Float waterMarkOpacity) {
		this.waterMarkOpacity = waterMarkOpacity;
	}

	public Integer getHorizontalPage() {
		return horizontalPage;
	}

	public void setHorizontalPage(Integer horizontalPage) {
		this.horizontalPage = horizontalPage;
	}

	public JSONArray getPageDivider() {
		return pageDivider;
	}

	public void setPageDivider(JSONArray pageDivider) {
		this.pageDivider = pageDivider;
	}
}
