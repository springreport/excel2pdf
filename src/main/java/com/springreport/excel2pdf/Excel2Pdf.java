package com.springreport.excel2pdf;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;

import java.io.IOException;
import java.io.OutputStream;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.List;


/**
 * Created by cary on 6/15/17.
 */
public class Excel2Pdf extends PdfTool{
    protected List<ExcelObject> objects = new ArrayList<ExcelObject>();

    /**
     * <p>Description: 导出单项PDF，不包含目录</p>
     * @param object
     */
    public Excel2Pdf(ExcelObject object , OutputStream os) {
        this.objects.add(object);
        this.os = os;
    }

    /**
     * <p>Description: 导出多项PDF，包含目录</p>
     * @param objects
     */
    public Excel2Pdf(List<ExcelObject> objects , OutputStream os) {
        this.objects = objects;
        this.os = os;
    }

    /**
     * <p>Description: 转换调用</p>
     * @throws DocumentException
     * @throws MalformedURLException
     * @throws IOException
     */
    public void convert() throws DocumentException, MalformedURLException, IOException {
    	if(this.objects.get(0).getPrintSettings() != null)
    	{
    		switch (this.objects.get(0).getPrintSettings().getPageType().intValue()) {
			case 1:
				if(this.objects.get(0).getPrintSettings().getPageLayout().intValue() == 1)
				{
					getDocument().setPageSize(PageSize.A3);//纵向pdf
				}else {
					getDocument().setPageSize(PageSize.A3.rotate());//横向pdf
				}
				break;
			case 2:
				if(this.objects.get(0).getPrintSettings().getPageLayout().intValue() == 1)
				{
					getDocument().setPageSize(PageSize.A4);//纵向pdf
				}else {
					getDocument().setPageSize(PageSize.A4.rotate());//横向pdf
				}
				break;
			case 3:
				if(this.objects.get(0).getPrintSettings().getPageLayout().intValue() == 1)
				{
					getDocument().setPageSize(PageSize.A5);//纵向pdf
				}else {
					getDocument().setPageSize(PageSize.A5.rotate());//横向pdf
				}
				break;
			case 4:
				if(this.objects.get(0).getPrintSettings().getPageLayout().intValue() == 1)
				{
					getDocument().setPageSize(PageSize.A6);//纵向pdf
				}else {
					getDocument().setPageSize(PageSize.A6.rotate());//横向pdf
				}
				break;
			case 5:
				if(this.objects.get(0).getPrintSettings().getPageLayout().intValue() == 1)
				{
					getDocument().setPageSize(PageSize.B2);//纵向pdf
				}else {
					getDocument().setPageSize(PageSize.B2.rotate());//横向pdf
				}
				break;
			case 6:
				if(this.objects.get(0).getPrintSettings().getPageLayout().intValue() == 1)
				{
					getDocument().setPageSize(PageSize.B3);//纵向pdf
				}else {
					getDocument().setPageSize(PageSize.B3.rotate());//横向pdf
				}
				break;
			case 7:
				if(this.objects.get(0).getPrintSettings().getPageLayout().intValue() == 1)
				{
					getDocument().setPageSize(PageSize.B4);//纵向pdf
				}else {
					getDocument().setPageSize(PageSize.B4.rotate());//横向pdf
				}
				break;
			case 8:
				if(this.objects.get(0).getPrintSettings().getPageLayout().intValue() == 1)
				{
					getDocument().setPageSize(PageSize.B5);//纵向pdf
				}else {
					getDocument().setPageSize(PageSize.B5.rotate());//横向pdf
				}
				break;
			case 9:
				if(this.objects.get(0).getPrintSettings().getPageLayout().intValue() == 1)
				{
					getDocument().setPageSize(PageSize.LETTER);//纵向pdf
				}else {
					getDocument().setPageSize(PageSize.LETTER.rotate());//横向pdf
				}
				break;
			case 10:
				if(this.objects.get(0).getPrintSettings().getPageLayout().intValue() == 1)
				{
					getDocument().setPageSize(PageSize.LEGAL);//纵向pdf
				}else {
					getDocument().setPageSize(PageSize.LEGAL.rotate());//横向pdf
				}
				break;
			default:
				getDocument().setPageSize(PageSize.A4);//纵向pdf
				break;
			}
    		if(this.objects.get(0).getPrintSettings().getCustomMargin().intValue() == 1) {
    			getDocument().setMargins(this.objects.get(0).getPrintSettings().getLeftMargin(), this.objects.get(0).getPrintSettings().getRightMargin(),
    					this.objects.get(0).getPrintSettings().getTopMargin(), this.objects.get(0).getPrintSettings().getBottomMargin());
    		}
    	}else {
    		getDocument().setPageSize(PageSize.A4);//纵向pdf
    	}
        PdfWriter writer = PdfWriter.getInstance(getDocument(), os);
        writer.setPageEvent(new PDFPageEvent(this.objects.get(0).getPrintSettings()));
        this.objects.get(0).setWriter(writer);
        //Open document
        getDocument().open();
        //Single one
        if(this.objects.size() <= 1){
        	List<PdfPTable> tables = this.toCreatePdfTable(this.objects.get(0) ,  getDocument());
        	for (int i = 0; i < tables.size(); i++) {
        		getDocument().add(tables.get(i));
        		getDocument().newPage();
			}
//            
        }
        //Multiple ones
//        if(this.objects.size() > 1){
//            toCreateContentIndexes(writer , this.getDocument() , this.objects);
//            //
//            for (int i = 0; i < this.objects.size(); i++) {
//                PdfPTable table = this.toCreatePdfTable(this.objects.get(i) , getDocument() , writer);
//                getDocument().add(table);
//            }
//        }
        getDocument().close();
    }

    protected List<PdfPTable> toCreatePdfTable(ExcelObject object , Document document) throws MalformedURLException, IOException, DocumentException{
    	document.addAuthor(object.getPrintSettings().getAuthor()==null?"SpringReport":object.getPrintSettings().getAuthor());
    	document.addTitle(object.getPrintSettings().getTitle() == null?"":object.getPrintSettings().getTitle());
    	document.addSubject(object.getPrintSettings().getSubject() == null?"":object.getPrintSettings().getSubject());
    	document.addKeywords(object.getPrintSettings().getKeyWords()==null?"":object.getPrintSettings().getKeyWords());
    	float tableHeight = document.getPageSize().getTop() - document.topMargin() - document.bottomMargin();
    	object.setTableHeight(tableHeight);
    	float taleWidth = document.getPageSize().getRight() - document.leftMargin() - document.rightMargin();
    	object.setTableWidth(taleWidth);
    	object.setDocument(document);
    	List<PdfPTable> table = new PdfTableExcel(object).getTable();
        return table;
    }

    /**
     * <p>Description: 内容索引创建</p>
     * @throws DocumentException
     */
    protected void toCreateContentIndexes(PdfWriter writer , Document document , List<ExcelObject> objects) throws DocumentException{
        PdfPTable table = new PdfPTable(1);
        table.setKeepTogether(true);
        table.getDefaultCell().setBorder(PdfPCell.NO_BORDER);
        //
        Font font = new Font(Resource.BASE_FONT_CHINESE , 12 , Font.NORMAL);
        font.setColor(new BaseColor(0,0,255));
        //
        for (int i = 0; i < objects.size(); i++) {
            ExcelObject o = objects.get(i);
            String text = o.getAnchorName();
            Anchor anchor = new Anchor(text , font);
            anchor.setReference("#" + o.getAnchorName());
            //
            PdfPCell cell = new PdfPCell(anchor);
            cell.setBorder(0);
            //
            table.addCell(cell);
        }
        //
        document.add(table);
    }

    /**
     * <p>ClassName: PDFPageEvent</p>
     * <p>Description: 事件 -> 页码控制</p>
     * <p>Author: Cary</p>
     * <p>Date: Oct 25, 2013</p>
     */
    private static class PDFPageEvent extends PdfPageEventHelper {
		protected PdfTemplate template;
        public BaseFont baseFont;
        private PrintSettingsDto printSettings;
        private com.itextpdf.text.Image img;
        public PDFPageEvent(PrintSettingsDto printSettings) {
			this.printSettings = printSettings;
		}

        @Override
        public void onStartPage(PdfWriter writer, Document document) {
            try{
                this.template = writer.getDirectContent().createTemplate(100, 100);
                this.baseFont = new Font(Resource.BASE_FONT_CHINESE , 10, Font.NORMAL).getBaseFont();
            } catch(Exception e) {
                throw new ExceptionConverter(e);
            }
        }
        
        @Override
        public void onEndPage(PdfWriter writer, Document document) {
            //在每页结束的时候把“第x页”信息写道模版指定位置
            PdfContentByte byteContent = writer.getDirectContent();
            float documentHeight = document.getPageSize().getHeight();
            float documentWidth = document.getPageSize().getWidth();
            if(printSettings == null)
            {
            	printSettings = new PrintSettingsDto();
            	printSettings.setPageType(2);
            	printSettings.setPageLayout(1);
            	printSettings.setPageHeaderShow(2);
            	printSettings.setWaterMarkShow(2);
            	printSettings.setPageShow(2);
            }
            if(printSettings.getPageHeaderShow().intValue() == 1)
            {
            	if(printSettings.getPageHeaderPosition().intValue() == 1)
            	{
            		ColumnText.showTextAligned(writer.getDirectContent(), Element.ALIGN_LEFT, new Phrase(printSettings.getPageHeaderContent(), new Font(this.baseFont, 10, Font.NORMAL)), document.left(), document.top() + document.topMargin()/2, 0);
            	}else if(printSettings.getPageHeaderPosition().intValue() == 2) {
            		ColumnText.showTextAligned(writer.getDirectContent(), Element.ALIGN_CENTER, new Phrase(printSettings.getPageHeaderContent(), new Font(this.baseFont, 10, Font.NORMAL)), documentWidth/2, document.top() + document.topMargin()/2, 0);
            	}else if(printSettings.getPageHeaderPosition().intValue() == 3) {
            		ColumnText.showTextAligned(writer.getDirectContent(), Element.ALIGN_RIGHT, new Phrase(printSettings.getPageHeaderContent(), new Font(this.baseFont, 10, Font.NORMAL)), document.right(), document.top()+document.topMargin()/2, 0);
            	}
            }
            byteContent.saveState();
            if(printSettings.getPageShow().intValue() == 1)
            {
            	String text = (writer.getPageNumber()+printSettings.getStartPage()-1) + "";
            	if(printSettings.getPagePosition().intValue() == 1)
            	{
            		ColumnText.showTextAligned(writer.getDirectContent(), Element.ALIGN_LEFT, new Phrase(text, new Font(this.baseFont, 10, Font.NORMAL)), document.left(), document.bottomMargin()/2, 0);
            	}else if(printSettings.getPagePosition().intValue() == 2) {
            		ColumnText.showTextAligned(writer.getDirectContent(), Element.ALIGN_CENTER, new Phrase(text, new Font(this.baseFont, 10, Font.NORMAL)), documentWidth/2, document.bottomMargin()/2, 0);
            	}else if(printSettings.getPagePosition().intValue() == 3) {
            		ColumnText.showTextAligned(writer.getDirectContent(), Element.ALIGN_RIGHT, new Phrase(text, new Font(this.baseFont, 10, Font.NORMAL)), document.right(), document.bottomMargin()/2, 0);
            	}
            }
            if(printSettings.getWaterMarkShow().intValue() == 1)
            {
            	if(printSettings.getWaterMarkType().intValue() == 1)
            	{
            		try {
                		PdfGState gs = new PdfGState();
                		// 设置填充字体不透明度为0.4f
                        gs.setFillOpacity(printSettings.getWaterMarkOpacity());
                        byteContent.setGState(gs);
    					BaseFont base = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", BaseFont.EMBEDDED);
    					byteContent.beginText();
    	            	byteContent.setFontAndSize(base, 36f);
    	            	byteContent.setColorFill(BaseColor.LIGHT_GRAY);
    	            	byteContent.showTextAligned(Element.ALIGN_CENTER, printSettings.getWaterMarkContent(), documentWidth/2, documentHeight/2, 45);
    	            	byteContent.endText();
    				} catch (Exception e) {
    					e.printStackTrace();
    				} 
            	}else {
            		try {
            			PdfGState gs = new PdfGState();
                		// 设置填充字体不透明度为0.4f
                        gs.setFillOpacity(printSettings.getWaterMarkOpacity());
                        byteContent.setGState(gs);
                        if(this.img == null)
                        {
                        	this.img = Image.getInstance(printSettings.getWaterMarkImg());
                        }
                        float imgWidth = this.img.getWidth();
                        this.img.setAbsolutePosition((documentWidth/2-imgWidth/2)>0?documentWidth/2-imgWidth/2:document.leftMargin(), (documentHeight/2-imgWidth/2)>0?documentHeight/2-imgWidth/2:document.bottomMargin());
                        this.img.setRotationDegrees(45);
						byteContent.addImage(this.img);
						byteContent.setColorFill(BaseColor.LIGHT_GRAY);
					} catch (Exception e) {
    					e.printStackTrace();
    				} 
            	}
            }
            byteContent.restoreState();
        }
        
        public void onCloseDocument(PdfWriter writer, Document document) {
        	
        }
    }
    /**  
     * @MethodName: getMobileInfos
     * @Description: 获取移动端信息
     * @author caiyang
     * @param object
     * @return
     * @throws BadElementException
     * @throws MalformedURLException
     * @throws IOException ResMobileInfos
     * @date 2023-06-28 09:14:22 
     */ 
    public ResMobileInfos getMobileInfos() throws BadElementException, MalformedURLException, IOException{
    	ResMobileInfos mobileInfos = new PdfTableExcel(this.objects.get(0)).getMobileInfos();
    	return mobileInfos;
    }
}