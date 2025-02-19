package com.springreport.excel2pdf;

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.BaseFont;

import cn.hutool.core.io.resource.ClassPathResource;

/**
 * Created by cary on 6/15/17.
 */
public class Resource {
    /**
     * 中文字体支持
     */
    protected static BaseFont BASE_FONT_CHINESE;
    
    protected static Map<String, BaseFont> baseFontMap = new HashMap<String, BaseFont>();
    static {
        try {
            BASE_FONT_CHINESE = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
            // 搜尋系統,載入系統內的字型(慢)
            FontFactory.registerDirectories();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    static {
    	try {
    		Map<String, String> fontName = new HashMap<>();
//    		fontName.put("宋体", "simsun.ttc");
    		fontName.put("黑体", "simhei.ttf");
    		fontName.put("楷体", "simkai.ttf");
    		fontName.put("仿宋", "simfang.ttf");
//    		fontName.put("新宋体", "simsun.ttc");
    		fontName.put("华文新魏", "STXINWEI.TTF");
    		fontName.put("华文行楷", "STXINGKA.TTF");
    		fontName.put("华文隶书", "STLITI.TTF");
    		fontName.forEach((key,value) -> {
    			byte[] bytes = null;
				try {
					bytes = getResourceAsStream("classpath:/font/"+value);
				} catch (Exception e) {
					e.printStackTrace();
				}
    			try {
					BaseFont baseFont = BaseFont.createFont(value, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED,true,bytes,null);
					baseFontMap.put(key, baseFont);
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} 
            });
		} catch (Exception e) {
			// TODO: handle exception
		}
    }
    
    private static byte[] getResourceAsStream(String path) throws IOException {
        ClassPathResource resource = new ClassPathResource(path);
        return resource.readBytes();
    }

    /**
     * 將 POI Font 轉換到 iText Font
     * @param font
     * @return
     */
    public static com.itextpdf.text.Font getFont(HSSFFont font) {
        try {
            com.itextpdf.text.Font iTextFont = FontFactory.getFont(font.getFontName(),
                    BaseFont.IDENTITY_H, BaseFont.EMBEDDED,
                    font.getFontHeightInPoints());
            return iTextFont;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
    
    public static BaseFont getFont(String fontName) {
    	BaseFont result = baseFontMap.get(fontName);
    	if(result == null) {
    		result = BASE_FONT_CHINESE;
    	}
    	return result;
    }
}