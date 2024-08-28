package com.springreport.excel2pdf;

import java.io.ByteArrayOutputStream;
import java.util.HashMap;
import java.util.Map;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.client.j2se.MatrixToImageConfig;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.QRCodeWriter;

public class QRCodeUtil {

	 public static byte[] generateQRCodeImage(String content, int width, int height) {
		    QRCodeWriter  qrcodeWriter = new QRCodeWriter();
	        BitMatrix bitMatrix = null;
	        Map<EncodeHintType, Object> hintMap = new HashMap<>();
	        hintMap.put(EncodeHintType.CHARACTER_SET, "UTF-8");
	        hintMap.put(EncodeHintType.MARGIN, 1);
	        try {
	            bitMatrix = qrcodeWriter.encode(content, BarcodeFormat.QR_CODE, width, height,hintMap);
	            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
	            MatrixToImageConfig config = new MatrixToImageConfig(MatrixToImageConfig.BLACK, MatrixToImageConfig.WHITE);
	            MatrixToImageWriter.writeToStream(bitMatrix, "PNG", outputStream, config);
	            return outputStream.toByteArray();
	        } catch (Exception e) {
	            e.printStackTrace();
	            return null;
	        }
	    }
}
