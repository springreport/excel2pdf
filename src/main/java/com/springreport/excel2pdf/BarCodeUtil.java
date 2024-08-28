package com.springreport.excel2pdf;
import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.client.j2se.MatrixToImageConfig;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.oned.Code128Writer;

import java.io.ByteArrayOutputStream;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;
public class BarCodeUtil {

    public static byte[] generateBarcodeImage(String content, int width, int height) {
        Code128Writer barcodeWriter = new Code128Writer();
        BitMatrix bitMatrix = null;
        Map<EncodeHintType, Object> hintMap = new HashMap<>();
        hintMap.put(EncodeHintType.CHARACTER_SET, "UTF-8");
        hintMap.put(EncodeHintType.MARGIN, 1);
        try {
            bitMatrix = barcodeWriter.encode(content, BarcodeFormat.CODE_128, width, height,hintMap);
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
