package com.springreport.excel2pdf;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.util.Date;


/**  
 * @ClassName: LuckysheetUtil
 * @Description: luckysheet工具类
 * @author caiyang
 * @date 2022-10-12 03:51:11 
*/  
public class LuckysheetUtil {

	/**  
	 * @MethodName: formatValue
	 * @Description: 格式化值
	 * @author caiyang
	 * @return 
	 * @return String
	 * @date 2022-10-12 03:53:10 
	 */  
	public static Object formatValue(String fa,Object value)
	{
		if(value == null || StringUtil.isNullOrEmpty(String.valueOf(value)))
		{
			return "";
		}
		if(fa.equals(CellFormatEnum.GENERAL.getCode()))
		{//自动直接返回
			return value;
		}else if(fa.equals(CellFormatEnum.TEXT.getCode())) {
			//纯文本
			return String.valueOf(value);
		}else if(fa.equals(CellFormatEnum.INTEGER.getCode())) {
			//整数
			if(CheckUtil.isNumber(String.valueOf(value)))
			{
				return numberTranform(value,0);
			}else {
				return value;
			}
		}else if(fa.equals(CellFormatEnum.INTEGER_2.getCode())) {
			//逗号分割整数
			if(CheckUtil.isNumber(String.valueOf(value)))
			{
				String numStr = numberTranform(value,0);
				return strAddComma(numStr);
			}else {
				return value;
			}
		}else if(fa.equals(CellFormatEnum.FLOAT1.getCode())) {
			//一位小数
			if(CheckUtil.isNumber(String.valueOf(value)))
			{
				return numberTranform(value,1);
			}else {
				return value;
			}
		}else if(fa.equals(CellFormatEnum.FLOAT2_1.getCode()) || fa.equals(CellFormatEnum.FLOAT2_2.getCode())) {
			//两位小数
			if(CheckUtil.isNumber(String.valueOf(value)))
			{
				return numberTranform(value,2);
			}else {
				return value;
			}
		}else if(fa.equals(CellFormatEnum.FLOAT2_3.getCode()) || fa.equals(CellFormatEnum.FLOAT2_4.getCode())) {
			//两位小数逗号分割
			if(CheckUtil.isNumber(String.valueOf(value)))
			{
				String numStr = numberTranform(value,2);
				return strAddComma(numStr);
			}else {
				return value;
			}
		}else if(fa.equals(CellFormatEnum.PERCENT1.getCode())) {
			//整数百分比
			if(CheckUtil.isNumber(String.valueOf(value)))
			{
				return parsePercent(value,fa);
			}else {
				return value;
			}
		}else if(fa.equals(CellFormatEnum.PERCENT2_1.getCode()) || fa.equals(CellFormatEnum.PERCENT2_2.getCode())) {
			//两位小数百分比
			if(CheckUtil.isNumber(String.valueOf(value)))
			{
				return parsePercent(value,fa.replace("#", ""));
			}else {
				return value;
			}
		}else if(fa.equals(CellFormatEnum.SCIENTIC_1.getCode()) || fa.equals(CellFormatEnum.SCIENTIC_2.getCode())) {
			//科学计数法
			if(CheckUtil.isNumber(String.valueOf(value)))
			{
				return double2Scientific((long) value);
			}else {
				return value;
			}
		}else if(fa.equals(CellFormatEnum.ACCOUNT.getCode())) {
			//会计
			if(CheckUtil.isNumber(String.valueOf(value)))
			{
				return accountFormat(value);
			}else {
				return value;
			}
		}else if(fa.equals(CellFormatEnum.WANYUAN.getCode())) {
			//万元
			if(CheckUtil.isNumber(String.valueOf(value)))
			{
				return wanyuanFormat(value);
			}else {
				return value;
			}
		}else if(fa.equals(CellFormatEnum.WANYUAN_2.getCode())) {
			//万元2位小数
			if(CheckUtil.isNumber(String.valueOf(value)))
			{
				return wanyuanFormat(value);
			}else {
				return value;
			}
		}else if(fa.equals(CellFormatEnum.CURRENCY_1.getCode()) || fa.equals(CellFormatEnum.CURRENCY_2.getCode())) {
			//货币(人民币)
			if(CheckUtil.isNumber(String.valueOf(value)))
			{
				BigDecimal bigDecimal = new BigDecimal(String.valueOf(value));
				DecimalFormat decimalFormat = new DecimalFormat(".00");
				if(bigDecimal.compareTo(new BigDecimal(0)) < 0)
				{
					return "-¥" + decimalFormat.format(Math.abs(bigDecimal.longValue()));
				}else {
					return "¥" + decimalFormat.format(bigDecimal.longValue());
				}
			}else {
				return value;
			}
		}else if(fa.equals(CellFormatEnum.DATE_1.getCode())) {
			//日期格式yyyy-MM-dd
			if(CheckUtil.isDate(String.valueOf(value), DateUtil.FORMAT_LONOGRAM))
			{
				return DateUtil.date2String(DateUtil.string2Date(String.valueOf(value), DateUtil.FORMAT_LONOGRAM), DateUtil.FORMAT_LONOGRAM);
			}else {
				if(CheckUtil.isNumeric(String.valueOf(value))) {
					 Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(String.valueOf(value)));
					 return DateUtil.date2String(date, DateUtil.FORMAT_LONOGRAM);
				}else {
					return value;
				}
			}
		}else if(fa.equals(CellFormatEnum.DATE_2.getCode())) {
			//日期格式yyyy-MM-dd hh:mm AM/PM
			if(CheckUtil.isDate(String.valueOf(value), DateUtil.FORMAT_FULL))
			{
				return DateUtil.date2String(DateUtil.string2Date(String.valueOf(value), DateUtil.FORMAT_FULL), DateUtil.FORMAT_FULL_12).replace("上午", "AM").replace("下午", "PM");
			}else {
				if(CheckUtil.isNumeric(String.valueOf(value))) {
					 Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(String.valueOf(value)));
					 return DateUtil.date2String(date, DateUtil.FORMAT_FULL);
				}else {
					return value;
				}
			}
		}else if(fa.equals(CellFormatEnum.DATE_3.getCode())) {
			//日期格式yyyy-MM-dd hh:mm
			if(CheckUtil.isDate(String.valueOf(value), DateUtil.FORMAT_FULL))
			{
				return DateUtil.date2String(DateUtil.string2Date(String.valueOf(value), DateUtil.FORMAT_WITHOUTSECONDS), DateUtil.FORMAT_WITHOUTSECONDS);
			}else {
				if(CheckUtil.isNumeric(String.valueOf(value))) {
					 Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(String.valueOf(value)));
					 return DateUtil.date2String(date, DateUtil.FORMAT_WITHOUTSECONDS);
				}else {
					return value;
				}
			}
		}else if(fa.equals(CellFormatEnum.DATE_4.getCode())) {
			if(CheckUtil.isDate(String.valueOf(value), DateUtil.FORMAT_FULL))
			{
				return DateUtil.date2String(DateUtil.string2Date(String.valueOf(value), DateUtil.FORMAT_HOURSMINUTES), DateUtil.FORMAT_HOURSMINUTES);
			}else {
				if(CheckUtil.isNumeric(String.valueOf(value))) {
					 Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(String.valueOf(value)));
					 return DateUtil.date2String(date, DateUtil.FORMAT_HOURSMINUTES).replace("上午", "AM").replace("下午", "PM");
				}else {
					return value;
				}
			}
		}else if(fa.equals(CellFormatEnum.DATE_5.getCode())) {
			if(CheckUtil.isDate(String.valueOf(value), DateUtil.FORMAT_FULL))
			{
				return DateUtil.date2String(DateUtil.string2Date(String.valueOf(value), DateUtil.FORMAT_LONOGRAM_CN_2), DateUtil.FORMAT_LONOGRAM_CN_2);
			}else {
				if(CheckUtil.isNumeric(String.valueOf(value))) {
					 Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(String.valueOf(value)));
					 return DateUtil.date2String(date, DateUtil.FORMAT_LONOGRAM_CN_2);
				}else {
					return value;
				}
			}
		}else if(fa.equals(CellFormatEnum.DATE_6.getCode())) {
			if(CheckUtil.isDate(String.valueOf(value), DateUtil.FORMAT_LONOGRAM_2))
			{
				return DateUtil.date2String(DateUtil.string2Date(String.valueOf(value), DateUtil.FORMAT_LONOGRAM_2), DateUtil.FORMAT_LONOGRAM_2);
			}else {
				if(CheckUtil.isNumeric(String.valueOf(value))) {
					 Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(String.valueOf(value)));
					 return DateUtil.date2String(date, DateUtil.FORMAT_LONOGRAM_2);
				}else {
					return value;
				}
			}
		}else if(fa.equals(CellFormatEnum.DATE_7.getCode())) {
			if(CheckUtil.isDate(String.valueOf(value), DateUtil.FORMAT_FULL))
			{
				return DateUtil.date2String(DateUtil.string2Date(String.valueOf(value), DateUtil.FORMAT_FULL), DateUtil.FORMAT_LONOGRAM_CN_2);
			}else {
				if(CheckUtil.isNumeric(String.valueOf(value))) {
					 Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(String.valueOf(value)));
					 return DateUtil.date2String(date, DateUtil.FORMAT_LONOGRAM_CN_2);
				}else {
					return value;
				}
			}
		}else if(fa.equals(CellFormatEnum.DATE_8.getCode())) {
			if(CheckUtil.isDate(String.valueOf(value), DateUtil.FORMAT_FULL))
			{
				return DateUtil.date2String(DateUtil.string2Date(String.valueOf(value), DateUtil.FORMAT_FULL), DateUtil.FORMAT_DATE);
			}else {
				if(CheckUtil.isNumeric(String.valueOf(value))) {
					 Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(String.valueOf(value)));
					 return DateUtil.date2String(date, DateUtil.FORMAT_DATE);
				}else {
					return value;
				}
			}
		}else if(fa.equals(CellFormatEnum.DATE_9.getCode())) {
			if(CheckUtil.isDate(String.valueOf(value), DateUtil.FORMAT_FULL))
			{
				return DateUtil.date2String(DateUtil.string2Date(String.valueOf(value), DateUtil.FORMAT_FULL), DateUtil.FORMAT_DATE_2);
			}else {
				if(CheckUtil.isNumeric(String.valueOf(value))) {
					 Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(String.valueOf(value)));
					 return DateUtil.date2String(date, DateUtil.FORMAT_DATE_2);
				}else {
					return value;
				}
			}
		}else if(fa.equals(CellFormatEnum.DATE_10.getCode())) {
			if(CheckUtil.isDate(String.valueOf(value), DateUtil.FORMAT_FULL))
			{
				return DateUtil.date2String(DateUtil.string2Date(String.valueOf(value), DateUtil.FORMAT_FULL), DateUtil.FORMAT_DATE_CN);
			}else {
				if(CheckUtil.isNumeric(String.valueOf(value))) {
					 Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(String.valueOf(value)));
					 return DateUtil.date2String(date, DateUtil.FORMAT_DATE_CN);
				}else {
					return value;
				}
			}
		}else if(fa.equals(CellFormatEnum.DATE_11.getCode())) {
			if(CheckUtil.isDate(String.valueOf(value), DateUtil.FORMAT_FULL))
			{
				return DateUtil.date2String(DateUtil.string2Date(String.valueOf(value), DateUtil.FORMAT_FULL), DateUtil.FORMAT_HOURSMINUTESSECONDS_2);
			}else {
				if(CheckUtil.isNumeric(String.valueOf(value))) {
					 Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(String.valueOf(value)));
					 return DateUtil.date2String(date, DateUtil.FORMAT_HOURSMINUTESSECONDS_2);
				}else {
					return value;
				}
			}
		}else if(fa.equals(CellFormatEnum.DATE_12.getCode())) {
			if(CheckUtil.isDate(String.valueOf(value), DateUtil.FORMAT_FULL))
			{
				return DateUtil.date2String(DateUtil.string2Date(String.valueOf(value), DateUtil.FORMAT_FULL), DateUtil.FORMAT_HOURSMINUTES_3);
			}else {
				if(CheckUtil.isNumeric(String.valueOf(value))) {
					 Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(Double.parseDouble(String.valueOf(value)));
					 return DateUtil.date2String(date, DateUtil.FORMAT_HOURSMINUTES_3);
				}else {
					return value;
				}
			}
		}
		
		return value;
	}
	
	/**  
	 * @MethodName: strAddComma
	 * @Description: 数字逗号分割
	 * @author caiyang
	 * @param str
	 * @return 
	 * @return String
	 * @date 2022-10-14 03:37:17 
	 */  
	private static String strAddComma(String str) {
		if (str == null) {
			str = "";
		}
		String addCommaStr = ""; // 需要添加逗号的字符串（整数）
		String tmpCommaStr = ""; // 小数，等逗号添加完后，最后在末尾补上
		if (str.contains(".")) {
			addCommaStr = str.substring(0,str.indexOf("."));
			tmpCommaStr = str.substring(str.indexOf("."),str.length());
		}else{
			addCommaStr = str;
		}
		// 将传进数字反转
		String reverseStr = new StringBuilder(addCommaStr).reverse().toString();
		String strTemp = "";
		for (int i = 0; i < reverseStr.length(); i++) {
			if (i * 3 + 3 > reverseStr.length()) {
				strTemp += reverseStr.substring(i * 3, reverseStr.length());
				break;
			}
			strTemp += reverseStr.substring(i * 3, i * 3 + 3) + ",";
		}
		// 将 "5,000,000," 中最后一个","去除
		if (strTemp.endsWith(",")) {
			strTemp = strTemp.substring(0, strTemp.length() - 1);
		}
		// 将数字重新反转,并将小数拼接到末尾
		String resultStr = new StringBuilder(strTemp).reverse().toString() + tmpCommaStr;
		return resultStr;
	}
	
	/**  
	 * @MethodName: numberTranform
	 * @Description: 数值转换成对应的格式
	 * @author caiyang
	 * @param value 数值字符串
	 * @param digit 小数位数
	 * @return 
	 * @return String
	 * @date 2022-10-12 04:28:56 
	 */  
	private static String numberTranform(Object value,int digit)
	{
		if(value == null) {
			value = 0;
		}
		BigDecimal bigDecimal = new BigDecimal(String.valueOf(value));
		bigDecimal = bigDecimal.setScale(digit, RoundingMode.HALF_UP);
		return String.valueOf(bigDecimal);
	}
	
	/**  
	 * @MethodName: parsePercent
	 * @Description: 数字转成百分比
	 * @author caiyang
	 * @param value
	 * @param digit
	 * @return 
	 * @return String
	 * @date 2022-10-12 04:41:29 
	 */  
	private static String parsePercent(Object value,String fa)
	{
		DecimalFormat decimalFormat = new DecimalFormat(fa);
		 
        // 格式化数字并输出结果
        String formattedNumber = decimalFormat.format(Double.parseDouble(value.toString()));
        return formattedNumber;
	}
	
	/**  
	 * @MethodName: double2Scientific
	 * @Description: 数字转科学计数法
	 * @author caiyang
	 * @param num
	 * @return 
	 * @return String
	 * @date 2022-10-13 08:15:52 
	 */  
	private static String double2Scientific(double num)
	{
		String str = String.format("%E", num);
		String temp = str.substring(0,str.indexOf(".")+4);
		String f = String.format("%.2f", Double.parseDouble(temp));
		str = f + str.substring(str.indexOf("E"));
		return str;
	}
	
	/**  
	 * @MethodName: accountFormat
	 * @Description: 会计格式
	 * @author caiyang
	 * @param value
	 * @return 
	 * @return String
	 * @date 2022-10-13 08:44:44 
	 */  
	private static String accountFormat(Object value) {
		String result = "";
		long l = Long.parseLong(String.valueOf(value));
		if(l >=0)
		{
			result = "¥(" + l + ")";
		}else {
			result = "-¥(" + Math.abs(l) + ")";
		}
		return result;
	}
	
	private static String wanyuanFormat(Object value)
	{
		String result = "";
		BigDecimal bigDecimal = new BigDecimal(String.valueOf(value));
		BigDecimal[] decimalResults = bigDecimal.divideAndRemainder(new BigDecimal(10000));
		if(decimalResults[0].compareTo(new BigDecimal(0)) == 0)
		{
			if(isIntegerValue(decimalResults[1]))
			{
				result = String.valueOf(decimalResults[1].intValue());
			}else {
				DecimalFormat decimalFormat = new DecimalFormat(".00");
				result = decimalFormat.format(decimalResults[1]);
			}
		}else {
			if(decimalResults[1].compareTo(new BigDecimal(0)) == 0)
			{
				result = decimalResults[0].intValue() + "万";
			}else {
				if(decimalResults[1].compareTo(new BigDecimal(1)) < 0)
				{
					DecimalFormat decimalFormat = new DecimalFormat(".00");
					result = decimalResults[0].intValue() + "万" + decimalFormat.format(Math.abs(decimalResults[1].doubleValue()));
				}else {
					if(isIntegerValue(decimalResults[1]))
					{
						DecimalFormat decimalFormat = new DecimalFormat("0000");
						result = decimalResults[0].intValue() + "万" + decimalFormat.format(Math.abs(decimalResults[1].doubleValue()));
					}else {
						DecimalFormat decimalFormat = new DecimalFormat("0000.00");
						result = decimalResults[0].intValue() + "万" + decimalFormat.format(Math.abs(decimalResults[1].doubleValue()));
					}
				}
			}
		}
		return result;
	}
	
	/**  
	 * @MethodName: isIntegerValue
	 * @Description: 校验bigdecimal是否是整数
	 * @author caiyang
	 * @param bd
	 * @return 
	 * @return boolean
	 * @date 2022-10-13 03:54:49 
	 */  
	private static boolean isIntegerValue(BigDecimal bd) {
	    boolean ret;
	    try {
	        bd.toBigIntegerExact();
	        ret = true;
	    } catch (ArithmeticException ex) {
	        ret = false;
	    }

	    return ret;
	}
}
