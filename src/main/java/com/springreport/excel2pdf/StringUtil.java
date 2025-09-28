package com.springreport.excel2pdf;


/**  
 * @ClassName: StringUtil
 * @Description: 工具类
 * @author caiyang
 * @date 2020-05-29 11:25:12 
*/  
public class StringUtil {
	
	/**
	 * 判断字符串是否为空，null,空字符串，空格字符串都是返回true
	 * 
	 * @param str
	 * @return 是空，返回true，否则false
	 */
	public static boolean isNullOrEmpty(String str) {
		if (str == null || str.trim().length() == 0) {
			return true;
		}
		return false;
	}

	/**
	 * 判断字符串是否不为空，null,空字符串，空格字符串都是返回false
	 * 
	 * @param str
	 * @return
	 */
	public static boolean isNotEmpty(String str) {
		if (str == null || str.trim().length() == 0) {
			return false;
		}
		return true;
	}
	
	/**  
     * @MethodName: countChineseCharaceters
     * @Description: 统计字符串中中文字符的数量
     * @author caiyang
     * @param Str
     * @return boolean
     * @date 2025-07-24 09:28:26 
     */ 
    public static int countChineseCharaceters(String str) {
    	int count = 0;
    	if(StringUtil.isNullOrEmpty(str)) {
    		return 0;
    	}
    	char[] c = str.toCharArray();
        for(int i = 0; i < c.length; i ++)
        {
            String len = Integer.toBinaryString(c[i]);
            if(len.length() > 8)
                count ++;
        }
        return count;
    }
    
    /**  
     * @MethodName: countDigits
     * @Description: 统计字符串中数字的数量
     * @author caiyang
     * @param str
     * @return int
     * @date 2025-09-23 09:19:57 
     */ 
    public static int countDigits(String str) {
    	String digitsOnly = str.replaceAll("\\D", "");
        int count = digitsOnly.length();
        return count;
    }
    
    public static void main(String[] args) {
		String str = "测试123 2025-09-06 abc";
		System.err.println(StringUtil.countDigits(str));
	}
}