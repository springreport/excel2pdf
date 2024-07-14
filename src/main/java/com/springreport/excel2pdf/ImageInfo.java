package com.springreport.excel2pdf;

public class ImageInfo {

	/**  
	 * @Fields width : 宽度
	 * @author caiyang
	 * @date 2023-06-28 08:35:41 
	 */  
	private float width;
	
	/**  
	 * @Fields height : 高度
	 * @author caiyang
	 * @date 2023-06-28 08:35:57 
	 */  
	private float height;
	
	/**  
	 * @Fields bytes : 图片信息 base64编码
	 * @author caiyang
	 * @date 2023-06-28 08:36:29 
	 */  
	private String image;

	public float getWidth() {
		return width;
	}

	public void setWidth(float width) {
		this.width = width;
	}

	public float getHeight() {
		return height;
	}

	public void setHeight(float height) {
		this.height = height;
	}

	public String getImage() {
		return image;
	}

	public void setImage(String image) {
		this.image = image;
	}

}
