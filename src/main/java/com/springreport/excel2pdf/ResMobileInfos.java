package com.springreport.excel2pdf;

import java.util.List;
import java.util.Set;


/**  
 * @ClassName: ResMobileInfos
 * @Description: 手机端信息返回结果用实体类
 * @author caiyang
 * @date 2023-06-28 08:19:38 
*/ 
public class ResMobileInfos {
	/**  
	 * @Fields tableCells : 单元格数据
	 * @author caiyang
	 * @date 2023-06-28 08:23:25 
	 */  
	List<List<TableCell>> tableCells;
	
	/**  
	 * @Fields imageInfos : 图片信息
	 * @author caiyang
	 * @date 2023-06-28 08:39:52 
	 */  
	List<ImageInfo> imageInfos;
	
	/**  
	 * @Fields mergeCells : 合并单元格
	 * @author caiyang
	 * @date 2023-07-10 07:41:17 
	 */  
	private Set<String> mergeCells;
	
	/**  
	 * @Fields emptyRows : 空行
	 * @author caiyang
	 * @date 2023-07-12 09:08:57 
	 */  
	private Set<Integer> emptyRows;
	
	public List<List<TableCell>> getTableCells() {
		return tableCells;
	}

	public void setTableCells(List<List<TableCell>> tableCells) {
		this.tableCells = tableCells;
	}

	public List<ImageInfo> getImageInfos() {
		return imageInfos;
	}

	public void setImageInfos(List<ImageInfo> imageInfos) {
		this.imageInfos = imageInfos;
	}

	public Set<String> getMergeCells() {
		return mergeCells;
	}

	public void setMergeCells(Set<String> mergeCells) {
		this.mergeCells = mergeCells;
	}

	public Set<Integer> getEmptyRows() {
		return emptyRows;
	}

	public void setEmptyRows(Set<Integer> emptyRows) {
		this.emptyRows = emptyRows;
	}
}
