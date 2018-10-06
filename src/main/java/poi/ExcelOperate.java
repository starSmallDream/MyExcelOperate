package poi;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.imageio.ImageIO;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;

/**
 * Excel 操作工具类
 * @author xhc
 *
 */
public class ExcelOperate {

	private static Map<String,Integer> imageExtend=new HashMap<String, Integer>();
	
	static {
		imageExtend.put("EMF", Workbook.PICTURE_TYPE_EMF);
		imageExtend.put("DIB", Workbook.PICTURE_TYPE_DIB);
		imageExtend.put("JPEG", Workbook.PICTURE_TYPE_JPEG);
		imageExtend.put("PICT", Workbook.PICTURE_TYPE_PICT);
		imageExtend.put("PNG", Workbook.PICTURE_TYPE_PNG);
		imageExtend.put("WMF", Workbook.PICTURE_TYPE_WMF);
	}
	
	public final static String XLS="XLS";
	
	public final static String XLSX="XLSX";
	
	private Workbook workbook;
	
	private Sheet sheet;
	
	private Workbook targetWorkBook;
	
	private Sheet targetSheet;
	
	private CellStyle newCellStyle ;
	
	private Map<String,CellStyle> cellStyleMap=new HashMap<String, CellStyle>();
	
	private Drawing drawing;
	
	/**
	 * 	枚举Excel文件的类型
	 * @author xhc
	 *
	 */
	public static enum ExcelSuffix{
		XLS("XLS"),XLSX("XLSX");
		private String suffix;
		public String getSuffix() {
			return suffix;
		}
		ExcelSuffix(String suffix){
			this.suffix=suffix;
		}
	}
	
	public ExcelOperate(String modelPath,Workbook targetWorkBook,ExcelSuffix suffix) {
		try {
			init(modelPath,targetWorkBook,suffix);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public Sheet getModuleSheet() {
		return this.sheet;
	}
	
	public Sheet getTargetSheet() {
		return this.targetSheet;
	}
	
	private void init(String modelPath,Workbook targetWorkBook,ExcelSuffix suffix) throws FileNotFoundException, IOException {
		if(suffix.getSuffix().equals("XLS")) {
			this.workbook=new HSSFWorkbook(POIFSFileSystem.createNonClosingInputStream(new FileInputStream(modelPath)));
			this.targetWorkBook=targetWorkBook;
		}else if(suffix.getSuffix().equals("XLSX")) {
			this.workbook=new XSSFWorkbook(POIFSFileSystem.createNonClosingInputStream(new FileInputStream(modelPath)));
			this.targetWorkBook=targetWorkBook;
			
			//复制样式
			StylesTable stylesSource = ((XSSFWorkbook)(this.workbook)).getStylesSource();
			StylesTable targetStylesSource = ((XSSFWorkbook)(this.targetWorkBook)).getStylesSource();
			
			List<XSSFCellBorder> borders = stylesSource.getBorders();
			targetStylesSource.getBorders().clear();
			for(XSSFCellBorder border:borders) {
				targetStylesSource.putBorder(border);
			}
			
			List<XSSFCellFill> fills = stylesSource.getFills();
			targetStylesSource.getFills().clear();
			for(XSSFCellFill xssfCellFill:fills) {
				targetStylesSource.putFill(xssfCellFill);
			}
			
			/**
			 * 未用到
			 */
//			int _getXfsSize = stylesSource._getXfsSize();
//			for(int i=0;i<_getXfsSize;i++) {
//				targetStylesSource.putCellXf(stylesSource.getCellXfAt(i));
//			}
//			
//			int _getNumberFormatSize = stylesSource._getNumberFormatSize();
//			for(int i=0;i<_getNumberFormatSize;i++) {
//				targetStylesSource.putNumberFormat(stylesSource.getNumberFormatAt(i));
//			}
//			
//			List<XSSFFont> fonts = stylesSource.getFonts();
//			targetStylesSource.getFonts().clear();
//			for(XSSFFont font:fonts) {
//				targetStylesSource.putFont(font);
//			}
//			
//			int numCellStyles = stylesSource._getStyleXfsSize();
//			for(int i=0;i<numCellStyles;i++) {
//				targetStylesSource.putCellStyleXf(stylesSource.getCellStyleXfAt(i));
//			}
			
		}

		this.sheet=workbook.getSheetAt(0);
		this.targetSheet=targetWorkBook.createSheet();
	}
	
	
	/**
	 * 目标Sheet追加复制行 到指定的 targetSheet中，累加行为
	 * @param originRow
	 */
	public void appendCopyRow(int originRow) {
		appendCopyRow(originRow,null);
	}
	
	/**
	 * 目标Sheet追加复制行 到指定的 targetSheet中，累加行为
	 * @param originRow
	 * @param data
	 */
	public void appendCopyRow(int originRow,Map<String,Object> data) {
		int[] coordinate = getRowContainsMerged(originRow);
		
		//获取目标单元格最后一行(基准新行)
		int lastRowNum = targetSheet.getLastRowNum();
		if(lastRowNum>0) lastRowNum++;
		//循环行的范围，进行复制
		for(int rowY=coordinate[1];rowY<=coordinate[3];rowY++,lastRowNum++) {
			Row row = sheet.getRow(rowY);
			if(row==null) continue;
			short firstCellNum = row.getFirstCellNum();
			short lastCellNum = row.getLastCellNum();
			for(int colX=firstCellNum;colX<lastCellNum;colX++) {
				//判断指定行是否已经包含合并单元格了
				if(!isContainsMergedCell(targetSheet,lastRowNum,colX)) {
					//判断此单元格是否是合并单元格
					if(isMergedRegion(rowY, colX)) {
						//获取指定范围的合并单元格的索引
						int mergedRegionIndex = getMergedRegionIndex(this.sheet,rowY, colX);
						if(mergedRegionIndex==-1) continue;
						CellRangeAddress mergedRegion = sheet.getMergedRegion(mergedRegionIndex);
						//添加合并单元格
						int tFirstRow = lastRowNum;
						int tLastRow = (mergedRegion.getLastRow()-mergedRegion.getFirstRow())+lastRowNum;
						int tFirstColumn = mergedRegion.getFirstColumn();
						int tLastColumn = mergedRegion.getLastColumn();
						Row tRow =targetSheet.getRow(tFirstRow); 
						if(tRow==null) tRow=targetSheet.createRow(tFirstRow);
						Cell tCell = tRow.createCell(tFirstColumn);
						Cell firstCell = getCell(mergedRegion.getFirstRow(),mergedRegion.getFirstColumn());
						//设置单元格的值
						String cellValue = getCellValue(firstCell);
						String[] allKey = getAllKey(cellValue);
						Object value=null;
						for(String str:allKey) {
							String key=getPattermKey(str);
							Object object = data.get(key);
							if((allKey.length>1 || cellValue.length()>str.length()) && (value==null || value instanceof String)) {
								if(value==null) {
									value=cellValue;
								}
								value=value.toString().replace(str, object.toString());
							}else {
								value=object;
							}
						}
						if(data!=null && value!=null) {
							setCellValue(tCell, value);
						}else{
							setCellValue(tCell,firstCell );
						}
						CellRangeAddress cellRangeAddress = new CellRangeAddress(tFirstRow,tLastRow,tFirstColumn,tLastColumn);
						targetSheet.addMergedRegion(cellRangeAddress);
						
						//设置行高
						tRow.setHeight(row.getHeight());
						tRow.setHeightInPoints(row.getHeightInPoints());
						setRegionStyle(mergedRegion,cellRangeAddress);
					}else {
						Cell cell = getCell(rowY, colX);
						String cellValue = getCellValue(cell);
						String value=getPattermKey(cellValue);
						Row tRow = targetSheet.getRow(lastRowNum);
						if(tRow==null) tRow=targetSheet.createRow(lastRowNum);
						Cell tCell = tRow.getCell(colX);
						if(tCell==null) tCell = tRow.createCell(colX);
						getOrPutCellStyleMap(cell.getCellStyle());
						tCell.setCellStyle(newCellStyle);
						if(data!=null && value!=null) {
							setCellValue(tCell,data.get(value) );
						}else if(!"".equals(value)){//等于""或null就没有必要再进行设置单元格的值了
							setCellValue(tCell,cell );
						}
						//设置行高
						tRow.setHeight(row.getHeight());
						tRow.setHeightInPoints(row.getHeightInPoints());
						//设置列宽
						targetSheet.setColumnWidth(colX, sheet.getColumnWidth(colX));
					}
				}
			}
		}
	}
	
	/**
	 * 复制行 到指定的 targetSheet中，累加行为
	 * @param originRow
	 * @param targetRow
	 */
	public void copyRow(int originRow,int targetRow) {
		copyRow(originRow,targetRow,null);
	}
	
	
	/**
	 * 复制行 到指定的 targetSheet中，累加行为
	 * @param originRow
	 * @param targetRow
	 * @param data
	 */
	public void copyRow(int originRow,int targetRow,Map<String,Object> data) {
		int[] coordinate = getRowContainsMerged(originRow);
		//如果目标Sheet已经包含了合并单元格，则进行忽略return
		if(isRowContainsMerged(targetSheet,targetRow)) return;
		
		//判断是否可以容纳合并的行数,如果目标行为null，容纳不下，则进行移动出空闲位置
		int firstRow=0;
		int offsetRow=0;
		for(int i=coordinate[1],tRow=targetRow;i<=coordinate[3];i++,tRow++) {
			Row row = targetSheet.getRow(tRow);
			if(row!=null) {
				if(firstRow==0)
					firstRow=tRow;
				offsetRow++;
			}
		}
		if(firstRow!=0) {
			targetSheet.shiftRows(firstRow, targetSheet.getLastRowNum(), offsetRow);
		}
		
		//循环行的范围，进行复制
		for(int rowY=coordinate[1];rowY<=coordinate[3];rowY++,targetRow++) {
			Row Row = sheet.getRow(rowY);
			if(Row==null) continue;
			//获取目标单元格最后一行
			int lastRowNum = targetRow;
			short tFirstCellNum = Row.getFirstCellNum();
			short tLastCellNum = Row.getLastCellNum();
			for(int colX=tFirstCellNum;colX<tLastCellNum;colX++) {
				//判断指定单元格是否已经包含合并单元格了
				if(!isContainsMergedCell(targetSheet,lastRowNum,colX)) {
					//判断此单元格是否是合并单元格
					if(isMergedRegion(rowY, colX)) {
						//获取合并单元格的下标
						int mergedRegionIndex = getMergedRegionIndex(this.sheet,rowY, colX);
						if(mergedRegionIndex==-1) continue;
						CellRangeAddress mergedRegion = sheet.getMergedRegion(mergedRegionIndex);
						//添加合并单元格
						int tFirstRow = lastRowNum;
						int tLastRow = (mergedRegion.getLastRow()-mergedRegion.getFirstRow())+lastRowNum;
						int tFirstColumn = mergedRegion.getFirstColumn();
						int tLastColumn = mergedRegion.getLastColumn();
						Row tRow =targetSheet.getRow(tFirstRow); 
						if(tRow==null) tRow=targetSheet.createRow(tFirstRow);
						Cell tCell = tRow.createCell(tFirstColumn);
						Cell firstCell = getCell(mergedRegion.getFirstRow(),mergedRegion.getFirstColumn());
						//设置单元格的值
						String cellValue = getCellValue(firstCell);
						String[] allKey = getAllKey(cellValue);
						Object value=null;
						for(String str:allKey) {
							String key=getPattermKey(str);
							Object object = data.get(key);
							if((allKey.length>1 || cellValue.length()>str.length()) && (value==null || value instanceof String)) {
								if(value==null) {
									value=cellValue;
								}
								value=value.toString().replace(str, object.toString());
							}else {
								value=object;
							}
						}
						if(data!=null && value!=null) {
							setCellValue(tCell, value);
						}else{
							setCellValue(tCell,firstCell );
						}
						CellRangeAddress cellRangeAddress = new CellRangeAddress(tFirstRow,tLastRow,tFirstColumn,tLastColumn);
						targetSheet.addMergedRegion(cellRangeAddress);
						
						//设置行高
						tRow.setHeight(Row.getHeight());
						tRow.setHeightInPoints(Row.getHeightInPoints());
						setRegionStyle(mergedRegion,cellRangeAddress);
					}else {
						Cell cell = getCell(rowY, colX);
						String cellValue = getCellValue(cell);
						String value=getPattermKey(cellValue);
						Row tRow = targetSheet.getRow(lastRowNum);
						if(tRow==null) tRow=targetSheet.createRow(lastRowNum);
						Cell tCell = tRow.getCell(colX);
						if(tCell==null) tCell = tRow.createCell(colX);
						getOrPutCellStyleMap(cell.getCellStyle());
						tCell.setCellStyle(newCellStyle);
						if(data!=null && value!=null) {
							setCellValue(tCell,data.get(value) );
						}else if(!"".equals(value)){//等于""或null就没有意义再进行设置值了
							setCellValue(tCell,cell );
						}
						//设置行高
						tRow.setHeight(Row.getHeight());
						tRow.setHeightInPoints(Row.getHeightInPoints());
						//设置列宽
						targetSheet.setColumnWidth(colX, sheet.getColumnWidth(colX));
					}
				}
			}
		}
	}
	
	/**
	 *	获取value中的全部key,返回String数组
	 * @param value
	 * @return
	 */
	private String[] getAllKey(String value) {
		String tempStr="";
		Pattern compile = Pattern.compile("(#.+?#)");
		Matcher matcher = compile.matcher(value);
		while(matcher.find()) {
			tempStr+=","+matcher.group(1);
		}
		if(tempStr.startsWith(",")) {
			tempStr=tempStr.substring(1);
		}
		if(tempStr.isEmpty()) {
			return new String[0];
		}
		return tempStr.split(",");
	}
	
	/**
	 * 	从缓存区获取指定样式，如果没有，则进行创建并进行缓存，如果已经存在，则直接从缓存里获取(解决Cell Style数量太多的问题)
	 * @param cellStyle
	 * @return
	 */
	private CellStyle getOrPutCellStyleMap(CellStyle cellStyle) {
		newCellStyle = cellStyleMap.get(String.valueOf(cellStyle.hashCode()));
		if(newCellStyle == null) {
			newCellStyle = targetWorkBook.createCellStyle();
			newCellStyle.cloneStyleFrom(cellStyle);
			newCellStyle.setFillPattern(cellStyle.getFillPattern()==1?CellStyle.SOLID_FOREGROUND:CellStyle.NO_FILL);
			cellStyleMap.put(String.valueOf(cellStyle.hashCode()), newCellStyle);
		}
		return newCellStyle;
	}
	
	/**
	 * 	从源Sheet合并范围的样式同步到目标Sheet指定合并范围的样式，并同步合并单元格的列宽
	 * @param originCellRangeAddress
	 * @param targetCellRangeAddress
	 */
	private void setRegionStyle(CellRangeAddress originCellRangeAddress,CellRangeAddress targetCellRangeAddress) {
		int firstRow = originCellRangeAddress.getFirstRow();
		int lastRow = originCellRangeAddress.getLastRow();
		int firstColumn = originCellRangeAddress.getFirstColumn();
		int lastColumn = originCellRangeAddress.getLastColumn();
		int fr = targetCellRangeAddress.getFirstRow();
		int fc = targetCellRangeAddress.getFirstColumn();
		int lr = targetCellRangeAddress.getLastRow();
		int lc = targetCellRangeAddress.getLastColumn();
		//合并方格范围相等，则进行填充，否则，选区一种样式进行填充
		if((lastRow-firstRow)==(lr-fr) && (lastColumn-firstColumn) == (lc-fc)) {
			for(int y=firstRow,offsetY=0;y<=lastRow;y++,offsetY++) {
				Row row = sheet.getRow(y);
				Row targetRow = targetSheet.getRow(offsetY+fr);
				if(targetRow == null) targetRow=targetSheet.createRow(offsetY+fr);
				for(int x=firstColumn,offsetX=0;x<=lastColumn;x++,offsetX++) {
					Cell cell = row.getCell(x);
					CellStyle newCellStyle = getOrPutCellStyleMap(cell.getCellStyle());
					Cell targetCell = targetRow.getCell(offsetX+fc);
					if(targetCell==null) targetCell = targetRow.createCell(offsetX+fc);
					targetCell.setCellStyle(newCellStyle);
					//设置列宽
					targetSheet.setColumnWidth(offsetX+fc, sheet.getColumnWidth(x));
				}
			}
		}else {
			Cell mmoduleFirstCell = getCell(firstRow, firstColumn);
			Cell mmoduleLastCell = getCell(lastRow, lastColumn);
			CellStyle firstStyle = mmoduleFirstCell.getCellStyle();
			CellStyle lastStyle = mmoduleLastCell.getCellStyle();
			firstStyle.setBorderRight(lastStyle.getBorderRight());
			firstStyle.setBorderBottom(lastStyle.getBorderBottom());
			firstStyle.setRightBorderColor(lastStyle.getRightBorderColor());
			firstStyle.setBottomBorderColor(lastStyle.getBottomBorderColor());
			CellStyle cellStyle = getOrPutCellStyleMap(firstStyle);
			
			for(int y=fr,offsetY=0;y<=lr;y++,offsetY++) {
				Row targetRow = targetSheet.getRow(y);
				if(targetRow == null) targetRow=targetSheet.createRow(y);
				for(int x=fc,offsetX=0;x<=lc;x++,offsetX++) {
					Cell targetCell = targetRow.getCell(x);
					if(targetCell==null) targetCell = targetRow.createCell(x);
					targetCell.setCellStyle(cellStyle);
					//设置列宽
					targetSheet.setColumnWidth(x, sheet.getColumnWidth(firstColumn+offsetX));
				}
				targetRow.setHeight(sheet.getRow(firstRow+offsetY).getHeight());
				targetRow.setHeightInPoints(sheet.getRow(firstRow+offsetY).getHeightInPoints());
			}
			
		}
	}
	
	/**
	 * 	获取源Sheet指定行列的单元格对象
	 * @param row
	 * @param col
	 * @return
	 */
	private Cell getCell(int row,int col) {
		Row _row = sheet.getRow(row);
		if(_row==null) _row=sheet.createRow(row);
		Cell _cell = _row.getCell(col);
		if(_cell==null) _cell=_row.createCell(col);
		return _cell;
	}
	
	/**
	 * 	获取指定Sheet 里的单元格所在的合并单元格的范围对象
	 * @param row
	 * @param col
	 * @return
	 */
	public CellRangeAddress getCellRangeAddress(Sheet sheet,int row,int col) {
		int numMergedRegions = sheet.getNumMergedRegions();
		for(int i=0;i<numMergedRegions;i++) {
			CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
			if(mergedRegion.isInRange(row, col)) {
				return mergedRegion;
			}
		}
		return null;
	}
	
	/**
	 *	判断指定Sheet 指定的单元格是否已经包含合并单元格(例外:一般的单元格不会重复创建的，无需判断)
	 * @return
	 */
	public boolean isContainsMergedCell(Sheet s,int row,int col) {
		int numMergedRegions = s.getNumMergedRegions();
		for(int _i=0;_i<numMergedRegions;_i++) {
			CellRangeAddress mergedRegion = s.getMergedRegion(_i);
			if(mergedRegion.isInRange(row, col)) {
				return true;
			}
		}
		return false;
	}

	
	/**
	 * 	获取源Sheet合并单元格的值
	 * @param row
	 * @param col
	 * @return
	 */
	public String getMergedRegionValue(int row,int col) {
		int numMergedRegions = sheet.getNumMergedRegions();
		for(int i=0;i<numMergedRegions;i++) {
			CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
			int firstRow = mergedRegion.getFirstRow();
			int lastRow = mergedRegion.getLastRow();
			int firstColumn = mergedRegion.getFirstColumn();
			int lastColumn = mergedRegion.getLastColumn();
			if(row >= firstRow && row <= lastRow && col >= firstColumn && col <= lastColumn) {
				Row r = sheet.getRow(firstRow);
				Cell c = r.getCell(firstColumn);
				return String.valueOf(getCellValue(c));
			}
		}
		return "";
	}
	
	
	/**
	 *	如果源Sheet的指定行有合并单元格，则获取这个合并单元格的范围[x,y,x1,y1],如果没有合并单元格，则进行返回指定的行数
	 * @param row
	 * @return 
	 */
	private int[] getRowContainsMerged(int row) {
		int[] r=new int[4];
		int _numMergedRegions = sheet.getNumMergedRegions();
		for(int i=0;i<_numMergedRegions;i++) {
			CellRangeAddress _mergedRegion = sheet.getMergedRegion(i);
			int firstRow = _mergedRegion.getFirstRow();
			int lastRow = _mergedRegion.getLastRow();
			int firstColumn = _mergedRegion.getFirstColumn();
			int lastColumn = _mergedRegion.getLastColumn();
			if(row>=firstRow && row<=lastRow) {
				r[0]=firstColumn;
				r[1]=firstRow;
				r[2]=lastColumn;
				r[3]=lastRow;
				return r;
			}
		}
		r[0]=row;
		r[1]=row;
		r[2]=row;
		r[3]=row;
		return r;
	}
	
	
	/**
	 * 	判断指定Sheet的指定行是否包含合并单元格,如果指定行包含合并单元格，则返回true，否则返回false
	 * @param row
	 * @return 
	 */
	private boolean isRowContainsMerged(Sheet sheet,int row) {
		int _numMergedRegions = sheet.getNumMergedRegions();
		for(int _i=0;_i<_numMergedRegions;_i++) {
			CellRangeAddress _mergedRegion = sheet.getMergedRegion(_i);
			int _firstRow = _mergedRegion.getFirstRow();
			int _lastRow = _mergedRegion.getLastRow();
			if(row>=_firstRow && row<=_lastRow) {
				return true;
			}
		}
		return false;
	}
	
	
	/**
	 *	 判断源Sheet指定的行列是否时合并单元格(的范围)
	 * @param row
	 * @param col
	 * @return
	 */
	private boolean isMergedRegion(int row,int col) {
		int numMergedRegions = sheet.getNumMergedRegions();
		for(int i=0;i<numMergedRegions;i++) {
			CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
			if(mergedRegion.isInRange(row, col)) {
				return true;
			}
		}
		return false;
	}
	
	/**
	 *	获取指定Sheet里的合并单元格的所在行(row)列(col)的Index
	 * @param row
	 * @param col
	 * @return
	 */
	public int getMergedRegionIndex(Sheet sheet,int row,int col) {
		int numMergedRegions = sheet.getNumMergedRegions();
		for(int i=0;i<numMergedRegions;i++) {
			CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
			if(mergedRegion.isInRange(row, col)) {
				return i;
			}
		}
		return -1;
	}
	
	/**
	 * 	获取指定单元格的值
	 * @return
	 */
	public String getCellValue(Cell cell) {
		if(cell!=null) {
			int cellType = cell.getCellType();
			if(Cell.CELL_TYPE_BLANK == cellType) {
				return "";
			}else if(Cell.CELL_TYPE_BOOLEAN == cellType) {
				return String.valueOf(cell.getBooleanCellValue());
			}else if(Cell.CELL_TYPE_STRING == cellType) {
				return cell.getStringCellValue();
			}else if(Cell.CELL_TYPE_FORMULA == cellType) {
				return cell.getCellFormula();
			}else if(Cell.CELL_TYPE_NUMERIC == cellType) {
				return String.valueOf(cell.getNumericCellValue());
			}else if(Cell.CELL_TYPE_ERROR == cellType) {
				return String.valueOf(cell.getErrorCellValue());
			}
		}
		return "";
	}
	
	/**
	 *	 把指定value值设置到这个单元格里
	 * @param targetCell
	 * @param value
	 */
	public void setCellValue(Cell targetCell,Object value) {
		String val = String.valueOf(value);
		if(value instanceof String) {
			targetCell.setCellType(Cell.CELL_TYPE_STRING);
			targetCell.setCellValue(val);
		}else if(value instanceof Number) {
			targetCell.setCellType(Cell.CELL_TYPE_NUMERIC);
			targetCell.setCellValue(Double.valueOf(val));
		}
	}
	
	/**
	 *	将源Cell里的值设置到targetCell值中
	 * @param targetCell
	 * @param originCell
	 */
	public void setCellValue(Cell targetCell,Object value,int cellType) {
		targetCell.setCellType(cellType);
		String cellValue = String.valueOf(value);
		switch (cellType) {
			case Cell.CELL_TYPE_BLANK:
			case Cell.CELL_TYPE_STRING:
				targetCell.setCellValue(cellValue);
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				targetCell.setCellValue(Boolean.valueOf(cellValue));
				break;
			case Cell.CELL_TYPE_NUMERIC:
				targetCell.setCellValue(Double.valueOf(cellValue));
				break;
			case Cell.CELL_TYPE_FORMULA:
				targetCell.setCellFormula(cellValue);
				break;
			case Cell.CELL_TYPE_ERROR:
				targetCell.setCellErrorValue(Byte.valueOf(cellValue));
				break;
		}
	}
	
	/**
	 *	将源Cell里的值设置到targetCell值中
	 * @param targetCell
	 * @param originCell
	 */
	public void setCellValue(Cell targetCell,Cell originCell) {
		int cellType = originCell.getCellType();
		targetCell.setCellType(cellType);
		String cellValue = getCellValue(originCell);
		switch (cellType) {
			case Cell.CELL_TYPE_BLANK:
			case Cell.CELL_TYPE_STRING:
				targetCell.setCellValue(cellValue);
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				targetCell.setCellValue(Boolean.valueOf(cellValue));
				break;
			case Cell.CELL_TYPE_NUMERIC:
				targetCell.setCellValue(Double.valueOf(cellValue));
				break;
			case Cell.CELL_TYPE_FORMULA:
				targetCell.setCellFormula(cellValue);
				break;
			case Cell.CELL_TYPE_ERROR:
				targetCell.setCellErrorValue(Byte.valueOf(cellValue));
				break;
		}
	}
	
	/**
	 * 	获取Excel的匹配Key名称
	 * @param key
	 * @return
	 */
	public String getPattermKey(String key) {
		Pattern compile = Pattern.compile("#(.+)#");
		Matcher matcher = compile.matcher(key);
		if(matcher.matches() && matcher.groupCount()>0) {
			return matcher.group(1);
		}
		return null;
	}
	
	/**
	 * 	向目标targetWorkBook添加图片，返回图片的索引
	 * @throws IOException 
	 * @throws FileNotFoundException 
	 */
	public int addPicture(File imageFile) throws FileNotFoundException, IOException {
		if(drawing==null) {
			drawing = targetSheet.createDrawingPatriarch();
		}
		if(imageFile.exists()) {
			if(imageFile.isFile()) {
				String filename = imageFile.getName();
				int extendOffsetIndex = filename.lastIndexOf(".");
				if(extendOffsetIndex==-1) {
					return -1;
				}
				String fileExtend = filename.substring(extendOffsetIndex+1);
				int imageExtendType = getImageExtendType(fileExtend);
				if(imageExtendType!=-1) {
					BufferedImage bufferedImage = ImageIO.read(new FileInputStream(imageFile));
					ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
					ImageIO.write(bufferedImage, fileExtend , byteArrayOutputStream);
					return targetWorkBook.addPicture(byteArrayOutputStream.toByteArray(), imageExtendType);
				}
			}
		}
		return -1;
	}
	
	/**
	 * 	创建锚点
	 * @param dx1
	 * @param dy1
	 * @param dx2
	 * @param dy2
	 * @param col1
	 * @param row1
	 * @param col2
	 * @param row2
	 * @param pictureIndex
	 */
	public void setAuthor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2,int pictureIndex) {
		ClientAnchor anchor = drawing.createAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
		drawing.createPicture(anchor, pictureIndex);
	}
	
	/**
	 * 	创建锚点并设置规模
	 * @param dx1
	 * @param dy1
	 * @param dx2
	 * @param dy2
	 * @param col1
	 * @param row1
	 * @param col2
	 * @param row2
	 * @param pictureIndex
	 * @param scale
	 */
	public void setAuthor(int dx1, int dy1, int dx2, int dy2, int col1, int row1, int col2, int row2,int pictureIndex,double scale) {
		ClientAnchor anchor = drawing.createAnchor(dx1, dy1, dx2, dy2, col1, row1, col2, row2);
		Picture picture = drawing.createPicture(anchor, pictureIndex);
		picture.resize(scale);
	}
	
	/**
	 * 	合并指定范围的单元格，如果范围内的单元格如已被合并，则进行删除，再进行合并
	 */
	public void addMergedRange(int moduleMergedindex,CellRangeAddress cellRangeAddress) {
		int firstRow = cellRangeAddress.getFirstRow();
		int lastRow = cellRangeAddress.getLastRow();
		int firstColumn = cellRangeAddress.getFirstColumn();
		int lastColumn = cellRangeAddress.getLastColumn();
		
		Set<Integer> indexSet=new LinkedHashSet<Integer>();
		
		for(int row=firstRow;row<=lastRow;row++) {
			for(int col=firstColumn;col<=lastColumn;col++) {
				int index = getMergedRegionIndex(this.targetSheet,row,col);
				if(index!=-1 && !indexSet.contains(index)) {
					indexSet.add(index);
				}
			}
		}
		
		Integer[] indexs=new Integer[indexSet.size()];
		indexSet.toArray(indexs);
		Arrays.sort(indexs);
		for(int i=indexs.length-1;i>=0;i--) {
			targetSheet.removeMergedRegion(indexs[i]);
		}
		
		targetSheet.addMergedRegion(cellRangeAddress);
		setRegionStyle(this.sheet.getMergedRegion(moduleMergedindex), cellRangeAddress);
	}
	
	/**
	 * 根据扩展类型 获取POI(WorkBook)支持的扩展类型
	 * @param extendType
	 * @return
	 */
	private int getImageExtendType(String extendType) {
		if(extendType!=null) {
			extendType=extendType.toUpperCase();
			if(Pattern.matches("^(JPEG|JPG)$", extendType)) {
				return imageExtend.get("JPEG").intValue();
			}else{
				 Integer integer = imageExtend.get(extendType);
				 if(integer!=null) {
					return integer.intValue();
				 }
			}
		}
		return -1;
	}
	
	/**
	 * 	将数据写入到指定路径文件下
	 * @param targerPath
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public void write(String targerPath) throws FileNotFoundException, IOException {
		targetWorkBook.write(new FileOutputStream(targerPath));
	}
	
	/**
	 * 	将数据写入到写入流中
	 * @param os
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public void write(OutputStream os) throws FileNotFoundException, IOException {
		targetWorkBook.write(os);
	}
}
