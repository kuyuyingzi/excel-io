package excel;

import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

import common.MessageException;

public class ExcelImportUtils {
	
	/**
	 * 解析Sheet
	 * 
	 * @param clss
	 *            结果bean
	 * @param verifyBuilder
	 *            校验器
	 * @param sheet
	 * @param dataStartRow
	 *            开始行:从0开始计
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> parseSheet(Class<T> clss, AbstractVerifyBuidler verifyBuilder,
			Sheet sheet, int dataStartRow) throws Exception {
		return parseSheet(clss, verifyBuilder, sheet, dataStartRow, null);
	}

	/**
	 * 解析Sheet
	 * 
	 * @param clss
	 *            结果bean
	 * @param verifyBuilder
	 *            校验器
	 * @param sheet
	 * @param dataStartRow
	 *            开始行
	 * @param callback
	 *            加入回调逻辑
	 * @return
	 * @throws Exception
	 */
	public static <T> List<T> parseSheet(Class<T> clss, AbstractVerifyBuidler verifyBuilder,
			Sheet sheet, int dataStartRow, ParseSheetCallback<T> callback) throws Exception {
		List<T> beans = new ArrayList<>();
		StringBuffer errors = new StringBuffer();
		StringBuffer rowErrors = new StringBuffer();
		int rowStart = sheet.getFirstRowNum() + dataStartRow;
		int rowEnd = sheet.getLastRowNum();
		for (int rowNum = rowStart; rowNum <= rowEnd; rowNum++) {
			Row r = sheet.getRow(rowNum);
			// 创建对象
			T t = clss.newInstance();
			int fieldNum = 0;
			for (int cellNum : POIConstant.convertToCellNum(verifyBuilder.cellRefs)) {
				// 列坐标
				CellReference cellRef = new CellReference(rowNum, cellNum);
				try {
					Object cellValue = getCellValue(r, cellNum);
					// 校验和格式化列值
					cellValue = verifyBuilder.verify(verifyBuilder.filedNames[fieldNum],
							cellValue);
					// 填充列值
					FieldUtils.writeDeclaredField(t, verifyBuilder.filedNames[fieldNum], cellValue, true);
				} catch (Exception e) {
					if (e instanceof MessageException) {
						rowErrors.append(cellRef.formatAsString()).append(":")
								.append(e.getMessage()).append("\t");
					}
					e.printStackTrace();
				}
				fieldNum++;
			}
			// 回调处理一下特殊逻辑
			if(callback!=null){
				callback.callback(t, rowNum);
			}
			beans.add(t);
			if (rowErrors.length() > 0) {
				errors.append(rowErrors).append("\r\n");
				rowErrors.setLength(0);
			}
		}

		// throw parse exception
		if (errors.length() > 0) {
			throw MessageException.newMessageException(errors.toString());
		}
		// 返回结果
		return beans;
	}
	
	public static Object getCellValue(Row r, int cellNum) {
		// 缺失列处理政策
		Cell cell = r.getCell(cellNum, Row.CREATE_NULL_AS_BLANK);
		Object obj = null;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			obj = cell.getRichStringCellValue().getString();
			break;
		case Cell.CELL_TYPE_NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				obj = cell.getDateCellValue();
			} else {
				obj = cell.getNumericCellValue();
			}
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			obj = cell.getBooleanCellValue();
			break;
		case Cell.CELL_TYPE_FORMULA:
			obj = cell.getCellFormula();
			break;
		}
		return obj;
	}
}
