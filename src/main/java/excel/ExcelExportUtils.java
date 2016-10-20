package excel;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;

public class ExcelExportUtils {

	public static void printSetup(Sheet sheet) {
		PrintSetup printSetup = sheet.getPrintSetup();
		// 打印方向，true：横向，false：纵向
		printSetup.setLandscape(true);
		sheet.setFitToPage(true);
		sheet.setHorizontallyCenter(true);
	}

	public static Map<String, CellStyle> initStyles(Workbook wb) {
		Map<String, CellStyle> styles = new HashMap<String, CellStyle>();
		CellStyle style;
		Font titleFont = wb.createFont();
		titleFont.setFontHeightInPoints((short) 18);
		titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_LEFT);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setFont(titleFont);
		styles.put("title", style);

		Font monthFont = wb.createFont();
		monthFont.setFontHeightInPoints((short) 11);
		monthFont.setColor(IndexedColors.WHITE.getIndex());
		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setFont(monthFont);
		style.setWrapText(true);
		styles.put("header", style);

		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setWrapText(true);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		styles.put("cell", style);

		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
		styles.put("formula", style);

		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
		styles.put("formula_2", style);

		style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setWrapText(true);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		styles.put("yellowcell", style);

		return styles;
	}
	
	/**
	 * @param list
	 * @param columns
	 *            单元格数据【标题,列字段,列宽】
	 * @param numberics
	 *            数字字段
	 * @return
	 * @throws Exception
	 */
	public static <T> Workbook createWorkbook(List<T> list,
			Object[][] columns) throws Exception {
		return createWorkbook(list, columns, false);
	}

	/**
	 * @param list
	 * @param columns
	 *            单元格数据【标题,列字段,列宽】
	 * @param withNum
	 *            带序号
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings("rawtypes")
	public static <T> Workbook createWorkbook(List<T> list,
			Object[][] columns, boolean withNum) throws Exception {
		Workbook wb = new HSSFWorkbook();
		Map<String, CellStyle> styles = ExcelExportUtils.initStyles(wb);
		Sheet sheet = wb.createSheet();
		ExcelExportUtils.printSetup(sheet);

		// header row
		Row headerRow = sheet.createRow(0);
		for (int i = 0; i < columns.length; i++) {
			CellUtil.createCell(headerRow, i, (String) columns[i][0],
					styles.get("header"));
			sheet.setColumnWidth(i, (Integer) columns[i][2]);
		}

		// body row
		for (int i = 0; i < list.size(); i++) {
			Row row = sheet.createRow(i + 1);
			for (int j = 0; j < columns.length; j++) {
				Cell cell = row.createCell(j);
				cell.setCellStyle(styles.get("cell"));
				// 有效数据行号
				if (withNum && j == 0) {
					cell.setCellValue(i + 1);
					continue;
				}
				T t = list.get(i);
				// 读取Map/Object对应字段值
				Object value = null;
				if(t instanceof Map){
					value = ((Map) t).get(columns[j][1]);
				}else{
					value = FieldUtils.readDeclaredField(t,
							(String) columns[j][1], true);
				}
				if (value == null) {
					value = "";
				}
				// 填充列值
				cell.setCellValue(String.valueOf(value));
			}
		}
		return wb;
	}

	/*
	 * public static Workbook createJXLSWorkbook(List<Map<String, Object>> list,
	 * String templateFile) throws Exception { XLSTransformer transformer = new
	 * XLSTransformer(); Map<String, Object> beanParams = new HashMap<>();
	 * beanParams.put("list", list); InputStream is = new
	 * BufferedInputStream(new FileInputStream(templateFile)); Workbook wb =
	 * transformer.transformXLS(is, beanParams); return wb; }
	 */

}
