package example.imp;

import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.alibaba.fastjson.JSON;

import excel.ExcelImportUtils;
import excel.ParseSheetCallback;

public class MainClass {

	public static void main(String[] args) throws Exception {
		parseSheet();
		parseSheetWithCallback();
	}
	
	/**
	 * 解析excel
	 * @throws Exception
	 */
	public static void parseSheet() throws Exception {
		Workbook wb = WorkbookFactory.create(new FileInputStream("src/test/java/example/imp/import.xlsx"));
		Sheet sheet = wb.getSheetAt(0);
		// parseSheet
		List<Member> list = ExcelImportUtils.parseSheet(
				Member.class, MemberVerifyBuilder.getInstance(), sheet, 1);
		System.out.println(JSON.toJSONString(list));
	}
	
	/**
	 * 解析excel带回调函数：做一些额外字段填充
	 * @throws Exception
	 */
	public static void parseSheetWithCallback() throws Exception {
		Workbook wb = WorkbookFactory.create(new FileInputStream("src/test/java/example/imp/import.xlsx"));
		Sheet sheet = wb.getSheetAt(0);
		final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		// parseSheet
		List<Member> list = ExcelImportUtils.parseSheet(
				Member.class, MemberVerifyBuilder.getInstance(), sheet, 1, new ParseSheetCallback<Member>() {
					@Override
					public void callback(Member t, int rowNum) throws Exception {
						t.setDateDesc(sdf.format(t.getDate()));
					}
				});
		System.out.println(JSON.toJSONString(list));
	}
}
