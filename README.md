# excel-io
excel导入导出通用工具包
#导入示例：
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
#导出示例：
	/**
	 * 定义：excel标题，提取数据的字段，占用列长度
	 */
	private static final Object[][] fields = new Object[][]{
		{"姓名","name",POIConstant.TIME},
		{"年龄","age",POIConstant.NUMBER},
		{"国家","country",POIConstant.NAME},
		{"创建日期","dateDesc",POIConstant.TIME},
	};
	
	private static final Object[][] fieldsWithNum = new Object[][]{
		{"序号","",POIConstant.NUMBER},
		{"姓名","name",POIConstant.TIME},
		{"年龄","age",POIConstant.NUMBER},
		{"国家","country",POIConstant.NAME},
		{"创建日期","dateDesc",POIConstant.TIME},
	};
	
	public static void main(String[] args) throws Exception {
		exportBean();
		exportBeanWithNum();
		exportMap();
	}
	
	/**
	 * 1、导出对象
	 * @throws Exception
	 */
	public static void exportBean() throws Exception{
		List<Member> list = new ArrayList<>();
		list.add(new Member("张三", 28, 1, "2016-10-19"));
		list.add(new Member("李四", 25, 2, "2016-10-19"));
		Workbook bean = ExcelExportUtils.createWorkbook(list, fields);
		bean.write(new FileOutputStream("src/test/java/example/exp/exportBean.xlsx"));
	}
	
	/**
	 * 2、导出对象带序列号
	 * @throws Exception
	 */
	public static void exportBeanWithNum() throws Exception{
		List<Member> list = new ArrayList<>();
		list.add(new Member("张三", 28, 1, "2016-10-19"));
		list.add(new Member("李四", 25, 2, "2016-10-19"));
		Workbook bean = ExcelExportUtils.createWorkbook(list, fieldsWithNum, true);
		bean.write(new FileOutputStream("src/test/java/example/exp/exportBeanWithNum.xlsx"));
	}
	
	/**
	 * 导出map
	 * @throws Exception
	 */
	public static void exportMap() throws Exception{
		List<Map<String, Object>> list = new ArrayList<>();
		Map<String, Object> map = new HashMap<>();
		map.put("name", "张三");
		map.put("age", 28);
		map.put("country", 1);
		map.put("dateDesc", "2016-10-19");
		list.add(map);
		map = new HashMap<>();
		map.put("name", "张三");
		map.put("age", 28);
		map.put("country", 1);
		map.put("dateDesc", "2016-10-19");
		list.add(map);
		Workbook bean = ExcelExportUtils.createWorkbook(list, fields);
		bean.write(new FileOutputStream("src/test/java/example/exp/exportMap.xlsx"));
	}
