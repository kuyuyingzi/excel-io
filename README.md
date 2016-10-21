# excel-io
excel导入导出通用工具包，详细介绍http://blog.csdn.net/kuyuyingzi/article/details/51323072
#特性
	导入模板具有以下特性：
	1、列格式化和列值校验，是否允许空判断
	2、可指定列提取
	3、提供回调函数，进行额外字段填充和业务逻辑
	4、报错机制，报错提示批量、准确具体行列

	导出模板具有以下特性：

	1、可生成序列号

	2、可指定列长度

	此套导入和导出开发模板简单易用，可读性强，维护方便，让程序员避开复杂的代码操作，专注业务。
#导入示例：
	/**
	 * 导入用户校验类
	 * @author Administrator
	 *
	 */
	public class MemberVerifyBuilder extends AbstractVerifyBuidler {

		private static MemberVerifyBuilder builder = new MemberVerifyBuilder();

		public static MemberVerifyBuilder getInstance() {
			return builder;
		}

		/**
		 * 定义列校验实体：提取的字段、提取列、校验规则
		 */
		private MemberVerifyBuilder() {
			cellEntitys.add(new CellVerifyEntity("name", "A", new StringVerify("姓名", true)));
			cellEntitys.add(new CellVerifyEntity("age", "B", new IntegerVerify("年龄", true)));
			cellEntitys.add(new CellVerifyEntity("country", "D", new StringToIntegerVerify("国家",
					new AbstractCellValueVerify() {
						@Override
						public Object verify(Object fileValue) {
							// TODO 转换：从excel中得到string转成需要的integer
							return 1;
						}
					}, true)));
			cellEntitys.add(new CellVerifyEntity("date", "F", new DateTimeVerify("创建日期", "yyyy/MM/dd",
					true)));

			// 必须调用
			super.init();
		}
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
#待优化
	校验实体类可改为注解的方式
