package excel;

import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;


public class POIConstant {
	/**
	 * 定义允许上传的文件扩展名
	 */
	public static Set<String> extSet = new HashSet<String>(Arrays.asList("xls","xlsx"));
	//单元格坐标对应数字Map
	public static Map<String, Integer> cellRefNums = new HashMap<>();
	static{
		cellRefNums.put("A",0);	cellRefNums.put("B",1);	cellRefNums.put("C",2);	cellRefNums.put("D",3);	cellRefNums.put("E",4);	cellRefNums.put("F",5);
		cellRefNums.put("G",6);	cellRefNums.put("H",7);	cellRefNums.put("I",8);	cellRefNums.put("J",9);	cellRefNums.put("K",10);cellRefNums.put("L",11);
		cellRefNums.put("M",12);cellRefNums.put("N",13);cellRefNums.put("O",14);cellRefNums.put("P",15);cellRefNums.put("Q",16);cellRefNums.put("R",17);
		cellRefNums.put("S",18);cellRefNums.put("T",19);cellRefNums.put("U",20);cellRefNums.put("V",21);cellRefNums.put("W",22);cellRefNums.put("X",23);
		cellRefNums.put("Y",24);cellRefNums.put("Z",25);cellRefNums.put("AA",26);cellRefNums.put("AB",27);cellRefNums.put("AC",28);cellRefNums.put("AD",29);
		cellRefNums.put("AE",30);cellRefNums.put("AF",31);cellRefNums.put("AG",32);cellRefNums.put("AH",33);cellRefNums.put("AI",34);cellRefNums.put("AJ",35);
		cellRefNums.put("AK",36);cellRefNums.put("AL",37);cellRefNums.put("AM",38);cellRefNums.put("AN",39);cellRefNums.put("AO",40);cellRefNums.put("AP",41);
		cellRefNums.put("AQ",42);cellRefNums.put("AR",43);cellRefNums.put("AS",44);cellRefNums.put("AT",45);cellRefNums.put("AU",46);cellRefNums.put("AV",47);
		cellRefNums.put("AW",48);cellRefNums.put("AX",49);cellRefNums.put("AY",50);cellRefNums.put("AZ",51);}
	//----------------------------------------start 列宽度定义 -------------------------------------
	/**
	 * 列宽单位-字符
	 */
	public static int CHARUNIT = 2*256;
	/**
	 * 列宽单位-数字
	 */
	public static int NUMBERUNIT = new Double(CHARUNIT*0.65).intValue();
	/**
	 * 默认名称宽度
	 */
	public static int NAME = 5 * CHARUNIT;
	/**
	 * 默认银行账户宽度
	 */
	public static int BANK = 20 * NUMBERUNIT;
	/**
	 * 默认数字宽度
	 */
	public static int NUMBER = 8 * NUMBERUNIT;
	/**
	 * 默认金额宽度
	 */
	public static int MONEY = 12 * NUMBERUNIT;
	/**
	 * 默认时间宽度
	 */
	public static int TIME = 10 * NUMBERUNIT;
	/**
	 * 默认时间宽度-长时间
	 */
	public static int LONGTIME = 20 * NUMBERUNIT;
	/**
	 * 默认备注宽度
	 */
	public static int REMARK = 10 * CHARUNIT;
	/**
	 * 默认长备注宽度
	 */
	public static int LONGREMARK = 20 * CHARUNIT;
			
	public static int charWidth(int charNum){
		return CHARUNIT * charNum;
	}
	
	public static int numberWidth(int charNum){
		return NUMBERUNIT * charNum;
	}
	
	/**
	 * 转换列坐标为数字
	 * @param cellRefs 列坐标
	 * @return
	 */
	public static int[] convertToCellNum(String[] cellRefs){
		int[] nums = new int[cellRefs.length];
		for(int i=0; i<cellRefs.length; i++){
			nums[i] = cellRefNums.get(cellRefs[i]);
		}
		return nums;
	}
	//----------------------------------------end 列宽度定义 -------------------------------------
}
