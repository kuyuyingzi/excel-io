package excel;

import org.apache.commons.lang3.StringUtils;

import common.MessageException;

/**
 * 列字符串转数字校验
 * @author Administrator
 *
 */
public class StringToIntegerVerify extends AbstractCellVerify{
	private String cellName;
	private AbstractCellValueVerify cellValueVerify;
	private boolean allowNull;
	
	public StringToIntegerVerify(String cellName, boolean allowNull) {
		super();
		this.cellName = cellName;
		this.allowNull = allowNull;
	}

	public StringToIntegerVerify(String cellName, AbstractCellValueVerify cellValueVerify, boolean allowNull) {
		this.cellName = cellName;
		this.cellValueVerify = cellValueVerify;
		this.allowNull = allowNull;
	}

	@Override
	public Object verify(Object cellValue) throws Exception {
		if(allowNull){
			if(cellValue!=null && StringUtils.isNotEmpty(format(cellValue))){
				cellValue = format(cellValue);
				if(null!=cellValueVerify){
					cellValue = cellValueVerify.verify(cellValue);
				}
				return cellValue;
			}
			return cellValue;
		}
		
		if(cellValue==null || StringUtils.isEmpty(format(cellValue))){
			throw MessageException.newMessageException(cellName+"不能为空");
		}
		cellValue = format(cellValue);
		if(null!=cellValueVerify){
			cellValue = cellValueVerify.verify(cellValue);
		}
		return cellValue;
	}

	private String format(Object fileValue) {
		return String.valueOf(fileValue);
	}
}
