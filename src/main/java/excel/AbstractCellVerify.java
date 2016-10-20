package excel;

/**
 * 列校验和格式化接口
 * @author lixiangjing
 *
 */
public abstract class AbstractCellVerify {
	public abstract Object verify(Object cellValue) throws Exception;
}
