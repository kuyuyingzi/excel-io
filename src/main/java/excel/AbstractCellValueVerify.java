package excel;

/**
 * 列值校验
 * @author lixiangjing
 *
 */
public abstract class AbstractCellValueVerify {
	public abstract Object verify(Object fileValue) throws Exception;
}
