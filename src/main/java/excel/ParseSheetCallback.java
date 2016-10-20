package excel;

public interface ParseSheetCallback<T> {
	void callback(T t, int rowNum) throws Exception;
}
