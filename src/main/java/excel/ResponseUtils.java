package excel;


public class ResponseUtils {

	/*public static void write(String excelName, Workbook wb) throws Exception{
		HttpServletResponse response = HttpInvokeHelper.getHttpServletResponse();
		response.setHeader("Content-Disposition", "attachment;filename=" + new String((excelName).getBytes("GBK"), "ISO8859-1"));//设定输出文件头
		response.setContentType("application/vnd.ms-excel;charset=UTF-8");// 定义输出类型
		OutputStream out = response.getOutputStream();
		wb.write(out);
		out.flush();
		out.close();
	}

	public static void writeFile(String path) throws Exception{
		File file = new File(HttpInvokeHelper.getHttpServletRequest().getServletContext().getRealPath("/")+path);
		InputStream fis = new BufferedInputStream(new FileInputStream(file));
		byte[] buffer = new byte[fis.available()];
		fis.read(buffer); fis.close();
		
		HttpServletResponse response = HttpInvokeHelper.getHttpServletResponse();
		response.reset();
		response.addHeader("Content-Disposition", "attachment;filename=" + new String((file.getName()).getBytes("GBK"), "ISO8859-1"));
		response.addHeader("Content-Length", "" + file.length());
		OutputStream toClient = new BufferedOutputStream(response.getOutputStream());
		toClient.write(buffer);
		toClient.flush();
		toClient.close();
	}*/

}
