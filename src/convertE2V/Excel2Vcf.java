package convertE2V;

import java.io.BufferedInputStream;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;

/**
 * 
 * @author liuxiuquan
 * excel转vcf的工具类 
 * 
 */
public class Excel2Vcf {

	/**
	 * 读取excel 返回泛型为Excel2VcfBean的集合
	 * 
	 * @param excelPath 文件路径
	 * @return 返回List<Excel2VcfBean>
	 */
	public static List<Excel2VcfBean> readExcel(String excelPath)
			throws Exception {
		File file = new File(excelPath);
		// 初始化输入流
		FileInputStream is = new FileInputStream(file);
		// 定义返回List
		List<Excel2VcfBean> excelList = new ArrayList<Excel2VcfBean>();
		// 读入文件成功后 创建Excel,并指定Excel读取位置
		HSSFWorkbook wb = new HSSFWorkbook(is);
		// 根据Workbook得到第0个下标的工作薄
		HSSFSheet sheet1 = wb.getSheetAt(0);
		int rowMax = sheet1.getLastRowNum();
		System.out.println("--rowMax--" + rowMax);
		Map<String, String> titleMap = readProperties("title.properties");

		Row row = null;
		Map<String, Integer> numNameMap = new HashMap<String, Integer>();
		// 从前15行取标题列名称
		for (int j = 0; j <= ((15 < rowMax) ? 15 : rowMax); j++) {
			row = sheet1.getRow(j);
			int maxColumn = row.getPhysicalNumberOfCells();
			// 如果一行的列数大于0
			if (maxColumn > 0) {
				for (int k = 0; k < maxColumn; k++) {
					String res = whichCellType(row.getCell(k));
					if (null != res && titleMap.containsKey(res)) {
						numNameMap.put(titleMap.get(res), k);
					}
				}
			}
		}

		// 遍历工作薄中的所有行
		for (int i = 0; i <= rowMax; i++) {
			row = sheet1.getRow(i);
			int maxColumn = row.getPhysicalNumberOfCells();
			// 如果一行的列数大于0
			if (maxColumn > 0) {
				// 保存一条记录
				Excel2VcfBean excelBean = new Excel2VcfBean();
				// 姓名
				String name = null;
				// 手机号码
				String tel = null;
				// 应急电话
				String emergencyTel = null;
				// 公司
				String company = "精益有容";
				// 工作地点
				String address = null;
				if (numNameMap.containsKey("name")) {
					name = whichCellType(row.getCell(numNameMap.get("name")));
				}
				if (numNameMap.containsKey("tel")) {
					tel = whichCellType(row.getCell(numNameMap.get("tel")));
				}
				if (numNameMap.containsKey("TelEme")) {
					emergencyTel = whichCellType(row.getCell(numNameMap
							.get("TelEme")));
				}
				if (numNameMap.containsKey("address")) {
					address = whichCellType(row.getCell(numNameMap
							.get("address")));
				}
				if (null != name && null != tel && !"姓名".equals(name)
						&& !"手机号码".equals(tel) && !"工作地点".equals(tel)) {
					excelBean.setName(name);
					excelBean.setTel(tel);
					excelBean.setTelEme(emergencyTel);
					excelBean.setCompany(company + "-" + address);
					excelList.add(excelBean);
				}
			}
		}
		return excelList;
	}

	/**
	 * 解析properties为Map
	 * properties中存放的是实际用到的excel表的列字段映射，根据不同情况修改对应的name列
	 * 
	 * @param filePath 读取properties的全部信息
	 * @return 解析生成的Map
	 */
	public static Map<String, String> readProperties(String filePath) {
		Properties props = new Properties();
		Map<String, String> prop = new HashMap<String, String>();
		try {
			InputStream in = new BufferedInputStream(new FileInputStream(
					filePath));
			props.load(in);
			Enumeration en = props.propertyNames();
			while (en.hasMoreElements()) {
				String key = (String) en.nextElement();
				String Property = props.getProperty(key);
				// System.out.println(key +":"+ Property);
				prop.put(key, Property);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return prop;
	}

	/**
	 * 根据单元格内容格式的不同返回不同的字符串 
	 * 
	 * @param cell 单元格
	 * @return 转换后的字符串
	 */
	private static String whichCellType(Cell cell) {
		// 返回的字符串
		String result = null;
		if (null != cell) {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:// String类型单元格
				// 文本类型
				result = cell.getRichStringCellValue().getString();
				break;
			case Cell.CELL_TYPE_NUMERIC:// 数字类型
				// 检查单元格是否包含一个Date类型
				if (DateUtil.isCellDateFormatted(cell)) {
					// 日期类型
					result = cell.getDateCellValue().toString();
				} else {
					// 数字类型
					// 将1.3031113083E10转为13031113083
					DecimalFormat df = new DecimalFormat("#");
					result = df.format(cell.getNumericCellValue());
				}
				break;
			}
		}
		return result;
	}

	/**
	 * 写vcf工具类
	 * @param excelList 写入的数据
	 * @param vcfPath vcf的路径
	 * @return 是否成功执行
	 */

	public static Boolean WriteVcf(List<Excel2VcfBean> excelList, String vcfPath) {
		// 标识是否成功执行
		boolean ok = false;
		File file = new File(vcfPath);
		if (file.exists()) {
			file.delete();
		}
		try {
			file.createNewFile();
			FileOutputStream fos = new FileOutputStream(vcfPath, false);
			// 指定编码格式为utf-8
			OutputStreamWriter osw = new OutputStreamWriter(fos, "utf-8");
			BufferedWriter bw = new BufferedWriter(osw);

			for (Excel2VcfBean excelBean : excelList) {
				// bw.write(a.toString());
				StringBuffer sb = new StringBuffer();
				// BEGIN:VCARD
				// VERSION:3.0
				// N:;;;;
				// FN:刘达
				// TEL;TYPE=CELL:123456789
				// ORG:河北企业 石家庄
				// END:VCARD
				if (null != excelBean.getName() && null != excelBean.getTel()) {
					sb.append("BEGIN:VCARD").append("\r\n");
					sb.append("VERSION:3.0").append("\r\n");
					sb.append("N:;;;;").append("\r\n");
					sb.append("FN:").append(excelBean.getName()).append("\r\n");
					sb.append("TEL;TYPE=CELL:").append(excelBean.getTel())
							.append("\r\n");
					if (null != excelBean.getTelEme()) {
						sb.append("TEL;TYPE=CELL:")
								.append(excelBean.getTelEme()).append("\r\n");
					}
					sb.append("ORG:").append(excelBean.getCompany())
							.append("\r\n");
					sb.append("END:VCARD").append("\r\n");
					bw.write(sb.toString());
				}
			}
			bw.close();
			osw.close();
			fos.close();
			ok = true;
		} catch (FileNotFoundException e) {
			ok = false;
			// e.printStackTrace();
		} catch (UnsupportedEncodingException e) {
			ok = false;
			// e.printStackTrace();
		} catch (IOException e) {
			ok = false;
			// e.printStackTrace();
		}
		return ok;
	}

	public static void main(String[] args) throws Exception {
		List<Excel2VcfBean> excelList = readExcel("D:/workbook.xls");
		for (Excel2VcfBean a : excelList) {
			System.out.println(a);
		}
		boolean ok = WriteVcf(excelList, "D:/test1.vcf");
		String res = (ok == true) ? "成功" : "失败";
		System.out.println("--完成，执行结果为--:" + res);
	}
}
