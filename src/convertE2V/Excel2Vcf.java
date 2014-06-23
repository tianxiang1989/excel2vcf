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
 * excelתvcf�Ĺ����� 
 * 
 */
public class Excel2Vcf {

	/**
	 * ��ȡexcel ���ط���ΪExcel2VcfBean�ļ���
	 * 
	 * @param excelPath �ļ�·��
	 * @return ����List<Excel2VcfBean>
	 */
	public static List<Excel2VcfBean> readExcel(String excelPath)
			throws Exception {
		File file = new File(excelPath);
		// ��ʼ��������
		FileInputStream is = new FileInputStream(file);
		// ���巵��List
		List<Excel2VcfBean> excelList = new ArrayList<Excel2VcfBean>();
		// �����ļ��ɹ��� ����Excel,��ָ��Excel��ȡλ��
		HSSFWorkbook wb = new HSSFWorkbook(is);
		// ����Workbook�õ���0���±�Ĺ�����
		HSSFSheet sheet1 = wb.getSheetAt(0);
		int rowMax = sheet1.getLastRowNum();
		System.out.println("--rowMax--" + rowMax);
		Map<String, String> titleMap = readProperties("title.properties");

		Row row = null;
		Map<String, Integer> numNameMap = new HashMap<String, Integer>();
		// ��ǰ15��ȡ����������
		for (int j = 0; j <= ((15 < rowMax) ? 15 : rowMax); j++) {
			row = sheet1.getRow(j);
			int maxColumn = row.getPhysicalNumberOfCells();
			// ���һ�е���������0
			if (maxColumn > 0) {
				for (int k = 0; k < maxColumn; k++) {
					String res = whichCellType(row.getCell(k));
					if (null != res && titleMap.containsKey(res)) {
						numNameMap.put(titleMap.get(res), k);
					}
				}
			}
		}

		// �����������е�������
		for (int i = 0; i <= rowMax; i++) {
			row = sheet1.getRow(i);
			int maxColumn = row.getPhysicalNumberOfCells();
			// ���һ�е���������0
			if (maxColumn > 0) {
				// ����һ����¼
				Excel2VcfBean excelBean = new Excel2VcfBean();
				// ����
				String name = null;
				// �ֻ�����
				String tel = null;
				// Ӧ���绰
				String emergencyTel = null;
				// ��˾
				String company = "��������";
				// �����ص�
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
				if (null != name && null != tel && !"����".equals(name)
						&& !"�ֻ�����".equals(tel) && !"�����ص�".equals(tel)) {
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
	 * ����propertiesΪMap
	 * properties�д�ŵ���ʵ���õ���excel������ֶ�ӳ�䣬���ݲ�ͬ����޸Ķ�Ӧ��name��
	 * 
	 * @param filePath ��ȡproperties��ȫ����Ϣ
	 * @return �������ɵ�Map
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
	 * ���ݵ�Ԫ�����ݸ�ʽ�Ĳ�ͬ���ز�ͬ���ַ��� 
	 * 
	 * @param cell ��Ԫ��
	 * @return ת������ַ���
	 */
	private static String whichCellType(Cell cell) {
		// ���ص��ַ���
		String result = null;
		if (null != cell) {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:// String���͵�Ԫ��
				// �ı�����
				result = cell.getRichStringCellValue().getString();
				break;
			case Cell.CELL_TYPE_NUMERIC:// ��������
				// ��鵥Ԫ���Ƿ����һ��Date����
				if (DateUtil.isCellDateFormatted(cell)) {
					// ��������
					result = cell.getDateCellValue().toString();
				} else {
					// ��������
					// ��1.3031113083E10תΪ13031113083
					DecimalFormat df = new DecimalFormat("#");
					result = df.format(cell.getNumericCellValue());
				}
				break;
			}
		}
		return result;
	}

	/**
	 * дvcf������
	 * @param excelList д�������
	 * @param vcfPath vcf��·��
	 * @return �Ƿ�ɹ�ִ��
	 */

	public static Boolean WriteVcf(List<Excel2VcfBean> excelList, String vcfPath) {
		// ��ʶ�Ƿ�ɹ�ִ��
		boolean ok = false;
		File file = new File(vcfPath);
		if (file.exists()) {
			file.delete();
		}
		try {
			file.createNewFile();
			FileOutputStream fos = new FileOutputStream(vcfPath, false);
			// ָ�������ʽΪutf-8
			OutputStreamWriter osw = new OutputStreamWriter(fos, "utf-8");
			BufferedWriter bw = new BufferedWriter(osw);

			for (Excel2VcfBean excelBean : excelList) {
				// bw.write(a.toString());
				StringBuffer sb = new StringBuffer();
				// BEGIN:VCARD
				// VERSION:3.0
				// N:;;;;
				// FN:����
				// TEL;TYPE=CELL:123456789
				// ORG:�ӱ���ҵ ʯ��ׯ
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
		String res = (ok == true) ? "�ɹ�" : "ʧ��";
		System.out.println("--��ɣ�ִ�н��Ϊ--:" + res);
	}
}
