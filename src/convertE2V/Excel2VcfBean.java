package convertE2V;

/**
 * 
 * @author liuxiuquan
 * excelתvcf�õ�ʵ����
 *
 */
public class Excel2VcfBean {
	//����
	private String name;
	//�ֻ�����
	private String tel;
	//Ӧ���绰
	private String TelEme;
	//��˾
	private String company;
	
	//��дtoString���� �����ӡ���
	public String toString(){
		return "name:"+name+",tel:"+tel+",telEme:"+TelEme+",company:"+company;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getTel() {
		return tel;
	}
	public void setTel(String tel) {
		this.tel = tel;
	}
	public String getTelEme() {
		return TelEme;
	}
	public void setTelEme(String telEme) {
		this.TelEme = telEme;
	}

	public String getCompany() {
		return company;
	}
	public void setCompany(String company) {
		this.company = company;
	}
	
}
