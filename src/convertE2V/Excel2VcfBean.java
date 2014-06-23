package convertE2V;

/**
 * 
 * @author liuxiuquan
 * excel转vcf用的实体类
 *
 */
public class Excel2VcfBean {
	//姓名
	private String name;
	//手机号码
	private String tel;
	//应急电话
	private String TelEme;
	//公司
	private String company;
	
	//重写toString方法 方便打印输出
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
