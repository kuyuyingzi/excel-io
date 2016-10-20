package example.imp;

import java.util.Date;

public class Member {

	private String name;
	private Integer age;
	private Integer country;
	private Date date;
	private String dateDesc;
	public Member() {
		super();
	}
	public Member(String name, Integer age, Integer country, String dateDesc) {
		super();
		this.name = name;
		this.age = age;
		this.country = country;
		this.dateDesc = dateDesc;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public Integer getAge() {
		return age;
	}
	public void setAge(Integer age) {
		this.age = age;
	}
	public Integer getCountry() {
		return country;
	}
	public void setCountry(Integer country) {
		this.country = country;
	}
	public Date getDate() {
		return date;
	}
	public void setDate(Date date) {
		this.date = date;
	}
	public String getDateDesc() {
		return dateDesc;
	}
	public void setDateDesc(String dateDesc) {
		this.dateDesc = dateDesc;
	}
}
