package LichThi;

import java.util.ArrayList;
import java.util.List;

public class SinhVien {
	private String MSSV;
	private String maLop;
	private String ho;
	private String ten;
	private String email;
	private List<MonHoc> dsDangKy;
	
	public SinhVien(String MSSV, String maLop, String ho, String ten, String email) {
		this.MSSV = MSSV;
		this.maLop = maLop;
		this.ho = ho;
		this.ten = ten;
		this.email = email;
		dsDangKy = new ArrayList<MonHoc>();
	}
	
	public String getMSSV() {
		return MSSV;
	}
	public String getMaLop() {
		return maLop;
	}
	public String getHo() {
		return ho;
	}
	public String getTen() {
		return ten;
	}
	public String getEmail() {
		return email;
	}
	public List<MonHoc> getDsDangKy() {
		return dsDangKy;
	}
	
	public void themMon(MonHoc monHoc) {
		if(!dsDangKy.contains(monHoc)) {
			dsDangKy.add(monHoc);
		}
	}
	
	@Override
	public boolean equals(Object obj) {
		SinhVien another = (SinhVien)obj;
		return this.getMSSV().equals(another.getMSSV());
	}
	
	@Override
	public String toString() {
		String print = MSSV;
		print += (" - " + maLop + " - " + ho + " " + ten + " - " + email + "/r/n");
		print += "DSMHDK: /r/n";
		for(int i=0;i<dsDangKy.size();i++){
			print += (dsDangKy.get(i) + "/r/n");
		}
		return print;
	}
}
