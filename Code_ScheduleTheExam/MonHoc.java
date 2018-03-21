package LichThi;

import java.util.ArrayList;
import java.util.List;

class MonHoc{
	private String maMH;
	private String tenMH;
	private List<NhomMonHoc> dsNhomMH;
	private List<PhongThi> dsPhongThi;
	
	public MonHoc(String maMH, String tenMH) {
		this.maMH = maMH;
		this.tenMH = tenMH;
		dsNhomMH = new ArrayList<NhomMonHoc>();
		dsPhongThi= new ArrayList<PhongThi>();
	}

	public List<PhongThi> getPhongThi() {
		return dsPhongThi;
	}
	
	public String getMaMH() {
		return maMH;
	}
	
	public String getTenMH() {
		return tenMH;
	}
	
	public List<NhomMonHoc> getdsNhomMH(){
		return dsNhomMH;
	}
	
	public void themSV(SinhVien sinhVien, int nhom, int to) {
		boolean checker = true;
		for (NhomMonHoc nhomMonHoc : dsNhomMH) {
			if (nhomMonHoc.getNhom() == nhom && nhomMonHoc.getToTH() == to) {
				nhomMonHoc.themSV(sinhVien);
				nhomMonHoc.themTo(to);
				for(Group group: nhomMonHoc.getDsTo()){
					Group toMH=new Group(to);
					if(toMH==group){
						group.themSV(sinhVien);
					}
				}
				checker = false;
				break;
			}
		}
		if (checker) {
			NhomMonHoc nhomMoi = new NhomMonHoc(nhom, to);
			nhomMoi.themSV(sinhVien);
			dsNhomMH.add(nhomMoi);
			nhomMoi.themTo(to);
			for(Group group: nhomMoi.getDsTo()){
				Group toMH=new Group(to);
				if(toMH==group){
					group.themSV(sinhVien);
				}
			}
		}
	}
	
	public List<SinhVien> getTatCaSV() {
		List<SinhVien> allStudents = new ArrayList<>();
		for (NhomMonHoc nhomMonHoc : dsNhomMH) {
			allStudents.addAll(nhomMonHoc.getdsDangKy());
		}
		return allStudents;
	}
	
	public int getSoLuongSV() {
		return getTatCaSV().size();
	}
	
	@Override
	public boolean equals(Object obj) {
		MonHoc another = (MonHoc) obj;
		return this.getMaMH().equals(another.getMaMH());
	}
	
	public boolean checkExistedStudent(SinhVien student) {
		for (NhomMonHoc nhomMonHoc : dsNhomMH) {
			if (nhomMonHoc.checkExistedStudent(student)) {
				return true;
			}
		}
		return false;
	}
	
	@Override
	public String toString() {
		return  ((dsNhomMH.get(0).getToTH()>-1)?"TH":"") + maMH + " - " + tenMH + "\r\n";
	}
}