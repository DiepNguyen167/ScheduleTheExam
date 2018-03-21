package LichThi;

import java.util.ArrayList;
import java.util.List;

public class NhomMonHoc {
	private int nhom;
	private int toTH;
	private List<Group> dsTo;
	private List<SinhVien> dsDangKy;//sv dki theo nhom
	
	public NhomMonHoc(int nhom, int toTH) {
		this.nhom = nhom;
		this.toTH = toTH;
		dsTo    = new ArrayList<Group>();
		dsDangKy  = new ArrayList<SinhVien>();
	}
	
	public int getNhom() {
		return nhom;
	}
	
	public int getToTH() {
		return toTH;
	}
	
	public List<SinhVien> getdsDangKy(){
		return dsDangKy;
	}
	
	public void themTo( int to) {
		boolean checker = true;
		for (Group group : dsTo) {
			if (group.getIDGroup() == to) {
				checker = false;
				break;
			}
		}
		if (checker) {
			Group group=new Group(to);
			dsTo.add(group);
		}
	}

	public int soluongSV() {
		int count=0;
		for(int i=0;i<dsDangKy.size();i++){
			count++;
		}
		return count;
	}
	public String getName() {
		String print = "Nhóm " + String.valueOf(nhom);
		if (toTH > -1) {
			print += ("  - Tổ TH " + String.valueOf(toTH));
		}
		return print;
	}
	
	
	public List<Group> getDsTo() {
		return dsTo;
	}

	public List<SinhVien> getDsDangKy() {
		return dsDangKy;
	}

	public void themSV(SinhVien sinhVien) {
	
		if ((!dsDangKy.contains(sinhVien))) {
			dsDangKy.add(sinhVien);
		}
		
	}
	
		
	public boolean checkExistedStudent(SinhVien student) {
		return dsDangKy.contains(student);
	}
}
