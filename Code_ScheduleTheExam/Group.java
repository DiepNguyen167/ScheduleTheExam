package LichThi;

import java.util.ArrayList;
import java.util.List;

public class Group {
	public int IDgroup;
	public List<SinhVien> DsSVTheoTo;
	
	public Group (int group) {
		this.IDgroup=group;
		this.DsSVTheoTo=new ArrayList<SinhVien>();
	}

	public int getIDGroup() {
		return IDgroup;
	}

	public List<SinhVien> getDsSVTheoTo() {
		return DsSVTheoTo;
	}
	
	public void themSV(SinhVien sinhVien) {
		
		if ((!DsSVTheoTo.contains(sinhVien))) {
			DsSVTheoTo.add(sinhVien);
		}
		
	}
	
	public boolean checkExistedStudent(SinhVien student) {
		return DsSVTheoTo.contains(student);
	}
}
