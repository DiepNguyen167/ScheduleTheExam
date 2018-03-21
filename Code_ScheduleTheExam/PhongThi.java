package LichThi;

import java.util.ArrayList;
import java.util.List;

class PhongThi {
	private final String maPhong;
	private final int sucChua;
	private String loai;
	private String ghiChu;
	private List<SinhVien> dsSVDuThi; 
	
	public PhongThi(String maPhong, int sucChua, String loai, String ghiChu) {
		this.maPhong = maPhong;
		this.sucChua = sucChua;
		this.loai = loai;
		this.ghiChu = ghiChu;
		dsSVDuThi= new ArrayList<SinhVien>();
	}

	public List<SinhVien> getDsSVDuThi() {
		return dsSVDuThi;
	}
	
	public String getMaPhong() {
		return maPhong;
	}

	public int getSucChua() {
		return sucChua;
	}

	public String getLoai() {
		return loai;
	}

	public String getGhiChu() {
		return ghiChu;
	}

	@Override
	public boolean equals(Object obj) {
		PhongThi another = (PhongThi) obj;
		return this.getMaPhong().equals(another.getMaPhong());
	}

	@Override
	public String toString() {
		if (maPhong == null || loai == null || ghiChu == null) {
			return "No room";
		} else {
			return maPhong + " - " + loai + " - " + sucChua + "\r\n";
		}
	}

}