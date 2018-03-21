package LichThi;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.List;

public class Main {

	public static void main (String []args){
		in_output read=new in_output();
		try{
			FileInputStream b = new FileInputStream("KetQuaDangKyMonHoc.xlsx");
			FileInputStream c = new FileInputStream("DanhSachPhongThi.xlsx");
			read.readfile_KQDKMH(b);
			read.readfile_DSPT(c);
			b.close();
			c.close();
			read.writefile();
			Graph G =new Graph();
			G.createGraph(read.getDsMonHocLT());
			List<PhongThi> dsPTLT = read.getDsPhongLT();
			long totalCapacity = 0;
			for (PhongThi phongThi : dsPTLT) {
				totalCapacity += phongThi.getSucChua();
			}
			G.toMau(7, 10, totalCapacity, 10,true);
			RoomSchedule schedule= new RoomSchedule();
			
			schedule.schedule(G.getTimeSlot(), dsPTLT, 5, 10);
			
			G.createGraph(read.getDsMonHocTH());
			
			List<PhongThi> dsPTTH = read.getDsPhongTH();
			totalCapacity = 0;
			for (PhongThi phongThi : dsPTTH) {
				totalCapacity += phongThi.getSucChua();
			}
			G.toMau(7, 10, totalCapacity, 10, false);
			
			schedule.schedule(G.getTimeSlot(), dsPTTH, 5, 10);
			schedule.writeToExcel(new File("LichThi.xlsx"));
			schedule.writeDetailToExcel(new File("DetailExam.xlsx"));
			//read.writeDSMH(new File("Danh Sach Mon Hoc.xlsx"));
		}
		catch(IOException e){
			e.printStackTrace();
			
		}
	}
}
