package LichThi;

import java.io.*;
import java.util.*;
import java.lang.*;
import java.lang.reflect.Array;

import org.apache.poi.ss.usermodel.*;
import org.apache.log4j.chainsaw.Main;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.xsdschema.impl.NarrowMaxMinImpl.MaxOccursImpl;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class in_output {

	private String Input;
	private String Output;
	private ArrayList<MonHoc> dsMonHoc;
	private ArrayList<SinhVien> dsSinhVien;
	private ArrayList<PhongThi> dsPhong;
	private ArrayList<PhongThi> dsPhongLT;
	private ArrayList<PhongThi> dsPhongTH;
	private ArrayList<MonHoc> dsMonHocLT;
	private ArrayList<MonHoc> dsMonHocTH;
	private ArrayList<MonHoc> dsMonHocKhongThi;

	private int Height = 800;

	private String[] name;

	public void setI_FileName(String filename) {
		this.Input = filename;
	}

	public void setO_FileName(String filename) {
		this.Output = filename;
	}

	public void setHeight(int values) {
		this.Height = values;
	}

	public HSSFCellStyle createStyleForTitle(HSSFWorkbook output) {
		HSSFFont font = output.createFont();
		font.setItalic(true);
		HSSFCellStyle style = output.createCellStyle();
		style.setFont(font);
		return style;
	}

	public void writefile() throws IOException {
		HSSFWorkbook outputexcel = new HSSFWorkbook();
		HSSFSheet sheet = outputexcel.createSheet("LichThi.xlsx");
		HSSFRow row = sheet.createRow((short) 0);
		HSSFCellStyle style = outputexcel.createCellStyle();
		style.setWrapText(true);
		row.setRowStyle(style);
		Cell cell;
		// IdMH
		cell = row.createCell(0, CellType.STRING);
		cell.setCellValue("Mã Môn");
		cell.setCellStyle(style);
		// EmpName
		cell = row.createCell(1, CellType.STRING);
		cell.setCellValue("Tên Môn");
		cell.setCellStyle(style);
		// IdNhom
		cell = row.createCell(2, CellType.STRING);
		cell.setCellValue("Mã Nhóm");
		cell.setCellStyle(style);
		// IdTo
		cell = row.createCell(3, CellType.STRING);
		cell.setCellValue("Mã Tổ");
		cell.setCellStyle(style);
		// NgayThi
		cell = row.createCell(4, CellType.STRING);
		cell.setCellValue("Ngày Thi");
		cell.setCellStyle(style);
		//
		cell = row.createCell(5, CellType.STRING);
		cell.setCellValue("Ca Thi");
		cell.setCellStyle(style);
		//
		cell = row.createCell(6, CellType.STRING);
		cell.setCellValue("Phòng Thi");
		cell.setCellStyle(style);
		//
		cell = row.createCell(7, CellType.STRING);
		cell.setCellValue("sức chứa");
		cell.setCellStyle(style);
		//
		cell = row.createCell(8, CellType.STRING);
		cell.setCellValue("Sỉ Số");
		cell.setCellStyle(style);
		//
		cell = row.createCell(9, CellType.STRING);
		cell.setCellValue("Loại Phòng");
		cell.setCellStyle(style);
		// row.getCell(0).setCellStyle(style);
		File file = new File("LichThi.xlsx");
		FileOutputStream outFile = new FileOutputStream(file);
		outputexcel.write(outFile);
		outputexcel.close();
	}

	public void writeDSMH(File file) throws IOException {
		XSSFWorkbook outputexcel = new XSSFWorkbook();
		XSSFSheet sheet = outputexcel.createSheet("Danh Sach Mon Hoc");
		FileOutputStream outFile;
		try {
			outFile = new FileOutputStream(file);
			XSSFRow row = sheet.createRow(0);
			XSSFCellStyle borderStyle = outputexcel.createCellStyle();
			XSSFFont boldFont = outputexcel.createFont();
			boldFont.setBold(true);
			borderStyle.setWrapText(true);
			borderStyle.setBorderBottom(BorderStyle.THIN);
			borderStyle.setBorderLeft(BorderStyle.THIN);
			borderStyle.setBorderRight(BorderStyle.THIN);
			borderStyle.setBorderTop(BorderStyle.THIN);
			borderStyle.setAlignment(HorizontalAlignment.RIGHT);

			// row.setRowStyle(borderStyle);
			Cell cell;
			// IdMH
			cell = row.createCell(0, CellType.STRING);
			XSSFRichTextString rts = new XSSFRichTextString("Mã Môn");
			rts.applyFont(boldFont);
			cell.setCellValue(rts);
			cell.setCellStyle(borderStyle);
			// EmpName
			cell = row.createCell(1, CellType.STRING);
			rts = new XSSFRichTextString("Tên Môn");
			rts.applyFont(boldFont);
			cell.setCellValue(rts);
			cell.setCellStyle(borderStyle);
			int indexRow = 1;
			// if(false)
			for (MonHoc monHoc : dsMonHoc) {
				XSSFRow rowTime = sheet.createRow(indexRow++);
				rowTime.setRowStyle(borderStyle);
				Cell cellTime = rowTime.createCell(0, CellType.STRING);
				cellTime.setCellValue(monHoc.getMaMH());
				cellTime.setCellStyle(borderStyle);
				// System.out.print(timeSlot.getListOfCourses().get(i).getMaMH()
				// + "\t");
				// EmpName
				cellTime = rowTime.createCell(1, CellType.STRING);
				cellTime.setCellValue(monHoc.getTenMH());
				// System.out.print(timeSlot.getListOfCourses().get(i).getTenMH()
				// + "\t");
				cellTime.setCellStyle(borderStyle);
			}
			outputexcel.write(outFile);
			outputexcel.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public void addPhong(String loai, PhongThi P) {
		String name1 = "PM";
		String name2 = "LT";
		if (loai.equals(name1)) {
			dsPhongTH.add(P);
		} else if (loai.equals(name2)) {
			dsPhongLT.add(P);
		}

	}

	public ArrayList<PhongThi> getDsPhongLT() {
		return dsPhongLT;
	}

	public ArrayList<PhongThi> getDsPhongTH() {
		return dsPhongTH;
	}

	public void readfile_DSPT(FileInputStream input1) throws IOException {
		dsPhong = new ArrayList<>();
		dsPhongLT = new ArrayList<>();
		dsPhongTH = new ArrayList<>();
		dsSinhVien = new ArrayList<>();
		XSSFWorkbook DSPT = new XSSFWorkbook(input1);
		XSSFSheet sheetData = DSPT.getSheetAt(0);
		FormulaEvaluator formula = DSPT.getCreationHelper().createFormulaEvaluator();

		Iterator<Row> rowIterator = sheetData.iterator();
		rowIterator.next();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			if (row.getCell(0).toString() != "/n" || row.getCell(0).toString() != "/r") {
				if ((row.getCell(3) == null) || (row.getCell(3).getCellType() == Cell.CELL_TYPE_BLANK)) {
					row.createCell(3, Cell.CELL_TYPE_STRING);
					row.getCell(3).setCellValue("");
				}
				Cell maPhong = row.getCell(0);
				Cell sucChuaStr = row.getCell(1);
				int sucChua = Integer.parseInt(sucChuaStr.toString());
				Cell loai = row.getCell(2);
				Cell ghiChu = row.getCell(3);

				PhongThi phongThi = new PhongThi(maPhong.toString(), sucChua, loai.toString(), ghiChu.toString());
				dsPhong.add(phongThi);
				addPhong(loai.toString(), phongThi);
			}
			// System.out.println(ghiChu.toString());
		}
		DSPT.close();
		sortPhongThi(dsPhongTH);
		sortPhongThi(dsPhongLT);
	}

	private Integer tryParseInt(String numStr) {
		try {
			Integer result = Integer.parseInt(numStr);
			return result;
		} catch (Exception ex) {
			return null;
		}
	}

	private void sortPhongThi(List<PhongThi> sortedList) {
		Collections.sort(sortedList, new Comparator<PhongThi>() {

			@Override
			public int compare(PhongThi arg0, PhongThi arg1) {
				// TODO Auto-generated method stub
				int sizeP1 = arg0.getSucChua();
				int sizeP2 = arg1.getSucChua();
				if (sizeP1 > sizeP2) {
					return 1;
				} else if (sizeP1 == sizeP2) {
					return 0;
				} else if (sizeP1 < sizeP2) {
					return -1;
				}

				return 0;
			}

		});
	}

	public void readfile_KQDKMH(FileInputStream input1) throws IOException {

		XSSFWorkbook KQDKMH = new XSSFWorkbook(input1);
		XSSFSheet sheetData = KQDKMH.getSheetAt(0);
		FormulaEvaluator formula = KQDKMH.getCreationHelper().createFormulaEvaluator();
		dsMonHocLT = new ArrayList<MonHoc>();
		dsMonHocTH = new ArrayList<MonHoc>();
		dsSinhVien = new ArrayList<SinhVien>();
		dsMonHoc = new ArrayList<MonHoc>();
		dsMonHocKhongThi = new ArrayList<MonHoc>();
		Iterator<Row> rowIterator = sheetData.iterator();
		rowIterator.next();

		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			if (row.getCell(0).toString() != "/n" || row.getCell(0).toString() != "/r") {
				if ((row.getCell(0).toString().equals("500002")) || (row.getCell(0).toString().equals("500003"))) {
					int indexCell = -1;
					for (Cell cell : row) {
						indexCell++;
						if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
							cell = row.createCell(indexCell);
							cell.setCellType(Cell.CELL_TYPE_STRING);
							cell.setCellValue("");
						}
					}
					String maMH = row.getCell(0).toString();
					String tenMH = row.getCell(1).toString();
					MonHoc monHoc = new MonHoc(maMH, tenMH);

					dsMonHocKhongThi.add(monHoc);

				} else {
					int indexCell = -1;
					for (Cell cell : row) {
						indexCell++;
						if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
							cell = row.createCell(indexCell);
							cell.setCellType(Cell.CELL_TYPE_STRING);
							cell.setCellValue("");
						}
					}
					String maMH = row.getCell(0).toString();
					String tenMH = row.getCell(1).toString();
					String nhomStr = row.getCell(2).toString();
					int nhom = Integer.parseInt(nhomStr);
					String toTHStr = row.getCell(3).toString();
					int toTH;
					if (tryParseInt(toTHStr) != null) {
						toTH = tryParseInt(toTHStr);
					} else {
						toTH = -1;
					}

					String maLop = row.getCell(4).toString();
					String mssv = row.getCell(5).toString();
					String hoLot = row.getCell(6).toString();
					String tenSV = row.getCell(7).toString();
					String email = row.getCell(8).toString();
					// System.out.println(toTH);
					MonHoc monHoc = new MonHoc(maMH, tenMH);
					SinhVien sinhVien = new SinhVien(mssv, maLop, hoLot, tenSV, email);
					dsSinhVien.add(sinhVien);
					sinhVien.themMon(monHoc);
					int index = dsMonHoc.indexOf(monHoc);
					if (index < 0) {
						dsMonHoc.add(monHoc);
						if (toTH == -1) {
							dsMonHocLT.add(monHoc);
						} else {
							dsMonHocTH.add(monHoc);
						}
						monHoc.themSV(sinhVien, nhom, toTH);
					} else {
						dsMonHoc.get(index).themSV(sinhVien, nhom, toTH);
					}
				}
			}
		}
		KQDKMH.close();
	}

	public String getInput() {
		return Input;
	}

	public String getOutput() {
		return Output;
	}

	public ArrayList<MonHoc> getDsMonHoc() {
		return dsMonHoc;
	}

	public ArrayList<SinhVien> getDsSinhVien() {
		return dsSinhVien;
	}

	public ArrayList<PhongThi> getDsPhong() {
		return dsPhong;
	}

	public ArrayList<MonHoc> getDsMonHocLT() {
		return dsMonHocLT;
	}

	public ArrayList<MonHoc> getDsMonHocTH() {
		return dsMonHocTH;
	}

	public int getHeight() {
		return Height;
	}

	public String[] getName() {
		return name;
	}

	public List<MonHoc> getDSMH() {
		return dsMonHoc;
	}
}
