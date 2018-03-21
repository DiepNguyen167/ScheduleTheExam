package LichThi;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RoomSchedule {
	private List<TimeSlot> listOfTimeSlot;
	private List<PhongThi> listOfRooms;
	

	List<PhongThi> listOfUsedRooms;
	List<PhongThi> listOfRemainRooms;
	int minIndex;
	int maxIndex;

	public boolean schedule(List<TimeSlot> listOfRefTimeSlot, List<PhongThi> listOfRefRooms, int threshold,
			int threshold2) {
		this.listOfTimeSlot = listOfRefTimeSlot;
		this.listOfRooms = listOfRefRooms;
		reoderRoom(listOfRefRooms);
		List<MonHoc> fuckedCourses = new ArrayList<MonHoc>();
		for (int indexTimeSlot = 0; indexTimeSlot < listOfTimeSlot.size(); indexTimeSlot++) {
			TimeSlot currentTimeSlot = listOfTimeSlot.get(indexTimeSlot);
			List<MonHoc> listOfUnwantedCourse = new ArrayList<MonHoc>();
			List<MonHoc> listOfCourse = currentTimeSlot.getListOfUnscheduledCourses();
			listOfUsedRooms = new ArrayList<PhongThi>();
			listOfRemainRooms = new ArrayList<PhongThi>();
			listOfRemainRooms.addAll(listOfRooms);
			reoderCourse(listOfCourse);
			listOfUsedRooms.addAll(currentTimeSlot.getAllRooms());
			listOfRemainRooms.removeAll(listOfUsedRooms);
			minIndex = 0;
			maxIndex = listOfRooms.size() - 1;
			List<MonHoc> listOfOptionTwo = new ArrayList<MonHoc>();
			for (int indexCourse = 0; indexCourse < listOfCourse.size(); indexCourse++) {
				MonHoc currentCourse = listOfCourse.get(indexCourse);
				List<PhongThi> listOfScheduledRooms = scheduleOptionalOne(currentCourse, threshold);
				if (listOfScheduledRooms.size() > 0) {
					for (int indexNhom = 0; indexNhom < currentCourse.getdsNhomMH().size(); indexNhom++) {
						listOfUsedRooms.add(listOfScheduledRooms.get(indexNhom));
						listOfRemainRooms.remove(listOfScheduledRooms.get(indexNhom));
						currentTimeSlot.setRoom(currentCourse, listOfScheduledRooms.get(indexNhom));
						currentTimeSlot.setAllStudents(currentCourse, listOfScheduledRooms.get(indexNhom),
								currentCourse.getdsNhomMH().get(indexNhom).getDsDangKy());
						if (indexNhom == minIndex) {
							minIndex++;
						} else if (indexNhom == maxIndex) {
							maxIndex--;
						}
					}
				} else {
					listOfOptionTwo.add(currentCourse);//ds luu mon hoc k săp dc o gdoan 1
				}
			}

			reoderCourse(listOfOptionTwo);
			for (int indexCourse = 0; indexCourse < listOfOptionTwo.size(); indexCourse++) {
				MonHoc currentCourse = listOfOptionTwo.get(indexCourse);
				List<PhongThi> listOfScheduledRooms = scheduleOptionalTwo(currentCourse, threshold2);
				if (listOfScheduledRooms.size() > 0) {
					int indexNhom = 0;
					int indexStudent = 0;
					for (int indexRoom = 0; indexRoom < listOfScheduledRooms.size(); indexRoom++) {
						if (indexRoom == minIndex) {// giữ 2 đầu ds phong => giam do phuc tap
							minIndex++;
						} else if (indexRoom == maxIndex) {
							maxIndex--;
						}
						listOfUsedRooms.add(listOfScheduledRooms.get(indexRoom));
						listOfRemainRooms.remove(listOfScheduledRooms.get(indexRoom));
						currentTimeSlot.setRoom(currentCourse, listOfScheduledRooms.get(indexRoom));
						int counter = 0;
						long max = listOfScheduledRooms.get(indexRoom).getSucChua();
						List<NhomMonHoc> listOfSubCourses = currentCourse.getdsNhomMH();
						while (counter < max) {
							if (indexStudent < listOfSubCourses.get(indexNhom).getDsDangKy().size()) {
								currentTimeSlot.setStudent(currentCourse, listOfScheduledRooms.get(indexRoom),
										listOfSubCourses.get(indexNhom).getDsDangKy().get(indexStudent));
								indexStudent++;
								counter++;
							} else {
								if (indexNhom < listOfSubCourses.size() - 1) {
									indexNhom++;
									indexStudent = 0;
								} else {
									indexNhom++;
									break;
								}
							}
						}

						if (indexNhom >= listOfSubCourses.size()) {
							break;
						}
					}
				} else {
					listOfUnwantedCourse.add(currentCourse);//
				}
			}

			reoderCourse(listOfUnwantedCourse);
			List<MonHoc> listOfScheduledCourses = new ArrayList<MonHoc>();
			for (int indexCourse = 0; indexCourse < listOfUnwantedCourse.size(); indexCourse++) {
				MonHoc currentCourse = listOfUnwantedCourse.get(indexCourse);
				List<PhongThi> listOfScheduledRooms = scheduleOptionalTwo(currentCourse, threshold2 + 5);
				if (listOfScheduledRooms.size() > 0) {
					int indexNhom = 0;
					int indexStudent = 0;
					for (int indexRoom = 0; indexRoom < listOfScheduledRooms.size(); indexRoom++) {
						if (indexRoom == minIndex) {
							minIndex++;
						} else if (indexRoom == maxIndex) {
							maxIndex--;
						}
						listOfUsedRooms.add(listOfScheduledRooms.get(indexRoom));
						listOfRemainRooms.remove(listOfScheduledRooms.get(indexRoom));
						currentTimeSlot.setRoom(currentCourse, listOfScheduledRooms.get(indexRoom));
						int counter = 0;
						long max = listOfScheduledRooms.get(indexRoom).getSucChua();
						List<NhomMonHoc> listOfSubCourses = currentCourse.getdsNhomMH();
						while (counter < max) {
							if (indexStudent < listOfSubCourses.get(indexNhom).getDsDangKy().size()) {
								currentTimeSlot.setStudent(currentCourse, listOfScheduledRooms.get(indexRoom),
										listOfSubCourses.get(indexNhom).getDsDangKy().get(indexStudent));
							} else {
								if (indexNhom < listOfSubCourses.size() - 1) {
									indexNhom++;
									indexStudent = 0;
								} else {
									indexNhom++;
									break;
								}
							}
							indexStudent++;
							counter++;
						}

						if (indexNhom >= listOfSubCourses.size()) {
							break;
						}
					}
					listOfScheduledCourses.add(currentCourse);
				} else {

				}
			}
			listOfUnwantedCourse.removeAll(listOfScheduledCourses);

			listOfScheduledCourses.clear();
			for (int indexCourse = 0; indexCourse < listOfUnwantedCourse.size(); indexCourse++) {
				MonHoc currentCourse = listOfUnwantedCourse.get(indexCourse);
				List<PhongThi> listOfScheduledRooms = scheduleOptionalTwo(currentCourse, threshold2 + 10);
				if (listOfScheduledRooms.size() > 0) {
					int indexNhom = 0;
					int indexStudent = 0;
					for (int indexRoom = 0; indexRoom < listOfScheduledRooms.size(); indexRoom++) {
						if (indexRoom == minIndex) {
							minIndex++;
						} else if (indexRoom == maxIndex) {
							maxIndex--;
						}
						listOfUsedRooms.add(listOfScheduledRooms.get(indexRoom));
						listOfRemainRooms.remove(listOfScheduledRooms.get(indexRoom));
						currentTimeSlot.setRoom(currentCourse, listOfScheduledRooms.get(indexRoom));
						int counter = 0;
						long max = listOfScheduledRooms.get(indexRoom).getSucChua();
						List<NhomMonHoc> listOfSubCourses = currentCourse.getdsNhomMH();
						while (counter < max) {
							if (indexStudent < listOfSubCourses.get(indexNhom).getDsDangKy().size()) {
								currentTimeSlot.setStudent(currentCourse, listOfScheduledRooms.get(indexRoom),
										listOfSubCourses.get(indexNhom).getDsDangKy().get(indexStudent));
							} else {
								if (indexNhom < listOfSubCourses.size() - 1) {
									indexNhom++;
									indexStudent = 0;
								} else {
									indexNhom++;
									break;
								}
							}
							indexStudent++;
							counter++;
						}

						if (indexNhom >= listOfSubCourses.size()) {
							break;
						}
					}
					listOfScheduledCourses.add(currentCourse);
				} else {

				}
			}
			listOfUnwantedCourse.removeAll(listOfScheduledCourses);

			listOfScheduledCourses.clear();
			for (int indexCourse = 0; indexCourse < listOfUnwantedCourse.size(); indexCourse++) {
				MonHoc currentCourse = listOfUnwantedCourse.get(indexCourse);
				List<PhongThi> listOfScheduledRooms = scheduleOptionalTwo(currentCourse, threshold2 + 15);
				if (listOfScheduledRooms.size() > 0) {
					int indexNhom = 0;
					int indexStudent = 0;
					for (int indexRoom = 0; indexRoom < listOfScheduledRooms.size(); indexRoom++) {
						if (indexRoom == minIndex) {
							minIndex++;
						} else if (indexRoom == maxIndex) {
							maxIndex--;
						}
						listOfUsedRooms.add(listOfScheduledRooms.get(indexRoom));
						listOfRemainRooms.remove(listOfScheduledRooms.get(indexRoom));
						currentTimeSlot.setRoom(currentCourse, listOfScheduledRooms.get(indexRoom));
						int counter = 0;
						long max = listOfScheduledRooms.get(indexRoom).getSucChua();
						List<NhomMonHoc> listOfSubCourses = currentCourse.getdsNhomMH();
						while (counter < max) {
							if (indexStudent < listOfSubCourses.get(indexNhom).getDsDangKy().size()) {
								currentTimeSlot.setStudent(currentCourse, listOfScheduledRooms.get(indexRoom),
										listOfSubCourses.get(indexNhom).getDsDangKy().get(indexStudent));
							} else {
								if (indexNhom < listOfSubCourses.size() - 1) {
									indexNhom++;
									indexStudent = 0;
								} else {
									indexNhom++;
									break;
								}
							}
							indexStudent++;
							counter++;
						}

						if (indexNhom >= listOfSubCourses.size()) {
							break;
						}
					}
					listOfScheduledCourses.add(currentCourse);
				} else {

				}
			}
			listOfUnwantedCourse.removeAll(listOfScheduledCourses);

			listOfScheduledCourses.clear();
			for (int indexCourse = 0; indexCourse < listOfUnwantedCourse.size(); indexCourse++) {
				MonHoc currentCourse = listOfUnwantedCourse.get(indexCourse);
				if (currentCourse.getSoLuongSV() < currentTimeSlot.getCapatiyRoom()) {
					List<PhongThi> listOfScheduledRooms = scheduleOptionalThree(currentCourse);
					if (listOfScheduledRooms.size() > 0) {
						int indexNhom = 0;
						int indexStudent = 0;
						for (int indexRoom = 0; indexRoom < listOfScheduledRooms.size(); indexRoom++) {
							if (indexRoom == minIndex) {
								minIndex++;
							} else if (indexRoom == maxIndex) {
								maxIndex--;
							}
							listOfUsedRooms.add(listOfScheduledRooms.get(indexRoom));
							listOfRemainRooms.remove(listOfScheduledRooms.get(indexRoom));
							currentTimeSlot.setRoom(currentCourse, listOfScheduledRooms.get(indexRoom));
							int counter = 0;
							long max = listOfScheduledRooms.get(indexRoom).getSucChua();
							List<NhomMonHoc> listOfSubCourses = currentCourse.getdsNhomMH();
							while (counter < max) {
								if (indexStudent < listOfSubCourses.get(indexNhom).getDsDangKy().size()) {
									currentTimeSlot.setStudent(currentCourse, listOfScheduledRooms.get(indexRoom),
											listOfSubCourses.get(indexNhom).getDsDangKy().get(indexStudent));
								} else {
									if (indexNhom < listOfSubCourses.size() - 1) {
										indexNhom++;
										indexStudent = 0;
									} else {
										indexNhom++;
										break;
									}
								}
								indexStudent++;
								counter++;
							}

							if (indexNhom >= listOfSubCourses.size()) {
								break;
							}
						}
						listOfScheduledCourses.add(currentCourse);
					} else {
						currentTimeSlot.removeCourse(currentCourse);
					}
				} else {
					currentTimeSlot.removeCourse(currentCourse);
				}
			}
			listOfUnwantedCourse.removeAll(listOfScheduledCourses);
			fuckedCourses.addAll(listOfUnwantedCourse);
		}
		
		for (MonHoc fuckedCourse : fuckedCourses) {
			for (TimeSlot timeSlot : listOfTimeSlot) {
				minIndex = 0;
				maxIndex = listOfRooms.size() - 1;
				if (fuckedCourse.getSoLuongSV() < timeSlot.getCapatiyRoom()) {
					List<PhongThi> listOfScheduledRooms = scheduleOptionalThree(fuckedCourse);
					if (listOfScheduledRooms.size() > 0) {
						int indexNhom = 0;
						int indexStudent = 0;
						timeSlot.addCourse(fuckedCourse);
						for (int indexRoom = 0; indexRoom < listOfScheduledRooms.size(); indexRoom++) {
							listOfUsedRooms.add(listOfScheduledRooms.get(indexRoom));
							listOfRemainRooms.remove(listOfScheduledRooms.get(indexRoom));
							timeSlot.setRoom(fuckedCourse, listOfScheduledRooms.get(indexRoom));
							int counter = 0;
							long max = listOfScheduledRooms.get(indexRoom).getSucChua();
							List<NhomMonHoc> listOfSubCourses = fuckedCourse.getdsNhomMH();
							while (counter < max) {
								if (indexStudent < listOfSubCourses.get(indexNhom).getDsDangKy().size()) {
									timeSlot.setStudent(fuckedCourse, listOfScheduledRooms.get(indexRoom),
											listOfSubCourses.get(indexNhom).getDsDangKy().get(indexStudent));
								} else {
									if (indexNhom < listOfSubCourses.size() - 1) {
										indexNhom++;
										indexStudent = 0;
									} else {
										indexNhom++;
										break;
									}
								}
								indexStudent++;
								counter++;
							}

							if (indexNhom >= listOfSubCourses.size()) {
								break;
							}
						}
						fuckedCourses.remove(fuckedCourse);
					} 
				}
			}
		}
		if (fuckedCourses.size() > 0) {
			System.out.println("Không đủ phòng sắp cho các Môn: ");
			for (MonHoc monHoc : fuckedCourses) {
				System.out.println(monHoc);
			}
		}
		return true;

	}

	// support function
	private void reoderRoom(List<PhongThi> listOfRooms) {
		Collections.sort(listOfRooms, new Comparator<PhongThi>() {

			@Override
			public int compare(PhongThi o1, PhongThi o2) {
				return o2.getSucChua() - o1.getSucChua();
			}
		});
	}

	private void reoderCourse(List<MonHoc> listOfCourse) {
		Collections.sort(listOfCourse, new Comparator<MonHoc>() {

			@Override
			public int compare(MonHoc o1, MonHoc o2) {
				return o2.getSoLuongSV() - o1.getSoLuongSV();
			}
		});
	}

	// handling function
	private List<PhongThi> scheduleOptionalOne(MonHoc course, int threshold) {
		List<PhongThi> result = new ArrayList<PhongThi>();
		List<NhomMonHoc> listOfSections = course.getdsNhomMH();
		int beginIndex = minIndex;
		int endIndex = maxIndex;
		if (1 == listOfSections.size()) {//có 1 nhóm thì chia ra làm đôi để tìm phòng.
			int mid = binarySearch(course.getSoLuongSV(), beginIndex, endIndex);
			while (listOfUsedRooms.contains(listOfRooms.get(mid))
					|| listOfRooms.get(mid).getSucChua() < course.getSoLuongSV()) {
				if (beginIndex == mid) {
					return result;
				}
				mid--;
			}
			float balanceValue = listOfRooms.get(mid).getSucChua() - course.getSoLuongSV();
			if (balanceValue < 0) {
				return result;
			} else if (balanceValue <= threshold)
				result.add(listOfRooms.get(mid));
			return result;
		} else {
			for (NhomMonHoc section : listOfSections) {
				int mid = binarySearch(section.getDsDangKy().size(), beginIndex, endIndex);
				while ((listOfUsedRooms.contains(listOfRooms.get(mid)) || result.contains(mid))
						|| listOfRooms.get(mid).getSucChua() < section.getDsDangKy().size()) {
					if (beginIndex == mid) {
						result.clear();
						return result;
					}
					mid--;
				}
				float balanceValue = listOfRooms.get(mid).getSucChua() - course.getSoLuongSV();
				if (balanceValue < 0 || balanceValue > threshold) {
					result.clear();
					return result;
				}
				result.add(listOfRooms.get(mid));
			}
			return result;
		}
	}

	private List<PhongThi> scheduleOptionalTwo(MonHoc course, int threshold2) {
		List<PhongThi> listOfScheduledRooms = new ArrayList<PhongThi>();
		if (checkingTable(course.getSoLuongSV(), threshold2, listOfScheduledRooms)) {

		}
		return listOfScheduledRooms;
	}

	private List<PhongThi> scheduleOptionalThree(MonHoc course) {
		int total = course.getSoLuongSV();
		int beginIndex = minIndex;
		int endIndex = maxIndex;
		List<PhongThi> listOfScheduledRooms = new ArrayList<PhongThi>();
		while (beginIndex <= maxIndex && listOfRooms.get(beginIndex).getSucChua() < total) {
			if (!listOfUsedRooms.contains(listOfRooms.get(beginIndex))) {
				listOfScheduledRooms.add(listOfRooms.get(beginIndex));
				total -= listOfRooms.get(beginIndex).getSucChua();
			}
			beginIndex++;
		}
		if (beginIndex > endIndex) {
			listOfScheduledRooms.clear();
			return listOfScheduledRooms;
		}
		if (total > 0) {
			int mid = binarySearch(total, beginIndex, endIndex);
			while (listOfUsedRooms.contains(listOfRooms.get(mid)) || listOfRooms.get(mid).getSucChua() < total) {
				if (beginIndex == mid) {
					listOfScheduledRooms.clear();
					return listOfScheduledRooms;
				}
				mid--;
			}
			listOfScheduledRooms.add(listOfRooms.get(mid));
		}
		return listOfScheduledRooms;
	}

	// sort function

	private int binarySearch(int target, int beginIndex, int endIndex) {
		while (beginIndex < endIndex) {
			if (beginIndex == endIndex || endIndex == (beginIndex + 1)) {
				return endIndex;
			}
			int mid = (beginIndex + endIndex) / 2;
			if (target < listOfRooms.get(mid).getSucChua()) {
				beginIndex = mid;
				// return binarySearch(target, beginIndex, mid);
			} else if (target > listOfRooms.get(mid).getSucChua()) {
				endIndex = mid;
				// return binarySearch(target, mid, endIndex);
			} else {
				return mid;
			}
		}
		return endIndex;
	}

	private boolean checkingTable(int target, int threshold2, List<PhongThi> result) {
		int size = listOfRemainRooms.size();
		long[][] checkingTable = new long[size][size];
		for (int indexRow = 0; indexRow < size; indexRow++) {
			for (int indexCol = 0; indexCol < size; indexCol++) {//create table
				if (indexRow == indexCol) {
					checkingTable[indexRow][indexCol] = 0;
				} else if (0 == indexCol) {//first col 
					checkingTable[indexRow][indexCol] = listOfRemainRooms.get(indexRow).getSucChua()
							+ listOfRemainRooms.get(indexCol).getSucChua();
					long balance = checkingTable[indexRow][indexCol] - target;
					if (balance > 0) {
						if (balance <= threshold2) {
							result.add(listOfRemainRooms.get(indexRow));
							for (int i = 0; i <= indexCol; i++) {
								if (i != indexRow) {
									result.add(listOfRemainRooms.get(i));
								}
							}

						}
						return true;
					}
				} else {
					if (indexRow != indexCol - 1) {
						checkingTable[indexRow][indexCol] = checkingTable[indexRow][indexCol - 1]
								+ listOfRemainRooms.get(indexCol).getSucChua();
					} else if ((indexCol - 1) == 0) {
						checkingTable[indexRow][indexCol] = listOfRemainRooms.get(indexRow).getSucChua()
								+ listOfRemainRooms.get(indexCol).getSucChua();
					} else {
						checkingTable[indexRow][indexCol] = checkingTable[indexRow][indexCol - 2]
								+ listOfRemainRooms.get(indexCol).getSucChua();
					}
					long balance = checkingTable[indexRow][indexCol] - target;
					if (balance > 0) {
						if (balance <= threshold2) {
							result.add(listOfRemainRooms.get(indexRow));
							for (int i = 0; i <= indexCol; i++) {
								if (i != indexRow) {
									result.add(listOfRemainRooms.get(i));
								}
							}
							return true;
						}
					}
				}
			}
		}
		return false;
	}

	public void write(PrintStream outStream) {
		outStream.println("TimeSlot");
		for (TimeSlot timeSlot : listOfTimeSlot) {
			outStream.println("Day: " + timeSlot.getDay() + " Slot: " + timeSlot.getSlot() + " Avaliable: "
					+ timeSlot.getCapatiyRoom());
			for (int i = 0; i < timeSlot.getListOfCourses().size(); i++) {
				outStream.print("\t" + timeSlot.getListOfCourses().get(i).getMaMH() + " ("
						+ String.valueOf(timeSlot.getListOfCourses().get(i).getSoLuongSV()) + ")" + " - ");
				int total = 0;
				for (int j = 0; j < timeSlot.getListOfRooms().get(i).size(); j++) {
					outStream.print(timeSlot.getListOfRooms().get(i).get(j).getMaPhong() + " ("
							+ timeSlot.getListOfRooms().get(i).get(j).getSucChua() + ")" + " ");
					total += timeSlot.getListOfRooms().get(i).get(j).getSucChua();
				}
				outStream.print(" - Empty: " + (total - timeSlot.getListOfCourses().get(i).getSoLuongSV()));
				outStream.println();
			}
			outStream.println();
		}
	}

	public void writeToExcel(File file) {
		XSSFWorkbook outputexcel = new XSSFWorkbook();
		XSSFSheet sheet = outputexcel.createSheet("LichThi");
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
			// IdNhom
			cell = row.createCell(2, CellType.STRING);
			rts = new XSSFRichTextString("Mã Nhóm");
			rts.applyFont(boldFont);
			cell.setCellValue(rts);
			cell.setCellStyle(borderStyle);
			// IdTo
			cell = row.createCell(3, CellType.STRING);
			rts = new XSSFRichTextString("Mã Tổ");
			rts.applyFont(boldFont);
			cell.setCellValue(rts);
			cell.setCellStyle(borderStyle);
			// NgayThi
			cell = row.createCell(4, CellType.STRING);
			rts = new XSSFRichTextString("Ngày Thi");
			rts.applyFont(boldFont);
			cell.setCellValue(rts);
			cell.setCellStyle(borderStyle);
			//
			cell = row.createCell(5, CellType.STRING);
			rts = new XSSFRichTextString("Ca Thi");
			rts.applyFont(boldFont);
			cell.setCellValue(rts);
			cell.setCellStyle(borderStyle);
			//
			cell = row.createCell(6, CellType.STRING);
			rts = new XSSFRichTextString("Phòng Thi");
			rts.applyFont(boldFont);
			cell.setCellValue(rts);
			cell.setCellStyle(borderStyle);
			//
			cell = row.createCell(7, CellType.STRING);
			rts = new XSSFRichTextString("Sức Chứa");
			rts.applyFont(boldFont);
			cell.setCellValue(rts);
			cell.setCellStyle(borderStyle);
			//
			cell = row.createCell(8, CellType.STRING);
			rts = new XSSFRichTextString("Sỉ Số");
			rts.applyFont(boldFont);
			cell.setCellValue(rts);
			cell.setCellStyle(borderStyle);
			//
			cell = row.createCell(9, CellType.STRING);
			rts = new XSSFRichTextString("Loại Phòng");
			rts.applyFont(boldFont);
			cell.setCellValue(rts);
			cell.setCellStyle(borderStyle);
			// row.getCell(0).setCellStyle(borderStyle);

			// render values
			int indexRow = 1;
			// if(false)
			for (TimeSlot timeSlot : listOfTimeSlot) {
				for (int i = 0; i < timeSlot.getListOfCourses().size(); i++) {
					int total = 0;
					int totalRoom = 0;
					int indexRoom = 0;
					totalRoom += timeSlot.getListOfRooms().get(i).get(indexRoom).getSucChua();
					for (int j = 0; j < timeSlot.getListOfCourses().get(i).getdsNhomMH().size(); j++) {
						XSSFRow rowTime = sheet.createRow(indexRow++);
						rowTime.setRowStyle(borderStyle);
						total += timeSlot.getListOfCourses().get(i).getdsNhomMH().get(j).getDsDangKy().size();

						Cell cellTime = rowTime.createCell(0, CellType.STRING);
						cellTime.setCellValue(timeSlot.getListOfCourses().get(i).getMaMH());
						cellTime.setCellStyle(borderStyle);
						//System.out.print(timeSlot.getListOfCourses().get(i).getMaMH() + "\t");
						// EmpName
						cellTime = rowTime.createCell(1, CellType.STRING);
						cellTime.setCellValue(timeSlot.getListOfCourses().get(i).getTenMH());
						//System.out.print(timeSlot.getListOfCourses().get(i).getTenMH() + "\t");
						cellTime.setCellStyle(borderStyle);
						// IdNhom
						cellTime = rowTime.createCell(2, CellType.STRING);
						cellTime.setCellValue(timeSlot.getListOfCourses().get(i).getdsNhomMH().get(j).getNhom());
						//System.out.print(timeSlot.getListOfCourses().get(i).getdsNhomMH().get(j).getNhom() + "\t");
						cellTime.setCellStyle(borderStyle);
						// IdTo
						cellTime = rowTime.createCell(3, CellType.STRING);
						cellTime.setCellValue((timeSlot.getListOfCourses().get(i).getdsNhomMH().get(j).getToTH() > -1)
								? String.valueOf(timeSlot.getListOfCourses().get(i).getdsNhomMH().get(j).getToTH())
								: " ");
						cellTime.setCellStyle(borderStyle);
						// NgayThi
						cellTime = rowTime.createCell(4, CellType.STRING);
						cellTime.setCellValue(timeSlot.getDay() + 1);
						//System.out.print((timeSlot.getDay() + 1) + "\t");
						cellTime.setCellStyle(borderStyle);
						//
						cellTime = rowTime.createCell(5, CellType.STRING);
						cellTime.setCellValue(timeSlot.getSlot() + 1);
						//System.out.print((timeSlot.getSlot() + 1) + "\t");
						cellTime.setCellStyle(borderStyle);
						//
						cellTime = rowTime.createCell(6, CellType.STRING);
						cellTime.setCellValue(timeSlot.getListOfRooms().get(i).get(indexRoom).getMaPhong());
						//System.out.print(timeSlot.getListOfRooms().get(i).get(indexRoom).getMaPhong() + "\t");
						cellTime.setCellStyle(borderStyle);

						//
						cellTime = rowTime.createCell(7, CellType.STRING);
						cellTime.setCellValue(timeSlot.getListOfRooms().get(i).get(indexRoom).getSucChua());
						//System.out.print(timeSlot.getListOfRooms().get(i).get(indexRoom).getSucChua() + "\t");
						cellTime.setCellStyle(borderStyle);
						//
						cellTime = rowTime.createCell(8, CellType.STRING);
						cellTime.setCellValue(timeSlot.getStudentSlots().get(i).get(indexRoom).size());
						//System.out.print(timeSlot.getStudentSlots().get(i).get(indexRoom).size() + "\t");
						cellTime.setCellStyle(borderStyle);
						//
						cellTime = rowTime.createCell(9, CellType.STRING);
						cellTime.setCellValue(timeSlot.getListOfRooms().get(i).get(indexRoom).getLoai());
						//System.out.print(timeSlot.getListOfRooms().get(i).get(indexRoom).getLoai() + "\t");
						cellTime.setCellStyle(borderStyle);

						while ((totalRoom - total) < 0) {
							indexRoom++;
							System.out.println();
							XSSFRow rowRoom = sheet.createRow(indexRow++);
							Cell cellRoom = rowRoom.createCell(6, CellType.STRING);
							cellRoom.setCellValue(timeSlot.getListOfRooms().get(i).get(indexRoom).getMaPhong());
							//System.out.print("\t\t\t\t\t\t\t"
								//	+ timeSlot.getListOfRooms().get(i).get(indexRoom).getMaPhong() + "\t");
							cellRoom.setCellStyle(borderStyle);

							//
							cellRoom = rowRoom.createCell(7, CellType.STRING);
							cellRoom.setCellValue(timeSlot.getListOfRooms().get(i).get(indexRoom).getSucChua());
							//System.out.print(timeSlot.getListOfRooms().get(i).get(indexRoom).getSucChua() + "\t");
							cellRoom.setCellStyle(borderStyle);
							//
							cellRoom = rowRoom.createCell(8, CellType.STRING);
							cellRoom.setCellValue(timeSlot.getStudentSlots().get(i).get(indexRoom).size());
							//System.out.print(timeSlot.getStudentSlots().get(i).get(indexRoom).size() + "\t");
							cellRoom.setCellStyle(borderStyle);
							//
							cellRoom = rowRoom.createCell(9, CellType.STRING);
							cellRoom.setCellValue(timeSlot.getListOfRooms().get(i).get(indexRoom).getLoai());
							//System.out.print(timeSlot.getListOfRooms().get(i).get(indexRoom).getLoai() + "\t");
							cellRoom.setCellStyle(borderStyle);
							totalRoom += timeSlot.getListOfRooms().get(i).get(indexRoom).getSucChua();
						}
						if (totalRoom == total) {
							indexRoom++;

							if (indexRoom < timeSlot.getListOfRooms().get(i).size())
								totalRoom += timeSlot.getListOfRooms().get(i).get(indexRoom).getSucChua();
						}
						System.out.println();
					}
				}
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

	public void writeDetailToExcel(File file) {
		XSSFWorkbook workBook = new XSSFWorkbook();
		for (TimeSlot timeSlot : listOfTimeSlot) {
			for (int indexCourse = 0; indexCourse < timeSlot.getListOfCourses().size(); indexCourse++) {
				MonHoc course = timeSlot.getListOfCourses().get(indexCourse);
				XSSFSheet sheet = workBook.createSheet(course.getMaMH());
				
				XSSFCellStyle borderStyle = workBook.createCellStyle();
				XSSFFont boldFont = workBook.createFont();
				boldFont.setBold(true);
				borderStyle.setWrapText(true);
				borderStyle.setBorderBottom(BorderStyle.THIN);
				borderStyle.setBorderLeft(BorderStyle.THIN);
				borderStyle.setBorderRight(BorderStyle.THIN);
				borderStyle.setBorderTop(BorderStyle.THIN);
				borderStyle.setAlignment(HorizontalAlignment.RIGHT);

				XSSFRow headerRow = sheet.createRow(0);
				Cell headerCell = headerRow.createCell(0, CellType.STRING);
				XSSFRichTextString rts = new XSSFRichTextString("MSSV");
				rts.applyFont(boldFont);
				headerCell.setCellValue(rts);
				headerCell.setCellStyle(borderStyle);

				headerCell = headerRow.createCell(1, CellType.STRING);
				rts = new XSSFRichTextString("Lớp");
				rts.applyFont(boldFont);
				headerCell.setCellValue(rts);
				headerCell.setCellStyle(borderStyle);

				headerCell = headerRow.createCell(2, CellType.STRING);
				rts = new XSSFRichTextString("Họ Lót");
				rts.applyFont(boldFont);
				headerCell.setCellValue(rts);
				headerCell.setCellStyle(borderStyle);

				headerCell = headerRow.createCell(3, CellType.STRING);
				rts = new XSSFRichTextString("Tên");
				rts.applyFont(boldFont);
				headerCell.setCellValue(rts);
				headerCell.setCellStyle(borderStyle);
				
				headerCell = headerRow.createCell(4, CellType.STRING);
				rts = new XSSFRichTextString("Môn học");
				rts.applyFont(boldFont);
				headerCell.setCellValue(rts);
				headerCell.setCellStyle(borderStyle);
				

				headerCell = headerRow.createCell(5, CellType.STRING);
				rts = new XSSFRichTextString("Ngày Thi");
				rts.applyFont(boldFont);
				headerCell.setCellValue(rts);
				headerCell.setCellStyle(borderStyle);

				headerCell = headerRow.createCell(6, CellType.STRING);
				rts = new XSSFRichTextString("Ca Thi");
				rts.applyFont(boldFont);
				headerCell.setCellValue(rts);
				headerCell.setCellStyle(borderStyle);

				headerCell = headerRow.createCell(7, CellType.STRING);
				rts = new XSSFRichTextString("Phòng");
				rts.applyFont(boldFont);
				headerCell.setCellValue(rts);
				headerCell.setCellStyle(borderStyle);

				int indexRow = 1;
				for (SinhVien student : course.getTatCaSV()) {
					int indexRoom = -1;
					for (indexRoom = 0; indexRoom < timeSlot.getStudentSlots().get(indexCourse).size(); indexRoom++) {
						if (timeSlot.getStudentSlots().get(indexCourse).get(indexRoom).contains(student)) {
							break;
						}
					}
					XSSFRow contentRow = sheet.createRow(indexRow++);
					Cell conentCell = contentRow.createCell(0, CellType.STRING);
					conentCell.setCellValue(student.getMSSV());
					conentCell.setCellStyle(borderStyle);

					conentCell = contentRow.createCell(1, CellType.STRING);
					conentCell.setCellValue(student.getMaLop());
					conentCell.setCellStyle(borderStyle);

					conentCell = contentRow.createCell(2, CellType.STRING);
					conentCell.setCellValue(student.getHo());
					conentCell.setCellStyle(borderStyle);

					conentCell = contentRow.createCell(3, CellType.STRING);
					conentCell.setCellValue(student.getTen());
					conentCell.setCellStyle(borderStyle);
					
					conentCell = contentRow.createCell(4, CellType.STRING);
					conentCell.setCellValue(course.getTenMH());
					
					conentCell.setCellStyle(borderStyle);

					conentCell = contentRow.createCell(5, CellType.STRING);
					conentCell.setCellValue(timeSlot.getDay() + 1);
					conentCell.setCellStyle(borderStyle);

					conentCell = contentRow.createCell(6, CellType.STRING);
					conentCell.setCellValue(timeSlot.getSlot() + 1);
					conentCell.setCellStyle(borderStyle);

					conentCell = contentRow.createCell(7, CellType.STRING);
					conentCell.setCellValue(listOfRooms.get(indexRoom).getMaPhong());
					conentCell.setCellStyle(borderStyle);
				}

			}
		}

		FileOutputStream outStream;
		try {
			outStream = new FileOutputStream(file);
			workBook.write(outStream);
			workBook.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
