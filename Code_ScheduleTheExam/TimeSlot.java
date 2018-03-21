package LichThi;

import java.util.ArrayList;
import java.util.List;

public class TimeSlot {
	private int day; // you can change to datetime for clearer
	private int slot;
	private long capacity;
	private long capacityRoom;
	private long totalCapacity;
	private List<MonHoc> listOfCourses;
	private List<ArrayList<PhongThi>> listOfRooms; // mon nao thuoc phong nao
													// (mon bao gom nhieu nhom)
	private List<ArrayList<ArrayList<SinhVien>>> studentSlots; // sinh vine nao
																// thuoc phong
																// nao

	public TimeSlot(int day, int slot, long capacity) {
		this.day = day;
		this.slot = slot;
		this.capacity = capacity;
		this.totalCapacity = capacity;
		this.capacityRoom = capacity;
		listOfCourses = new ArrayList<MonHoc>();
		listOfRooms = new ArrayList<ArrayList<PhongThi>>();
		studentSlots = new ArrayList<ArrayList<ArrayList<SinhVien>>>();
	}

	public void updateCapacity(long capacity) {
		this.capacity = capacity;
		this.totalCapacity = capacity;
	}

	public void addCourse(MonHoc course) {
		listOfCourses.add(course);
		capacity -= course.getTatCaSV().size();
		listOfRooms.add(new ArrayList<PhongThi>());
		studentSlots.add(new ArrayList<ArrayList<SinhVien>>());
	}

	public long getTotalCapacity() {
		return totalCapacity;
	}

	public List<ArrayList<PhongThi>> getListOfRooms() {
		return listOfRooms;
	}

	public List<PhongThi> getAllRooms() {
		List<PhongThi> allRooms = new ArrayList<>();
		for (List<PhongThi> subList : listOfRooms) {
			allRooms.addAll(subList);
		}
		return allRooms;
	}

	public List<ArrayList<ArrayList<SinhVien>>> getStudentSlots() {
		return studentSlots;
	}

	public int getDay() {
		return day;
	}

	public int getSlot() {
		return slot;
	}

	public List<MonHoc> getListOfCourses() {
		return listOfCourses;
	}

	public List<MonHoc> getListOfUnscheduledCourses() {
		List<MonHoc> listOfUnscheduled = new ArrayList<MonHoc>();
		for (int i = 0; i < listOfCourses.size(); i++) {
			if (listOfRooms.get(i).size() == 0) {
				listOfUnscheduled.add(listOfCourses.get(i));
			}
		}
		return listOfUnscheduled;
	}

	public long getCapacity() {
		return capacity;
	}

	public long getCapatiyRoom() {
		return capacityRoom;
	}

	public boolean canContain(long requestCapacity, int threshold) {
		if (requestCapacity >= capacity) {
			return false;
		} else {
			double percent = (double) (capacity - requestCapacity) / totalCapacity;
			percent *= 100;
			return percent >= threshold;
		}
	}

	@Override
	public boolean equals(Object obj) {
		TimeSlot another = (TimeSlot) obj;
		return (this.getDay() == another.getDay()) && (this.getSlot() == another.getSlot());
	}

	public boolean checkExistedStudent(SinhVien student) {
		for (MonHoc monHoc : listOfCourses) {
			if (monHoc.checkExistedStudent(student))
				return true;
		}
		return false;
	}

	public void setRoom(MonHoc course, PhongThi room) {
		// System.out.println("Schedule for " + course );
		int indexCourse = listOfCourses.indexOf(course);
		if (indexCourse > -1) {
			listOfRooms.get(indexCourse).add(room);
			capacityRoom -= room.getSucChua();
			studentSlots.get(indexCourse).add(new ArrayList<SinhVien>());
		}
	}

	public void setStudent(MonHoc course, PhongThi room, SinhVien student) {
		int indexCourse = listOfCourses.indexOf(course);
		if (indexCourse > -1) {
			int indexRoom = listOfRooms.get(indexCourse).indexOf(room);
			if (indexRoom > -1) {
				studentSlots.get(indexCourse).get(indexRoom).add(student);
			}
		}
	}

	public void setAllStudents(MonHoc course, PhongThi room, List<SinhVien> students) {
		int indexCourse = listOfCourses.indexOf(course);
		if (indexCourse > -1) {
			int indexRoom = listOfRooms.get(indexCourse).indexOf(room);
			if (indexRoom > -1) {
				studentSlots.get(indexCourse).get(indexRoom).addAll(students);
			}
		}
	}

	public void removeCourse(MonHoc currentCourse) {
		if (listOfCourses.remove(currentCourse)) {
			capacity += currentCourse.getSoLuongSV();
		}
	}
}
