package LichThi;

import java.io.PrintStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;

public class Graph {
	private List<Vertex> listOfVertexs;
	private List<Edge> listOfEdges;
	private List<TimeSlot> listOfTimeSlots;
	private List<MonHoc> listOfCourses;
	private int numberOfDays;
	private int numberOfSlots;
	private int threshold;

	public List<Vertex> getListOfVertexs() {
		return listOfVertexs;
	}

	public List<Edge> getListOfEdges() {
		return listOfEdges;
	}

	public List<MonHoc> getListOfCourses() {
		return listOfCourses;
	}
	
	public List<TimeSlot> getTimeSlot() {
		return listOfTimeSlots;
	}
	
	public Graph() {
	}

	public void createGraph(List<MonHoc> dsMonHoc) {
		listOfCourses = dsMonHoc;
		listOfEdges = new ArrayList<>();
		listOfVertexs = new ArrayList<>();
		for (int i = 0; i < dsMonHoc.size() - 1; i++) {
			MonHoc firstMH = dsMonHoc.get(i);
			Vertex firstVertex = new Vertex(firstMH.getMaMH());
			int index = listOfVertexs.indexOf(firstVertex);
			if (index < 0) {
				listOfVertexs.add(firstVertex);
			} else {
				firstVertex = listOfVertexs.get(index);
			}
			for (int j = i + 1; j < dsMonHoc.size(); j++) {
				MonHoc secondMH = dsMonHoc.get(j);
				Vertex secondVertex = new Vertex(secondMH.getMaMH());
				index = listOfVertexs.indexOf(secondVertex);
				if (index < 0) {
					listOfVertexs.add(secondVertex);
				} else {
					secondVertex = listOfVertexs.get(index);
				}

				int counter = 0;
				for (SinhVien firstSV : firstMH.getTatCaSV()) {
						if (secondMH.checkExistedStudent(firstSV)) {
							counter++;
						}
				}
				if (counter > 0) {
					Edge newEdge = new Edge(firstVertex.getId() + "-" + secondVertex.getId(), counter);
					newEdge.addVertex(firstVertex);
					newEdge.addVertex(secondVertex);
					firstVertex.addEdge(newEdge);
					secondVertex.addEdge(newEdge);
					listOfEdges.add(newEdge);
				}
			}
		}
	}

	private void addCourseToSlot(int slotKey, MonHoc course) {
		listOfTimeSlots.get(slotKey).addCourse(course);
	}

	private MonHoc getCourseFromId(String Id) {
		for (MonHoc course : listOfCourses) {
			if (course.getMaMH().equals(Id))
				return course;
		}
		return null;
	}

	public boolean toMau(int numberOfDays, int numberOfSlots, long concurrencyLevel, int threshold,
			boolean resetTimeSlot) {
		this.numberOfDays = numberOfDays;
		this.numberOfSlots = numberOfSlots;
		this.threshold = threshold;
		sortVertexs(listOfVertexs);// sort vexter have high->nhỏ degree.
//		int total = 0;
//		for (Vertex vertex : listOfVertexs) {
//			MonHoc mh = getCourseFromId(vertex.getId()); // take last vertex
//		}
		
		if (resetTimeSlot || listOfTimeSlots == null) { 
			listOfTimeSlots = new ArrayList<>();
			for (int indexDay = 0; indexDay < numberOfDays; indexDay++) {
				for (int indexSlot = 0; indexSlot < numberOfSlots; indexSlot++) {
					TimeSlot newSlot = new TimeSlot(indexDay, indexSlot, concurrencyLevel);
					listOfTimeSlots.add(newSlot);
				}
			} //create list day and slot .
		} else {
			for (int i = 0; i < listOfTimeSlots.size();i++) {
				listOfTimeSlots.get(i).updateCapacity(concurrencyLevel); //set sum capacity for this slot.
			}
		}
		int numberOfColoredVertexs = 0;
		if (!listOfVertexs.isEmpty()) {
			int slotKey = 0 * numberOfSlots + 0;
			listOfVertexs.get(0).setColor(slotKey); // set color first vertex.
			numberOfColoredVertexs++; // increase number color.
			MonHoc course = getCourseFromId(listOfVertexs.get(0).getId());
			if (null != course)
				addCourseToSlot(slotKey, course); // add first vertex in this slot.

		}
		for (int indexVertex = 0; indexVertex < listOfVertexs.size(); indexVertex++) {
			if (numberOfColoredVertexs >= listOfVertexs.size()) {
				return true; // numberColor > number of vertex => finish.
			}
			Vertex currentVertex = listOfVertexs.get(indexVertex);
			if (currentVertex.getColor() == Vertex.NOT_COLORED) {
				int availableColor = getNearestAvailableColor(currentVertex);//take color
				if (availableColor >= 0) {
					currentVertex.setColor(availableColor);
					MonHoc course = getCourseFromId(currentVertex.getId());
					if (null != course)
						addCourseToSlot(availableColor, course);
					numberOfColoredVertexs++;

					// coloring neighbours
					List<Vertex> listOfNeighbours = new ArrayList<>();
					listOfNeighbours.addAll(currentVertex.getNeighbours());
					sortVertexs(listOfNeighbours);
					for (int indexNeighbour = 0; indexNeighbour < listOfNeighbours.size(); indexNeighbour++) {
						Vertex currentNeighbour = listOfNeighbours.get(indexNeighbour);
						if (currentNeighbour.getColor() == Vertex.NOT_COLORED) {
							int availableColorForNeighbour = getNearestAvailableColor(currentNeighbour);
							if (availableColorForNeighbour >= 0) {
								currentNeighbour.setColor(availableColorForNeighbour);
								numberOfColoredVertexs++;
								course = getCourseFromId(currentNeighbour.getId());
								if (null != course)
									addCourseToSlot(availableColorForNeighbour, course);
							}
						} else {
							// Not implemented
						}
					}
				}
			} else {
				// Not implemented
			}
		}
		return numberOfColoredVertexs == listOfVertexs.size();
	}

	private int getNearestAvailableColor(Vertex currentVertex) {
		// check this slot can contain allSTUsent of this course .
		List<Vertex> listOfNeighbours = currentVertex.getNeighbours();
		for (int indexDay = 0; indexDay < numberOfDays; indexDay++) {
			for (int indexSlot = 0; indexSlot < numberOfSlots; indexSlot++) {
				// check other conditions before checking neighbours
				// because if it dont same as neghbours,
				// but other conditions are failed
				// this color also cant use
				boolean scheduled = true;
				int slotKey = indexDay * numberOfSlots + indexSlot;
				if (!listOfTimeSlots.get(slotKey).canContain(getCourseFromId(currentVertex.getId()).getTatCaSV().size(),
						threshold)) {
					scheduled = false;
					continue;
				} else {
					// Not Implemented
				}

				if (!checkConstraints(currentVertex, slotKey)) {

					scheduled = false;
					continue;
				} else {
					// Not Implemented
				}
// check this course with it neighbour have same day or not same  and same color?.
				for (int indexNeighbour = 0; indexNeighbour < listOfNeighbours.size(); indexNeighbour++) {
					Vertex neighbour = listOfNeighbours.get(indexNeighbour);
					int neighbourColor = neighbour.getColor();
					if (neighbourColor != Vertex.NOT_COLORED) {
						int selectedColor = indexDay * numberOfSlots + indexSlot;
						if (selectedColor != neighbourColor) { // Not same color
							int internalDistance = Math.abs(listOfTimeSlots.get(neighbourColor).getSlot()
									- listOfTimeSlots.get(selectedColor).getSlot());//check slot
							int externalDistance = Math.abs(listOfTimeSlots.get(neighbourColor).getDay()
									- listOfTimeSlots.get(selectedColor).getDay());//check day.
							if (externalDistance == 0) { // same day
								if (internalDistance <= 1) {
									scheduled = false;
									break;
								}
							} else { // if not same day
								// Just wait for the result
							}
						} else { // Same color
							scheduled = false;
							break;
						}
					}
				}
				if (scheduled) { //not same color and not same slot in one day.
					return slotKey;
				}
			}
		}
		return -1; // error code
	}

	private boolean checkConstraints(Vertex currentVertex, int slotKey) {
		// At this function, you can add a number of other constraints as you
		// like
		// TimeSlot per day for one student
		MonHoc course = getCourseFromId(currentVertex.getId());
		for (SinhVien student : course.getTatCaSV()) {
			if (listOfTimeSlots.get(slotKey).checkExistedStudent(student)) {
				return false;
			}
		}

		// most 3 slot per day for one student
		int currentDay = slotKey / numberOfSlots;
		int minSlot = currentDay * numberOfSlots;
		int maxSlot = minSlot  + numberOfSlots - 1;
		int counter = 0;
		for (SinhVien student : course.getTatCaSV()) {
			for(int slot = minSlot; slot <= maxSlot; slot++) {
				if (listOfTimeSlots.get(slot).checkExistedStudent(student)) {
					counter++;
				}
			}
			if (counter > 3) {
				return false;
			}
		}
		return true;
	}

	private void sortVertexs(List<Vertex> sortedList) {
		Collections.sort(sortedList, new Comparator<Vertex>() {

			@Override
			public int compare(Vertex o1, Vertex o2) {
				int value = o2.getWeight() - o1.getWeight();
				if (value == 0) {
					return o2.getLargestEdgeWeight() - o1.getLargestEdgeWeight();
				} else {
					return value;
				}
			}
		});

	}

	public void writeGraph(PrintStream outStream) {
		for (Vertex vertex : listOfVertexs) {
			outStream.println(vertex.getId() + ": " + String.valueOf(vertex.getColor()));
			for (Edge edge : vertex.getListOfEdges()) {
				Vertex neighbor = edge.getTheOtherVertex(vertex);
				outStream.print("\t" + neighbor.getId() + ": " + String.valueOf(neighbor.getColor()));
				outStream.println();
				// String.valueOf(neighbor.getColor()));
			}
		}
		// sortPhongThi();
		outStream.println("TimeSlot");
		for (TimeSlot timeSlot : listOfTimeSlots) {
			outStream.println("Day: " + timeSlot.getDay() + " Slot: " + timeSlot.getSlot());
			for (MonHoc monHoc : timeSlot.getListOfCourses()) {
				outStream.print("\t" + monHoc);
			}
			outStream.println();
		}
	}
}
