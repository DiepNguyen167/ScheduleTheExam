package LichThi;

import java.util.ArrayList;
import java.util.List;

public class Vertex {
	private String id;
	private List<Edge> listOfEdges; // danh sach canh ke
	private int color;
	private int weight;
	public static final int NOT_COLORED = -1;
	
	public Vertex(String id) {
		this.id = id;
		listOfEdges = new ArrayList<>();
		color = NOT_COLORED; // Don't be colored
		weight = 0; //degree
	}

	public String getId() {
		return id;
	}

	public List<Edge> getListOfEdges() {
		return listOfEdges;
	}
	
	public void addEdge(Edge newEgde) {
		if (!listOfEdges.contains(newEgde)) {
			listOfEdges.add(newEgde);
			weight++;
		}
	}
	
	public int getColor() {
		return color;
	}
	
	public void setColor(int color) {
		this.color = color;
	}
	
	public int getWeight() {
		return weight;
	}
	
	public int getLargestEdgeWeight() {
		int max = -1;
		for (int i = 0; i < listOfEdges.size(); i++) {
			int weight = listOfEdges.get(i).getWeight();
			if (weight > max) {
				max = weight;
			}
		}
		return max;
	}
	
	public List<Vertex> getNeighbours() {
		List<Vertex> listOfNeighbours = new ArrayList<>();
		for (Edge edge : listOfEdges) {
			listOfNeighbours.add(edge.getTheOtherVertex(this));
		}
		return listOfNeighbours;
	}
	
	@Override
	public boolean equals(Object obj) {
		Vertex another = (Vertex) obj;
		return this.getId().equals(another.getId());
	}
	
	public String toString() {
		return String.valueOf(weight) + " - " + String.valueOf(getLargestEdgeWeight()) + "\r\n";
	}
}
