package LichThi;

import java.util.ArrayList;
import java.util.List;

public class Edge {
	private String id;
	private List<Vertex> listOfVertexs; // danh sach canh ke
	private int weight;
	
	public Edge(String id, int weight) {
		this.id = id;
		this.weight = weight;
		listOfVertexs = new ArrayList<>(); //  co the set cung la 2 dinh 
	}

	public String getId() {
		return id;
	}

	public List<Vertex> getListOfVertexs() {
		return listOfVertexs;
	}
	
	public void addVertex(Vertex newVertex) {
		if (!listOfVertexs.contains(newVertex)) {
			listOfVertexs.add(newVertex);
		}
	}
	
	public int getWeight() {
		return weight;
	}
	
	
	public Vertex getTheOtherVertex(Vertex oneVertex) {
		if (listOfVertexs.get(0).equals(oneVertex)) {
			return listOfVertexs.get(1);
		} else {
			return listOfVertexs.get(0);
		}
	}
	
	@Override
	public boolean equals(Object obj) {
		Edge another = (Edge) obj;
		return this.getId().equals(another.getId());
	}
}