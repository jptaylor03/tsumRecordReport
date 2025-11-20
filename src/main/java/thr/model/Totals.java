package thr.model;

public class Totals {
	private String id;
	private int receivedCount;
	private int sentCount;
	
	public Totals() {}
	public Totals(String id, int receivedCount, int sentCount) {
		this.id = id;
		this.receivedCount = receivedCount;
		this.sentCount = sentCount;
	}
	
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	public int getReceivedCount() {
		return receivedCount;
	}
	public void setReceivedCount(int receivedCount) {
		this.receivedCount = receivedCount;
	}
	public int getSentCount() {
		return sentCount;
	}
	public void setSentCount(int sentCount) {
		this.sentCount = sentCount;
	}

	public String toString() {
		return "{ id=" + this.getId()
			+ ", receivedCount=" + this.getReceivedCount()
			+ ", sentCount=" + this.getSentCount()
			+ " }";
	}
}
