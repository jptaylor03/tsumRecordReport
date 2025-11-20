package thr.model;

import java.util.Map;

public class RecordData {
	// Common
	private String id;
	// Friend
	private long lastReceiveTime;
	Map<String, Integer> receiveCounts;
	// Totals
	private int receivedCount;
	private int sentCount;
	
	// Common
	public String getId() {
		return this.id;
	};
	public void setId(String id) {
		this.id = id;
	}
	
	// Friend
	public long getLastReceiveTime() {
		return lastReceiveTime;
	}
	public void setLastReceiveTime(long lastReceiveTime) {
		this.lastReceiveTime = lastReceiveTime;
	}
	public Map<String, Integer> getReceiveCounts() {
		return receiveCounts;
	}
	public void setReceiveCounts(Map<String, Integer> receiveCounts) {
		this.receiveCounts = receiveCounts;
	}
	
	// Totals
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
}
