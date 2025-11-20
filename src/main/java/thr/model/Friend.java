package thr.model;

import java.util.Date;
import java.util.Map;
import java.util.Objects;

public class Friend {
	private String id;
	private long lastReceiveTime;
	Map<String, Integer> receiveCounts;
	private String name;
	
	public Friend() {}
	public Friend(String id, long lastReceivedTime, Map<String, Integer> receiveCounts) {
		this.id = id;
		this.lastReceiveTime = lastReceivedTime;
		this.receiveCounts = receiveCounts;
	}
	
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	public long getLastReceiveTime() {
		return lastReceiveTime;
	}
	public Date getLastReceivedDate() {
		return new Date(this.lastReceiveTime);
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
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}

	@Override
	public boolean equals(Object other) {
		if (this == other) return true;
		if (other == null || this.getClass() != other.getClass()) return false;
		Friend otherCasted = (Friend)other;
		return Objects.equals(this.getId(), otherCasted.getId())
		//	&& Objects.equals(this.getName(), otherCasted.getName())
		//	&& Long.compare(this.getLastReceiveTime(), otherCasted.getLastReceiveTime()) == 0
		//	&& Objects.equals(this.getReceiveCounts(), otherCasted.getReceiveCounts())
        ;
	}

	@Override
	public int hashCode() {
		int result = Objects.hashCode(this.getId());
	//	result = 31 * result + Objects.hashCode(this.getName());
	//	result = 31 * result + Long.hashCode(this.getLastReceiveTime());
	//	result = 31 * result + Objects.hashCode(this.getReceiveCounts());
		return result;
	}

	@Override
	public String toString() {
		return "{ id=" + this.getId()
			+ ", name=" + this.getName()
			+ ", lastReceiveTime=" + this.getLastReceiveTime() + " (" + this.getLastReceivedDate() + ")"
			+ ", receiveCounts=" + this.getReceiveCounts()
			+ " }";
	}
}
