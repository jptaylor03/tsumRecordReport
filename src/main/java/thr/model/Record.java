package thr.model;

import java.util.LinkedHashMap;
import java.util.Map;

import com.fasterxml.jackson.annotation.JsonAnySetter;

public class Record {
	
	private Map<String, RecordData> data = new LinkedHashMap<String, RecordData>();
	
	public Map<String, RecordData> getData() {
		return this.data;
	}
	@JsonAnySetter
	public void setData(String key, RecordData value) {
		data.put(key, value);
	}
	
}
