package thr.utils;

import java.io.IOException;
import java.util.Properties;

public class MavenProperties {
	private static Properties properties = null;
	
	private static Properties obtainProperties() {
		if (properties == null) {
			properties = new Properties();
			try {
				properties.load(MavenProperties.class.getResourceAsStream("/pom.properties"));
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return properties;
	}
	
    public static String obtainGroupId() {
    	return (String)obtainProperties().get("groupId");
    }
    public static String obtainArtifactId() {
    	return (String)obtainProperties().get("artifactId");
    }
    public static String obtainVersion() {
    	return (String)obtainProperties().get("version");
    }
    public static String obtainName() {
    	return (String)obtainProperties().get("name");
    }
}
