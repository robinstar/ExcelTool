package com.robin.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Properties;

public class PropUtils {

	public static Properties read(String propertiesFilePath) throws IOException {
		checkPath(propertiesFilePath);

		InputStream propInputStream = new FileInputStream(propertiesFilePath);
		Properties prop = new Properties();
		prop.load(propInputStream);
		propInputStream.close();

		return prop;
	}

	public static void write(String propertiesFilePath, String key, String value) throws IOException {
		checkPath(propertiesFilePath);
		if (StringUtils.isEmpty(key)) {
			throw new IllegalArgumentException("empty property key =" + key);
		}

		Properties prop = new Properties();
		File file = new File(propertiesFilePath);
		if (!file.exists()) {
			file.createNewFile();
		}

		InputStream fis = new FileInputStream(file);
		prop.load(fis);
		fis.close();

		OutputStream fos = new FileOutputStream(propertiesFilePath);
		prop.setProperty(key, value);
		prop.store(fos, "Update '" + key + "' value");
		fos.close();
	}

	public static void write(String propertiesFilePath, Properties prop) throws IOException {
		checkPath(propertiesFilePath);
		if (prop == null) {
			throw new IllegalArgumentException("Properties is null");
		}

		OutputStream fos = new FileOutputStream(propertiesFilePath);
		prop.store(fos, null);
		fos.close();
	}

	static void checkPath(String path) {
		if (StringUtils.isEmpty(path)) {
			throw new IllegalArgumentException("empty file path =" + path);
		}
	}
}
