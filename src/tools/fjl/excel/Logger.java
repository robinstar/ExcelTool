package tools.fjl.excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Logger {

	private static final String USER_HOME = System.getProperty("user.home");
	private static final String LOG_NAME = "excel.log";

	public static void timestamp() {
		PrintWriter writer = getLogger();
		try {
			if (writer != null) {
				writer.println(currentTimestamp());
			}
		} finally {
			if (writer != null) {
				writer.close();
			}
		}
	}

	public static void log(String msg) {
		PrintWriter writer = getLogger();
		try {
			if (writer != null) {
				if (msg != null && !"".equals(msg)) {
					writer.print(currentTimestamp());
					writer.print(" ");
				}
				writer.println(msg);
			}
		} finally {
			if (writer != null) {
				writer.close();
			}
		}
	}

	public static void log(Exception e) {
		PrintWriter writer = getLogger();
		try {
			if (writer != null) {
				String msg = e.getMessage();
				if (msg == null || "".equals(msg)) {
					writer.println(msg);
				}

				e.printStackTrace(writer);
			}
		} finally {
			if (writer != null) {
				writer.close();
			}
		}
	}

	private static String currentTimestamp() {
		SimpleDateFormat format = new SimpleDateFormat("yyyy.mm.dd HH:mm:ss");
		return format.format(new Date(System.currentTimeMillis()));
	}

	private static PrintWriter getLogger() {
		File file = new File(USER_HOME, LOG_NAME);
		if (!file.exists()) {
			try {
				file.createNewFile();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		if (!file.exists()) {
			return null;
		}

		OutputStream os = null;
		try {
			os = new FileOutputStream(file, true);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}

		if (os == null) {
			return null;
		}

		return new PrintWriter(os);
	}
}
