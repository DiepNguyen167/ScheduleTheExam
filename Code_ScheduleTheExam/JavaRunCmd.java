package LichThi;

import java.io.IOException;

public class JavaRunCmd {
	public static void exec(String cmd) {
		Runtime rt = Runtime.getRuntime();
		try {
			Process pr = rt.exec(cmd);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
