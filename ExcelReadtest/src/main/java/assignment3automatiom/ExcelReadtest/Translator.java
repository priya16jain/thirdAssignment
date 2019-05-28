package assignment3automatiom.ExcelReadtest;
import java.io.BufferedReader;
import java.io.OutputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;

import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Translator {
	private static final String CLIENT_ID = "FREE_TRIAL_ACCOUNT";
	private static final String CLIENT_SECRET = "PUBLIC_SECRET";
	private static final String ENDPOINT = "http://api.whatsmate.net/v1/translation/translate";
	
   /**
	 * Entry Point
	 */
	public static void main(String[] args) throws Exception {
		// TODO: Specify your translation requirements here:
		String fromLang = "en";
		String toLang = "hi";
	    XSSFSheet data = ReadTranslatorFile.getSheetValues();
		Translator.translate(fromLang, toLang, data);
	}

   public static void translate(String fromLang, String toLang, XSSFSheet data) throws Exception {
		// TODO: Should have used a 3rd party library to make a JSON string from an object
		String jsonPayload = new StringBuilder()
				.append("{")
				.append("\"fromLang\":\"")
				.append(fromLang)
				.append("\",")
				.append("\"toLang\":\"")
				.append(toLang)
				.append("\",")
				.append("\"text\":\"")
				.append(data)
				.append("\"")
				.append("}")
				.toString();

		URL url = new URL(ENDPOINT);
		HttpURLConnection conn = (HttpURLConnection) url.openConnection();
		conn.setDoOutput(true);
		conn.setRequestMethod("POST");
		conn.setRequestProperty("X-WM-CLIENT-ID", CLIENT_ID);
		conn.setRequestProperty("X-WM-CLIENT-SECRET", CLIENT_SECRET);
		conn.setRequestProperty("Content-Type", "application/json");

		OutputStream os = conn.getOutputStream();
		os.write(jsonPayload.getBytes());
		os.flush();
		os.close();

		int statusCode = conn.getResponseCode();
		System.out.println("Status Code: " + statusCode);
		BufferedReader br = new BufferedReader(new InputStreamReader(
				(statusCode == 200) ? conn.getInputStream() : conn.getErrorStream()
				));
		String output;
		while ((output = br.readLine()) != null) {
			System.out.println(output);
		}
		conn.disconnect();
	}

}



