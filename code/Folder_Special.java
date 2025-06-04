package email.code;

import java.io.File;

import com.aspose.email.EmailClient;
import com.aspose.email.IConnection;
import com.aspose.email.ImapClient;
import com.aspose.email.SecurityOptions;

public class Folder_Special {
	static ImapClient clientforimap_output;

	public static void main(String[] args) {
		System.out.println(System.getProperty("user.home") + File.separator + "Desktop");
		try {
			clientforimap_output = new ImapClient("imap.gmail.com", 993, "devloperarysonrahul@gmail.com",
					"ntzkgzvfafnrfhgp");
			clientforimap_output.setSecurityOptions(SecurityOptions.Auto);
			EmailClient.setSocketsLayerVersion2(true);
			EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
			clientforimap_output.setTimeout(5 * 60 * 1000);
			IConnection iconnforimap_output = clientforimap_output.createConnection();
			String x1;
			System.out.println("before " + clientforimap_output.getDelimiter());
			System.out.println(" after " + clientforimap_output.getDelimiter());
			String path = "752575sd2752";
			x1 = path;
			String[] swd = { "Tue Nov 01 17 54 48 IST 2022", "deepaktestingid88@gmail.com", "Top of Personal Folders",
					"deepaktestingid88@gmailcom", "bhardwajdeepak@aolcom-29-10-22 16-18-39",
					"Fri Jun 10 18 18 10 IST 2022", "deepaktestingid8802@gmailcom", "INBOX", "A", "B", "C", "D", "E",
					"F", "G" };
			for (int i = 0; i < swd.length; i++) {
				System.out.println(" I " + i);
				try {
					String foldername = swd[i];
					if (swd[i].contains(".")) {
						foldername = swd[i].replace(".", " ");
					}
					// x1=INBOX + . + path
					x1 = x1 + clientforimap_output.getDelimiter() + foldername;
					System.out.println("x1 +  " + x1);
					if (clientforimap_output.existFolder(x1)) {
						clientforimap_output.selectFolder(x1);
					} else {
						clientforimap_output.createFolder(iconnforimap_output, x1);
						clientforimap_output.selectFolder(x1);
					}
				} catch (Exception e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
			System.out.println("creation done");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
