package email.code;

import java.io.File;
import com.aspose.email.HtmlFormatOptions;
import com.aspose.email.HtmlSaveOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MailMessageSaveType;
import com.aspose.email.MapiAttachment;
import com.aspose.email.MapiAttachmentCollection;
import com.aspose.email.MapiMessage;
import com.aspose.email.SaveOptions;

public class mapifile implements Runnable {

	private Main_Frame mf;
	private main_multiplefile mm;
	private ConvertPSTOST_file cf;
	private String filetype = "";
	private String filepath = "";
	private String destination_path = "";
	private String path = "";
	long count_destination;
	static boolean datevalidflag = false;
	private MapiMessage message = null;
	MailMessage mess = null;
	long k;
	File filecheck = null;

	public mapifile(main_multiplefile mm, Main_Frame mf, String filetype, String destination_path, String path,
			long count_destination, MapiMessage message, String filepath, ConvertPSTOST_file cf, MailMessage mess) {
		this.mf = mf;
		this.mess = mess;
		this.mm = mm;
		this.cf = cf;
		this.filetype = filetype;
		this.destination_path = destination_path;
		this.path = path;
		this.count_destination = count_destination;
		this.message = message;
		mf.stop = false;

	}

	public void run() {
		mapidupfile(message);
		ConvertPSTOST_file.count_destination = count_destination;
	}

	void mapidupfile(MapiMessage message) {

		if (main_multiplefile.datefilter.isSelected()) {
			datevalidflag = mm.checkdate(message, mess);
		}

		if (mm.chckbxRemoveDuplicacy.isSelected()) {
			String input = mm.duplicacymapi(message);
			if (!cf.listduplicacy.contains(input)) {
				cf.listduplicacy.add(input);

				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						Mapimessage_file(message);
						count_destination++;
						ConvertPSTOST_file.foldermessagecount++;
					}
				} else {
					Mapimessage_file(message);
					count_destination++;
					ConvertPSTOST_file.foldermessagecount++;
				}
			}
		} else {
			if (main_multiplefile.datefilter.isSelected()) {
				if (ConvertPSTOST_file.datevalidflag) {
					Mapimessage_file(message);
					count_destination++;
					ConvertPSTOST_file.foldermessagecount++;
				}
			} else {
				Mapimessage_file(message);
				count_destination++;
				ConvertPSTOST_file.foldermessagecount++;
			}
		}
	}

	@SuppressWarnings("deprecation")
	public void Mapimessage_file(MapiMessage message) {
		String subname = "";
		if (Main_Frame.fileoption.equals("OLM files (.olm)")) {
			subname = mf.namingconventionmapi(message, mf.c1);
		} else if (Main_Frame.fileoption.equals("EML File (.eml)") || Main_Frame.fileoption.equals("EMLX File (.emlx)")
				|| Main_Frame.fileoption.equals("Message File (.msg)") || Main_Frame.fileoption.equals("Maildir")) {
			subname = mm.namingconventionmapi(message, new File(filepath));
		} else {
			subname = mm.namingconventionmapi(message);
		}
//		String finalpath = destination_path + File.separator + path + File.separator
//				+ Main_Frame.getRidOfIllegalFileNameCharacters(subname);
		String finalpath = destination_path + File.separator + path + File.separator
				+ Main_Frame.getRidOfIllegalFileNameCharacters(mf.namingconventionmapi(message).trim());
		try {
			File fu = null;
			if (!mm.chckbxMigrateOrBackup.isSelected() && message.getAttachments().size() > 0
					&& mm.chckbxSavePdfAttachment.isSelected()) {
				fu = new File(destination_path + File.separator + path + File.separator + "Attachment");
				fu.mkdirs();
			}
			if (filetype.equalsIgnoreCase("EML")) {
				try {
					filecheck = new File(finalpath + ".eml");
					if (filecheck.exists()) {
						filecheck = new File(finalpath + "_" + count_destination + ".eml");
					}
					save_Attachment_File(message, fu, subname);
					message.save(filecheck.getAbsolutePath(), SaveOptions.getDefaultEml());

				} catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

			} else if (filetype.equalsIgnoreCase("HTML")) {

				try {
					filecheck = new File(finalpath + ".html");
					if (filecheck.exists()) {
						filecheck = new File(finalpath + "_" + count_destination + ".html");
					}
					HtmlSaveOptions options = SaveOptions.getDefaultHtml();
					options.setEmbedResources(false);
					options.setHtmlFormatOptions(
							HtmlFormatOptions.WriteHeader | HtmlFormatOptions.WriteCompleteEmailAddress);
					save_Attachment_File(message, fu, subname);

					message.save(filecheck.getAbsolutePath(), options);

				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

			} else if (filetype.equalsIgnoreCase("MSG")) {

				try {
					filecheck = new File(finalpath + ".msg");
					if (filecheck.exists()) {
						filecheck = new File(finalpath + "_" + count_destination + ".msg");
					}
					save_Attachment_File(message, fu, subname);
					message.save(filecheck.getAbsolutePath(), SaveOptions.getDefaultMsg());

				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

				catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

			} else if (filetype.equalsIgnoreCase("EMLX")) {
				MailMessageSaveType messagesavetype = MailMessageSaveType.getEmlxFormat();
				try {
					filecheck = new File(finalpath + ".emlx");
					if (filecheck.exists()) {
						filecheck = new File(finalpath + "_" + count_destination + ".emlx");
					}
					save_Attachment_File(message, fu, subname);
					message.save(filecheck.getAbsolutePath(), SaveOptions.createSaveOptions(messagesavetype));

				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

				catch (Exception e) {
					e.printStackTrace();
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}
			} else if (filetype.equalsIgnoreCase("MHTML")) {

				try {
					filecheck = new File(finalpath + ".mhtml");
					if (filecheck.exists()) {
						filecheck = new File(finalpath + "_" + count_destination + ".mhtml");
					}
					save_Attachment_File(message, fu, subname);
					message.save(filecheck.getAbsolutePath(), SaveOptions.getDefaultMhtml());

				} catch (Error e) {
					mf.logger.warning("ERROR" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				} catch (Exception e) {
					mf.logger.warning("Exception" + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}

			}
			fu = null;

			k = Main_Frame.count_destination++;

		} catch (Exception e) {

			e.printStackTrace();
		}
	}

	void save_Attachment_File(MapiMessage message, File fu, String subname) {
		if (mm.chckbxSavePdfAttachment.isSelected()) {
			MapiAttachmentCollection attachments = message.getAttachments();
			File f3 = null;
			if (attachments.size() > 0) {
				f3 = new File(fu.getAbsolutePath() + File.separator + subname.trim());
				if (f3.exists()) {
					f3 = new File(fu.getAbsolutePath() + File.separator + subname.trim() + "_" + count_destination);
				}
				f3.mkdirs();
				for (MapiAttachment attachment : message.getAttachments()) {
					try {
						attachment.save(f3.getAbsolutePath() + File.separator + attachment.getDisplayName().trim());
					} catch (Exception e) {
						attachment.save(f3.getAbsolutePath() + File.separator + attachment.getLongFileName().trim());
					}
				}
				message.getAttachments().clear();
			}
		}

	}

}
