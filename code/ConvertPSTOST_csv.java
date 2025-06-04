package email.code;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.aspose.email.FolderInfo;
import com.aspose.email.FolderInfoCollection;
import com.aspose.email.FollowUpManager;
import com.aspose.email.FollowUpOptions;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiAttachment;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiContactElectronicAddress;
import com.aspose.email.MapiContactEventPropertySet;
import com.aspose.email.MapiContactNamePropertySet;
import com.aspose.email.MapiContactOtherPropertySet;
import com.aspose.email.MapiContactPersonalInfoPropertySet;
import com.aspose.email.MapiContactPhysicalAddress;
import com.aspose.email.MapiContactPhysicalAddressPropertySet;
import com.aspose.email.MapiContactProfessionalPropertySet;
import com.aspose.email.MapiContactTelephonePropertySet;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiElectronicAddress;
import com.aspose.email.MapiMessage;
import com.aspose.email.MapiTask;
import com.aspose.email.MapiTaskUsers;
import com.aspose.email.MessageInfo;
import com.aspose.email.MessageInfoBase;
import com.aspose.email.MessageInfoCollection;
import com.aspose.email.PersonalStorage;
import com.opencsv.CSVWriter;

public class ConvertPSTOST_csv extends MessageInfoBase implements Runnable {
	String fOLdER = "";
	int number = 1;
	private List<String> filelist = new ArrayList<String>();
	File f3 = null;
	List<String> list_To = new ArrayList<String>();
	List<String> listduplictask = new ArrayList<String>();
	List<String> listdupliccontact = new ArrayList<String>();
	CSVWriter writer;
	long foldermessagecount;
	List<String> listdupliccal = new ArrayList<String>();
	List<String> listduplicacy = new ArrayList<String>();
	static Date fromdate;
	static Date todate;
	ArrayList<Date> fromList = new ArrayList<Date>();
	ArrayList<Date> toList = new ArrayList<Date>();
	String from;
	String to;
	String first = null, middle = null, last = null;
	private Main_Frame mf;
	private String path = "";
	private main_multiplefile mm;
	private String filetype = "";
	private String filepath = "";
	private String destination_path = "";
	static long count_destination;
	static PersonalStorage pst;
	String Folder;
	List<String> pstfolderlist;
	static boolean datevalidflag = false;
	String temppathm = "";
	MapiMessage msg = null;
	String mime_tag = "";

	public ConvertPSTOST_csv(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm) {
		this.mf = mf;
		this.mm = mm;
		this.filetype = filetype;
		this.filepath = filepath;
		this.destination_path = destination_path;
		ConvertPSTOST_csv.count_destination = count_destination;
		this.pstfolderlist = pstfolderlist;
		this.fromList = fromList;
		this.toList = toList;
		this.temppathm = temppathm;
	}

	public void run() {
		try {
			convertPSTOST_csv(mf, filetype, destination_path, count_destination, filepath, mm, pstfolderlist, fromList,
					toList, temppathm);
			main_multiplefile.count_destination = ConvertPSTOST_csv.count_destination;
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	private void convertPSTOST_csv(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList, String temppathm) throws IOException {
		pst = PersonalStorage.fromFile(filepath);
		ConvertPSTOST_csv.count_destination = 0;
		FolderInfo folderInfo2 = pst.getRootFolder();
		Folder = folderInfo2.getDisplayName();
		Folder = Folder.replace(",", "").replace(".", "");
		Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
		Folder = Folder.replaceAll("[\\[\\]]", "");
		Folder = Folder.trim();
		fOLdER = Folder;
		if (Folder.equalsIgnoreCase("")) {
			Folder = "Root Folder";
		}
		String path1 = "";
		path = path + File.separator + Folder;
		path1 = Folder;

		MessageInfoCollection messageInfoCollection1 = folderInfo2.getContents();

		new File(destination_path + File.separator + path).mkdirs();

		CSVWriter arr3[] = new CSVWriter[4];
		boolean arrA[] = new boolean[4];
		for (int i = 0; i < arrA.length; i++) {
			arrA[i] = false;
		}
		int messagesize1;
		listduplictask.clear();
		listdupliccal.clear();
		listduplicacy.clear();
		listdupliccontact.clear();
		if (main_multiplefile.demo) {
			if (messageInfoCollection1.size() <= All_Data.demo_count) {
				messagesize1 = messageInfoCollection1.size();
			} else {
				messagesize1 = All_Data.demo_count;
			}

		} else {
			messagesize1 = messageInfoCollection1.size();
		}

		for (int i = 0; i < messagesize1; i++) {

			try {

				if (mm.stop) {
					break;
				}
				if ((i % 100) == 0) {
					System.gc();
				}
				MessageInfo messageInfo = (MessageInfo) messageInfoCollection1.get_Item(i);

				MapiMessage message1 = pst.extractMessage(messageInfo);
				MapiConversionOptions d = MapiConversionOptions.getUnicodeFormat();
				MailConversionOptions de = new MailConversionOptions();
				MailMessage mess1 = message1.toMailMessage(de);

				if (mm.chckbxMigrateOrBackup.isSelected()) {
					mess1.getAttachments().clear();
				}
				MapiMessage message = MapiMessage.fromMailMessage(mess1, d);
				if (main_multiplefile.datefilter.isSelected()) {
					datevalidflag = mm.checkdate(message1, mess1);
				}
				if (message1.getMessageClass().equals("IPM.Contact")) {
					try {

						if (!arrA[0]) {

							String[] header = { "Account", "Address Selected", "Address Selector", "Anniversary",
									"Assistant's Name", "Assistant's Phone", "Attachment", "Billing Information",
									"Birthday", "Business Address", "Business Address City",
									"Business Address Country/Region", "Business Address PO Box",
									"Business Address Postal Code", "Business Address State", "Business Address Street",
									"Business Fax", "Business Home Page", "Business Phone", "Business Phone 2",
									"Callback", "Car Phone", "Categories", "Children", "City", "Company",
									"Company Main Phone", "Computer Network Name", "Contacts", "Country/Region",
									"Created", "Customer ID", "Department", "Email", "Email 2", "Email 3",
									"Email Address Type", "Email Display As", "Email Selected", "Email Selector",
									"Email2 Address Type", "Email2 Display As", "Email3 Address Type",
									"Email3 Display As", "File As", "First Name", "Flag Completed Date", "Flag Status",
									"Follow Up Flag", "FTP Site", "Full Name", "Gender", "Government ID Number",
									"Hobbies", "Home Address", "Home Address City", "Home Address Country/Region",
									"Home Address PO Box", "Home Address Postal Code", "Home Address State",
									"Home Address Street", "Home Fax", "Home Phone", "Home Phone 2", "Icon",
									"IM Address", "In Folder", "Initials", "Internet Free/Busy Address", "ISDN",
									"Job Title", "Journal", "Language", "Last Name", "Location", "Mailing Address",
									"Mailing Address Indicator", "Manager's Name", "Message Class", "Middle Name",
									"Mileage", "Mobile Phone", "Modified", "Nickname", "Office Location",
									"Organizational ID Number", "Other Address", "Other Address City",
									"Other Address Country/Region", "Other Address PO Box", "Other Address Postal Code",
									"Other Address State", "Other Address Street", "Other Fax", "Other Phone",
									"Outlook Data File", "Outlook Internal Version", "Outlook Version", "Pager",
									"Personal Home Page", "Phone 1 Selected", "Phone 1 Selector", "Phone 2 Selected",
									"Phone 2 Selector", "Phone 3 Selected", "Phone 3 Selector", "Phone 4 Selected",
									"Phone 4 Selector", "Phone 5 Selected", "Phone 5 Selector", "Phone 6 Selected",
									"Phone 6 Selector", "Phone 7 Selected", "Phone 7 Selector", "Phone 8 Selected",
									"Phone 8 Selector", "PO Box", "Primary Phone", "Private", "Profession",
									"Radio Phone", "Read", "Referred By", "Reminder", "Reminder Time", "Reminder Topic",
									"Sensitivity", "Size", "Size on Server", "Spouse/Partner", "State",
									"Street Address", "Subject", "Suffix", "Telex", "Title", "TTY/TDD Phone",
									"User Field 1", "User Field 2", "User Field 3", "User Field 4", "Web Page",
									"ZIP/Postal Code" };

							File file2 = new File(destination_path + File.separator + path + File.separator + Folder
									+ "_Contact" + ".csv");

							FileWriter outputfile2 = new FileWriter(file2);

							writer = new CSVWriter(outputfile2);
							writer.writeNext(header);
							arr3[0] = writer;
							arrA[0] = true;
						}

						MapiContact con = (MapiContact) message1.toMapiMessageItem();
						if (mm.chckbxRemoveDuplicacy.isSelected()) {

							String input = mm.duplicacymapiContact(con);
							input = input.replaceAll("\\s", "");
							input = input.trim();
							if (!listdupliccontact.contains(input)) {
								listdupliccontact.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										Mapimess_CSV(message1, arr3[0]);
										ConvertPSTOST_csv.count_destination++;
									}
								} else {
									Mapimess_CSV(message1, arr3[0]);
									ConvertPSTOST_csv.count_destination++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									Mapimess_CSV(message1, arr3[0]);
									ConvertPSTOST_csv.count_destination++;
								}
							} else {
								Mapimess_CSV(message1, arr3[0]);
								ConvertPSTOST_csv.count_destination++;
							}
						}
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " " + i
								+ mf.namingconventionmapi(message1) + System.lineSeparator());

						e.printStackTrace();
						continue;
					}

				} else if (message1.getMessageClass().equals("IPM.Appointment")
						|| message1.getMessageClass().contains("IPM.Schedule.Meeting")) {

					try {
						if (!arrA[2]) {
							String[] header = { "Subject", "Body", "From ", "To", "Cc", "Bcc", "Start Time", "End Time",
									"All day event", "Reminder on/off", "Meeting Workspace URL", "Reminder Time",
									"Categories", "Location", "Attachment Path" };

							File file2 = new File(destination_path + File.separator + path + File.separator + Folder
									+ "_Calendar" + ".csv");

							FileWriter outputfile2 = new FileWriter(file2);

							writer = new CSVWriter(outputfile2);
							writer.writeNext(header);
							arr3[2] = writer;
							arrA[2] = true;
						}

						MapiCalendar cal = (MapiCalendar) message1.toMapiMessageItem();
						if (mm.chckbxRemoveDuplicacy.isSelected()) {

							String input = mm.duplicacymapiCal(cal);
							input = input.replaceAll("\\s", "");
							input = input.trim();

							if (!listdupliccal.contains(input)) {
								listdupliccal.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										Mapimess_CSV(message1, arr3[2]);
										ConvertPSTOST_csv.count_destination++;
									}
								} else {
									Mapimess_CSV(message1, arr3[2]);
									ConvertPSTOST_csv.count_destination++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									Mapimess_CSV(message1, arr3[2]);
									ConvertPSTOST_csv.count_destination++;
								}
							} else {
								Mapimess_CSV(message1, arr3[2]);
								ConvertPSTOST_csv.count_destination++;
							}
						}

					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message1));
					} catch (Exception e) {
						mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " " + i
								+ mf.namingconventionmapi(message1) + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else if (message1.getMessageClass().equals("IPM.Task")) {

					String s = filetype;
					try {
						if (!arrA[1]) {
							String[] header = { "Subject", "Start Date", "Due Date", "Percentage Complete",
									"Estimate Effort", "Actual Effort", "Owner", "Last User", "Last Delegate",
									"Attende size", "Original Display Name", "Display Name", "Email Address",
									"Fax Number", "Address Type", "Comapanies", "Categories", "Mileage", "Billing",
									"Sensitivity", "Status", "History" };
							File file2 = new File(destination_path + File.separator + path + File.separator + Folder
									+ "_Task" + ".csv");

							FileWriter outputfile2 = new FileWriter(file2);

							writer = new CSVWriter(outputfile2);
							writer.writeNext(header);
							arr3[1] = writer;
							arrA[1] = true;
						}

						filetype = "MSG";

						MapiTask task = null;
						if (message1.getMessageClass().equals("IPM.Task")) {
							task = (MapiTask) message1.toMapiMessageItem();
						}

						if (mm.chckbxRemoveDuplicacy.isSelected()) {
							String input = "";
							if (message1.getMessageClass().equals("IPM.Task")) {
								input = mm.duplicacymapiTask(task);
							}
							input = input.replaceAll("\\s", "");
							input = input.trim();

							if (!listduplictask.contains(input)) {
								listduplictask.add(input);

								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										Mapimess_CSV(message1, arr3[1]);
										ConvertPSTOST_csv.count_destination++;
									}
								} else {
									Mapimess_CSV(message1, arr3[1]);
									ConvertPSTOST_csv.count_destination++;
								}
							}
						} else {

							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {
									Mapimess_CSV(message1, arr3[1]);
									ConvertPSTOST_csv.count_destination++;
								}
							} else {
								Mapimess_CSV(message1, arr3[1]);
								ConvertPSTOST_csv.count_destination++;
							}
						}

					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message1));
					} catch (Exception e) {
						mf.logger.warning("Exception : " + e.getMessage() + "Tasks" + " " + i
								+ mf.namingconventionmapi(message1) + System.lineSeparator());
						e.printStackTrace();
						continue;
					} finally {
						filetype = s;
					}

				} else {
					try {
						if (!arrA[3]) {

							String[] header = { "Date", "Subject", "Body", "From", "To", "CC:", "BCC:",
									"Attachment path" };

							File file = new File(destination_path + File.separator + path + File.separator + Folder
									+ "_MailItems" + ".csv");

							FileWriter outputfile = new FileWriter(file);

							writer = new CSVWriter(outputfile);
							writer.writeNext(header);
							arr3[3] = writer;
							arrA[3] = true;
						}
						mapicsv(message1, arr3[3]);
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message1));
					} catch (Exception e) {
						mf.logger.warning("Exception : " + e.getMessage() + "Message " + " " + i
								+ mf.namingconventionmapi(message1) + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				}

				mm.lbl_progressreport.setText("Total Message Saved Count  " + ConvertPSTOST_csv.count_destination + "  "
						+ Folder + "   Extracting messsage " + message1.getSubject());

			} catch (Exception e) {
				continue;
			}

		}
		for (int i = 0; i < arr3.length; i++) {
			if (arr3[i] != null) {
				arr3[i].close();
			}
		}

		FolderInfoCollection folderInf = pst.getRootFolder().getSubFolders();

		for (int j = 0; j < folderInf.size(); j++) {
			try {
				if (mm.stop) {
					break;
				}

				FolderInfo folderInfo = folderInf.get_Item(j);
				String Folder = folderInfo.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();
				fOLdER = Folder;
				path = path + File.separator + Folder;
				String path3 = path1 + File.separator + Folder;

				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (mm.stop) {
						break;
					}
					if (pstfolderlist.get(l).trim().equalsIgnoreCase(path3)) {

						mm.lbl_progressreport.setText("Getting Folder " + Folder);

						new File(destination_path + File.separator + path).mkdirs();

						CSVWriter arr1[] = new CSVWriter[4];
						boolean arrB[] = new boolean[4];
						for (int i = 0; i < arrB.length; i++) {
							arrB[i] = false;
						}

						MessageInfoCollection messageInfoCollection = folderInfo.getContents();
						listduplicacy.clear();
						listdupliccal.clear();
						listduplictask.clear();
						listdupliccontact.clear();
						int messagesize;
						if (main_multiplefile.demo) {
							if (messageInfoCollection.size() <= All_Data.demo_count) {
								messagesize = messageInfoCollection.size();
							} else {
								messagesize = All_Data.demo_count;
							}

						} else {
							messagesize = messageInfoCollection.size();
						}

						for (int i = 0; i < messagesize; i++) {

							try {

								if (mm.stop) {
									break;
								}
								if ((i % 100) == 0) {
									System.gc();
								}
								MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

								MapiMessage message1 = pst.extractMessage(messageInfo);
								MapiConversionOptions d = MapiConversionOptions.getUnicodeFormat();
								MailConversionOptions de = new MailConversionOptions();
								MailMessage mess = message1.toMailMessage(de);
								if (mm.chckbxMigrateOrBackup.isSelected()) {
									mess.getAttachments().clear();
								}
								MapiMessage message = MapiMessage.fromMailMessage(mess, d);

								if (main_multiplefile.datefilter.isSelected()) {
									datevalidflag = mm.checkdate(message1, mess);
								}
								if (message1.getMessageClass().equals("IPM.Contact")) {
									try {
										if (!arrB[0]) {

											String[] header = { "Account", "Address Selected", "Address Selector",
													"Anniversary", "Assistant's Name", "Assistant's Phone",
													"Attachment", "Billing Information", "Birthday", "Business Address",
													"Business Address City", "Business Address Country/Region",
													"Business Address PO Box", "Business Address Postal Code",
													"Business Address State", "Business Address Street", "Business Fax",
													"Business Home Page", "Business Phone", "Business Phone 2",
													"Callback", "Car Phone", "Categories", "Children", "City",
													"Company", "Company Main Phone", "Computer Network Name",
													"Contacts", "Country/Region", "Created", "Customer ID",
													"Department", "Email", "Email 2", "Email 3", "Email Address Type",
													"Email Display As", "Email Selected", "Email Selector",
													"Email2 Address Type", "Email2 Display As", "Email3 Address Type",
													"Email3 Display As", "File As", "First Name", "Flag Completed Date",
													"Flag Status", "Follow Up Flag", "FTP Site", "Full Name", "Gender",
													"Government ID Number", "Hobbies", "Home Address",
													"Home Address City", "Home Address Country/Region",
													"Home Address PO Box", "Home Address Postal Code",
													"Home Address State", "Home Address Street", "Home Fax",
													"Home Phone", "Home Phone 2", "Icon", "IM Address", "In Folder",
													"Initials", "Internet Free/Busy Address", "ISDN", "Job Title",
													"Journal", "Language", "Last Name", "Location", "Mailing Address",
													"Mailing Address Indicator", "Manager's Name", "Message Class",
													"Middle Name", "Mileage", "Mobile Phone", "Modified", "Nickname",
													"Office Location", "Organizational ID Number", "Other Address",
													"Other Address City", "Other Address Country/Region",
													"Other Address PO Box", "Other Address Postal Code",
													"Other Address State", "Other Address Street", "Other Fax",
													"Other Phone", "Outlook Data File", "Outlook Internal Version",
													"Outlook Version", "Pager", "Personal Home Page",
													"Phone 1 Selected", "Phone 1 Selector", "Phone 2 Selected",
													"Phone 2 Selector", "Phone 3 Selected", "Phone 3 Selector",
													"Phone 4 Selected", "Phone 4 Selector", "Phone 5 Selected",
													"Phone 5 Selector", "Phone 6 Selected", "Phone 6 Selector",
													"Phone 7 Selected", "Phone 7 Selector", "Phone 8 Selected",
													"Phone 8 Selector", "PO Box", "Primary Phone", "Private",
													"Profession", "Radio Phone", "Read", "Referred By", "Reminder",
													"Reminder Time", "Reminder Topic", "Sensitivity", "Size",
													"Size on Server", "Spouse/Partner", "State", "Street Address",
													"Subject", "Suffix", "Telex", "Title", "TTY/TDD Phone",
													"User Field 1", "User Field 2", "User Field 3", "User Field 4",
													"Web Page", "ZIP/Postal Code" };
											File file2 = new File(destination_path + File.separator + path
													+ File.separator + Folder + "_contact" + ".csv");

											FileWriter outputfile2 = new FileWriter(file2);

											writer = new CSVWriter(outputfile2);
											writer.writeNext(header);
											arr1[0] = writer;
											arrB[0] = true;
										}

										MapiContact con = (MapiContact) message1.toMapiMessageItem();
										if (mm.chckbxRemoveDuplicacy.isSelected()) {

											String input = mm.duplicacymapiContact(con);
											input = input.replaceAll("\\s", "");
											input = input.trim();

											if (!listdupliccontact.contains(input)) {
												listdupliccontact.add(input);

												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														Mapimess_CSV(message1, arr1[0]);
														ConvertPSTOST_csv.count_destination++;
													}
												} else {
													Mapimess_CSV(message1, arr1[0]);
													ConvertPSTOST_csv.count_destination++;
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													Mapimess_CSV(message1, arr1[0]);
													ConvertPSTOST_csv.count_destination++;
												}
											} else {
												Mapimess_CSV(message1, arr1[0]);
												ConvertPSTOST_csv.count_destination++;
											}
										}

									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message1));
									} catch (Exception e) {
										mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " " + i
												+ mf.namingconventionmapi(message1) + System.lineSeparator());

										e.printStackTrace();
										continue;
									}

								} else if (message1.getMessageClass().equals("IPM.Appointment")
										|| message1.getMessageClass().contains("IPM.Schedule.Meeting")) {

									try {
										if (!arrB[2]) {
											String[] header = { "Subject", "Body", "From ", "To", "Cc", "Bcc",
													"Start Time", "End Time", "All day event", "Reminder on/off",
													"Meeting Workspace URL", "Reminder Time", "Categories", "Location",
													"Attachment Path" };

											File file2 = new File(destination_path + File.separator + path
													+ File.separator + Folder + "_Calendar" + ".csv");

											FileWriter outputfile2 = new FileWriter(file2);

											writer = new CSVWriter(outputfile2);
											writer.writeNext(header);
											arr1[2] = writer;
											arrB[2] = true;
										}

										MapiCalendar cal = (MapiCalendar) message1.toMapiMessageItem();
										if (mm.chckbxRemoveDuplicacy.isSelected()) {
											String input = mm.duplicacymapiCal(cal);
											input = input.replaceAll("\\s", "");
											input = input.trim();
											if (!listdupliccal.contains(input)) {
												listdupliccal.add(input);
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														Mapimess_CSV(message1, arr1[2]);
														ConvertPSTOST_csv.count_destination++;
													}
												} else {
													Mapimess_CSV(message1, arr1[2]);
													ConvertPSTOST_csv.count_destination++;
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													Mapimess_CSV(message1, arr1[2]);
													ConvertPSTOST_csv.count_destination++;
												}
											} else {
												Mapimess_CSV(message1, arr1[2]);
												ConvertPSTOST_csv.count_destination++;
											}
										}
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message1));
									} catch (Exception e) {
										mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " " + i
												+ mf.namingconventionmapi(message1) + System.lineSeparator());
										e.printStackTrace();
										continue;
									}

								} else if (message1.getMessageClass().equals("IPM.Task")) {

									String s = filetype;
									try {
										if (!arrB[1]) {
											String[] header = { "Subject", "Start Date", "Due Date",
													"Percentage Complete", "Estimate Effort", "Actual Effort", "Owner",
													"Last User", "Last Delegate", "Attende size",
													"Original Display Name", "Display Name", "Email Address",
													"Fax Number", "Address Type", "Comapanies", "Categories", "Mileage",
													"Billing", "Sensitivity", "Status", "History" };
											File file2 = new File(destination_path + File.separator + path
													+ File.separator + Folder + "_Task" + ".csv");

											FileWriter outputfile2 = new FileWriter(file2);

											writer = new CSVWriter(outputfile2);
											writer.writeNext(header);
											arr1[1] = writer;
											arrB[1] = true;
										}

										filetype = "MSG";
										MapiTask task = null;
										if (message1.getMessageClass().equals("IPM.Task")) {
											task = (MapiTask) message1.toMapiMessageItem();
										}
										if (mm.chckbxRemoveDuplicacy.isSelected()) {
											String input = "";
											if (message1.getMessageClass().equals("IPM.Task")) {
												input = mm.duplicacymapiTask(task);
											}
											input = input.replaceAll("\\s", "");
											input = input.trim();

											if (!listduplictask.contains(input)) {
												listduplictask.add(input);

												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {
														Mapimess_CSV(message1, arr1[1]);
														ConvertPSTOST_csv.count_destination++;
													}
												} else {
													Mapimess_CSV(message1, arr1[1]);
													ConvertPSTOST_csv.count_destination++;
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {
													Mapimess_CSV(message1, arr1[1]);
													ConvertPSTOST_csv.count_destination++;
												}
											} else {
												Mapimess_CSV(message1, arr1[1]);
												ConvertPSTOST_csv.count_destination++;
											}
										}
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message1));
									} catch (Exception e) {
										mf.logger.warning("Exception : " + e.getMessage() + "Tasks" + " " + i
												+ mf.namingconventionmapi(message1) + System.lineSeparator());
										e.printStackTrace();
										continue;
									} finally {
										filetype = s;
									}

								} else {
									try {
										if (!arrB[3]) {

											String[] header = { "Date", "Subject", "Body", "From", "To", "CC:", "BCC:",
													"Attachment path" };

											File file = new File(destination_path + File.separator + path
													+ File.separator + Folder + "_MailItems" + ".csv");

											FileWriter outputfile = new FileWriter(file);

											writer = new CSVWriter(outputfile);
											writer.writeNext(header);
											arr1[3] = writer;
											arrB[3] = true;
										}
										mapicsv(message1, arr1[3]);
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message1));
									} catch (Exception e) {
										mf.logger.warning("Exception : " + e.getMessage() + "Message " + " " + i
												+ mf.namingconventionmapi(message1) + System.lineSeparator());
										e.printStackTrace();
										continue;
									}

								}
								mm.lbl_progressreport
										.setText("Total Message Saved Count  " + ConvertPSTOST_csv.count_destination
												+ "  " + Folder + "   Extracting messsage " + message1.getSubject());
							} catch (Exception e) {
								continue;
							}

						}
						for (int i = 0; i < arr1.length; i++) {
							if (arr1[i] != null) {
								arr1[i].close();
							}
						}
					}
				}
				if (folderInfo.hasSubFolders()) {
					getsubfolderforPSTOST_CSV(folderInfo, path3);
				}
				path = mm.removefolder(path);
			} catch (Exception e) {
				continue;
			}

		}

	}

	void getsubfolderforPSTOST_CSV(FolderInfo f, String path3) {

		FolderInfoCollection folderInf = f.getSubFolders();

		for (int j = 0; j < folderInf.size(); j++) {
			try {
				if (mm.stop) {
					break;
				}
				FolderInfo folderf = folderInf.get_Item(j);

				String Folder = folderf.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = main_multiplefile.getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();
				fOLdER = Folder;
				path = path + File.separator + Folder;
				path3 = path3 + File.separator + Folder;

				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (mm.stop) {
						break;
					}
					if (pstfolderlist.get(l).equalsIgnoreCase(path3)) {
						new File(destination_path + File.separator + path).mkdirs();

						try {
							CSVWriter arr[] = new CSVWriter[4];
							boolean arrB[] = new boolean[4];
							for (int i = 0; i < arrB.length; i++) {
								arrB[i] = false;
							}

							MessageInfoCollection messageInfoCollection = null;
							try {
								messageInfoCollection = folderf.getContents();
							} catch (Exception e1) {

								e1.printStackTrace();
							}

							if (!(messageInfoCollection == null)) {
								listdupliccal.clear();
								listduplicacy.clear();
								listduplictask.clear();
								listdupliccontact.clear();
								int messagesize;
								if (main_multiplefile.demo) {
									if (messageInfoCollection.size() <= All_Data.demo_count) {
										messagesize = messageInfoCollection.size();
									} else {
										messagesize = All_Data.demo_count;
									}
								} else {
									messagesize = messageInfoCollection.size();
								}

								for (int i = 0; i < messagesize; i++) {
									try {
										if (mm.stop) {
											break;
										}
										if ((i % 100) == 0) {
											System.gc();
										}
										MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(i);

										MapiMessage message1 = pst.extractMessage(messageInfo);
										MapiConversionOptions d = MapiConversionOptions.getUnicodeFormat();
										MailConversionOptions de = new MailConversionOptions();
										MailMessage mess = message1.toMailMessage(de);

										MapiMessage message = MapiMessage.fromMailMessage(mess, d);
										if (main_multiplefile.datefilter.isSelected()) {
											datevalidflag = mm.checkdate(message1, mess);
										}
										if (message1.getMessageClass().equals("IPM.Contact")) {
											try {
												if (!arrB[0]) {
													String[] header = { "Account", "Address Selected",
															"Address Selector", "Anniversary", "Assistant's Name",
															"Assistant's Phone", "Attachment", "Billing Information",
															"Birthday", "Business Address", "Business Address City",
															"Business Address Country/Region",
															"Business Address PO Box", "Business Address Postal Code",
															"Business Address State", "Business Address Street",
															"Business Fax", "Business Home Page", "Business Phone",
															"Business Phone 2", "Callback", "Car Phone", "Categories",
															"Children", "City", "Company", "Company Main Phone",
															"Computer Network Name", "Contacts", "Country/Region",
															"Created", "Customer ID", "Department", "Email", "Email 2",
															"Email 3", "Email Address Type", "Email Display As",
															"Email Selected", "Email Selector", "Email2 Address Type",
															"Email2 Display As", "Email3 Address Type",
															"Email3 Display As", "File As", "First Name",
															"Flag Completed Date", "Flag Status", "Follow Up Flag",
															"FTP Site", "Full Name", "Gender", "Government ID Number",
															"Hobbies", "Home Address", "Home Address City",
															"Home Address Country/Region", "Home Address PO Box",
															"Home Address Postal Code", "Home Address State",
															"Home Address Street", "Home Fax", "Home Phone",
															"Home Phone 2", "Icon", "IM Address", "In Folder",
															"Initials", "Internet Free/Busy Address", "ISDN",
															"Job Title", "Journal", "Language", "Last Name", "Location",
															"Mailing Address", "Mailing Address Indicator",
															"Manager's Name", "Message Class", "Middle Name", "Mileage",
															"Mobile Phone", "Modified", "Nickname", "Office Location",
															"Organizational ID Number", "Other Address",
															"Other Address City", "Other Address Country/Region",
															"Other Address PO Box", "Other Address Postal Code",
															"Other Address State", "Other Address Street", "Other Fax",
															"Other Phone", "Outlook Data File",
															"Outlook Internal Version", "Outlook Version", "Pager",
															"Personal Home Page", "Phone 1 Selected",
															"Phone 1 Selector", "Phone 2 Selected", "Phone 2 Selector",
															"Phone 3 Selected", "Phone 3 Selector", "Phone 4 Selected",
															"Phone 4 Selector", "Phone 5 Selected", "Phone 5 Selector",
															"Phone 6 Selected", "Phone 6 Selector", "Phone 7 Selected",
															"Phone 7 Selector", "Phone 8 Selected", "Phone 8 Selector",
															"PO Box", "Primary Phone", "Private", "Profession",
															"Radio Phone", "Read", "Referred By", "Reminder",
															"Reminder Time", "Reminder Topic", "Sensitivity", "Size",
															"Size on Server", "Spouse/Partner", "State",
															"Street Address", "Subject", "Suffix", "Telex", "Title",
															"TTY/TDD Phone", "User Field 1", "User Field 2",
															"User Field 3", "User Field 4", "Web Page",
															"ZIP/Postal Code" };
													File file = new File(destination_path + File.separator + path
															+ File.separator + Folder + "_contact" + ".csv");

													FileWriter outputfile = new FileWriter(file);

													writer = new CSVWriter(outputfile);
													writer.writeNext(header);
													arr[0] = writer;
													arrB[0] = true;
												}
												MapiContact con = (MapiContact) message1.toMapiMessageItem();
												if (mm.chckbxRemoveDuplicacy.isSelected()) {
													String input = mm.duplicacymapiContact(con);
													input = input.replaceAll("\\s", "");
													input = input.trim();

													if (!listdupliccontact.contains(input)) {
														listdupliccontact.add(input);

														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																Mapimess_CSV(message1, arr[0]);
																ConvertPSTOST_csv.count_destination++;
															}
														} else {
															Mapimess_CSV(message1, arr[0]);
															ConvertPSTOST_csv.count_destination++;
														}
													}
												} else {
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															Mapimess_CSV(message1, arr[0]);
															ConvertPSTOST_csv.count_destination++;
														}
													} else {
														Mapimess_CSV(message1, arr[0]);
														ConvertPSTOST_csv.count_destination++;
													}
												}
											} catch (Exception e) {
												mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " "
														+ System.lineSeparator());
												continue;
											}

										} else if (message1.getMessageClass().equals("IPM.Appointment")
												|| message1.getMessageClass().contains("IPM.Schedule.Meeting")) {

											try {
												MapiCalendar cal = (MapiCalendar) message1.toMapiMessageItem();
												if (!arrB[2]) {
													String[] header = { "Subject", "Body", "From ", "To", "Cc", "Bcc",
															"Start Time", "End Time", "All day event",
															"Reminder on/off", "Meeting Workspace URL", "Reminder Time",
															"Categories", "Location", "Attachment Path" };

													File file = new File(destination_path + File.separator + path
															+ File.separator + Folder + "_Calendar" + ".csv");

													FileWriter outputfile = new FileWriter(file);

													writer = new CSVWriter(outputfile);
													writer.writeNext(header);
													arr[2] = writer;
													arrB[2] = true;
												}

												if (mm.chckbxRemoveDuplicacy.isSelected()) {
													String input = mm.duplicacymapiCal(cal);
													input = input.replaceAll("\\s", "");
													input = input.trim();

													if (!listdupliccal.contains(input)) {
														listdupliccal.add(input);
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																Mapimess_CSV(message1, arr[2]);
																ConvertPSTOST_csv.count_destination++;
															}
														} else {
															Mapimess_CSV(message1, arr[2]);
															ConvertPSTOST_csv.count_destination++;
														}
													}
												} else {
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															Mapimess_CSV(message1, arr[2]);
															ConvertPSTOST_csv.count_destination++;
														}
													} else {
														Mapimess_CSV(message1, arr[2]);
														ConvertPSTOST_csv.count_destination++;
													}
												}
											} catch (Exception e) {
												mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " "
														+ System.lineSeparator());
												continue;
											}

										} else if (message1.getMessageClass().equals("IPM.Task")) {
											try {
												if (!arrB[1]) {

													String[] header = { "Subject", "Start Date", "Due Date",
															"Percentage Complete", "Estimate Effort", "Actual Effort",
															"Owner", "Last User", "Last Delegate", "Attende size",
															"Original Display Name", "Display Name", "Email Address",
															"Fax Number", "Address Type", "Comapanies", "Categories",
															"Mileage", "Billing", "Sensitivity", "Status", "History" };
													File file = new File(destination_path + File.separator + path
															+ File.separator + Folder + "_Task" + ".csv");

													FileWriter outputfile = new FileWriter(file);

													writer = new CSVWriter(outputfile);
													writer.writeNext(header);
													arr[1] = writer;
													arrB[1] = true;
												}
												MapiTask task = null;

												if (message1.getMessageClass().equals("IPM.Task")) {
													task = (MapiTask) message1.toMapiMessageItem();
												}

												if (mm.chckbxRemoveDuplicacy.isSelected()) {
													String input = "";
													if (message1.getMessageClass().equals("IPM.Task")) {
														input = mm.duplicacymapiTask(task);
													}
													input = input.replaceAll("\\s", "");
													input = input.trim();

													if (!listduplictask.contains(input)) {
														listduplictask.add(input);
														if (main_multiplefile.datefilter.isSelected()) {
															if (datevalidflag) {
																Mapimess_CSV(message1, arr[1]);
																ConvertPSTOST_csv.count_destination++;
															}
														} else {
															Mapimess_CSV(message1, arr[1]);
															ConvertPSTOST_csv.count_destination++;
														}
													}
												} else {
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {
															Mapimess_CSV(message1, arr[1]);
															ConvertPSTOST_csv.count_destination++;
														}
													} else {
														Mapimess_CSV(message1, arr[1]);
														ConvertPSTOST_csv.count_destination++;
													}
												}
											} catch (Exception e) {
												mf.logger.warning("Exception : " + e.getMessage() + "Task" + " "
														+ System.lineSeparator());
												continue;
											}

										} else {
											try {
												if (!arrB[3]) {

													String[] header = { "Date", "Subject", "Body", "From", "To", "CC:",
															"BCC:", "Attachment path" };

													File file = new File(destination_path + File.separator + path
															+ File.separator + Folder + "_MailItems" + ".csv");

													FileWriter outputfile = new FileWriter(file);

													writer = new CSVWriter(outputfile);
													writer.writeNext(header);
													arr[3] = writer;
													arrB[3] = true;
												}
												mapicsv(message1, arr[3]);
											}

											catch (Exception e) {
												e.printStackTrace();
												continue;
											}
										}
										mm.lbl_progressreport.setText("  Total Message Saved Count  "
												+ ConvertPSTOST_csv.count_destination + "  " + Folder
												+ "   Extracting messsage " + message.getSubject());

									} catch (Exception e) {
										e.printStackTrace();
										continue;
									}
								}
							}
							for (int i = 0; i < arr.length; i++) {
								if (arr[i] != null) {
									arr[i].close();
								}
							}
						} catch (Exception e) {
							continue;
						}
					}
				}
				if (folderf.hasSubFolders()) {
					getsubfolderforPSTOST_CSV(folderf, path3);
				}
				path3 = path3.replace(File.separator + Folder, "");
				path = mm.removefolder(path);
			} catch (Exception e) {
				continue;
			}
		}

	}

	void mapicsv(MapiMessage message, CSVWriter writer) {
		if (mm.chckbxRemoveDuplicacy.isSelected()) {

			String input = mm.duplicacymapi(message);

			if (!listduplicacy.contains(input)) {
				listduplicacy.add(input);

				if (main_multiplefile.datefilter.isSelected()) {
					if (datevalidflag) {
						Mapimess_CSV(message, writer);
						ConvertPSTOST_csv.count_destination++;
					}
				} else {
					Mapimess_CSV(message, writer);
					ConvertPSTOST_csv.count_destination++;
				}
			}
		} else {
			if (main_multiplefile.datefilter.isSelected()) {
				if (datevalidflag) {
					Mapimess_CSV(message, writer);
					ConvertPSTOST_csv.count_destination++;
				}
			} else {
				Mapimess_CSV(message, writer);
				ConvertPSTOST_csv.count_destination++;
			}
		}

	}

	void Mapimess_CSV(MapiMessage message, CSVWriter writer) {
		String subname = main_multiplefile.getRidOfIllegalFileNameCharacters(mf.namingconventionmapi(message));
		if (filelist.contains(subname)) {
			subname = subname + "_" + number;
		} else {
			filelist.add(subname);
		}
		if (message.getMessageClass().equals("IPM.Contact")) {

			String mileage = null;
			String Categories = null;
			String Children = null;
			String nickName = null;
			String Email1addresstype = null;
			String Email1displayname = null;
			String Email1address = null;
			String Email1fax = null;
			String Email2addresstype = null;
			String Email2displayname = null;
			String Email2address = null;
			String Email2fax = null;
			String Email3addresstype = null;
			String Email3displayname = null;
			String Email3address = null;
			String Email3fax = null;
			String homefaxaddresstype = null;
			String homefaxdisplayname = null;
			String homefaxaddress = null;
			String homefaxno = null;
			String primaryfaxdisplayname = null;
			String primaryfaxaddress = null;
			String primaryfaxno = null;
			String primaryfaxaddresstype = null;
			String bussinessfaxaddresstype = null;
			String bussinessfaxaddress = null;
			String bussinessfaxno = null;
			String bussinessfaxdisplayname = null;
			String birthday = null;
			String WeddingAnniversary = null;
			String firstName = null;
			String middleName = null;
			String lastname = null;
			String fileunder = null;
			String fileunderid = null;
			String Suffix = null;
			String Job_title = null;
			String title = null;
			String account = null;
			String BusinessHomePage = null;
			String ComputerNetworkName = null;
			String CustomerId = null;
			String FreeBusyLocation = null;
			String FtpSite = null;
			String Gender = null;
			String GovernmentIdNumber = null;
			String Hobbies = null;
			String Html = null;
			String InstantMessagingAddress = null;
			String Language = null;
			String Location = null;
			String Notes = null;
			String OrganizationalIdNumber = null;
			String PersonalHomePage = null;
			String ReferredByName = null;
			String SpouseName = null;
			String homeAddress = null;
			String homeCity = null;
			String homeCountry = null;
			String homeCountryCode = null;
			String homePostalCode = null;
			String homegetPostOfficeBox = null;
			String homeStateOrProvince = null;
			String homegetStreet = null;
			String otherAddress = null;
			String otherCity = null;
			String otherCountry = null;
			String otherCountryCode = null;
			String otherPostalCode = null;
			String othergetPostOfficeBox = null;
			String otherStateOrProvince = null;
			String othergetStreet = null;
			String workAddress = null;
			String workCity = null;
			String workCountry = null;
			String workCountryCode = null;
			String workPostalCode = null;
			String workgetPostOfficeBox = null;
			String workStateOrProvince = null;
			String workgetStreet = null;
			String Assistant = null;
			String CompanyName = null;
			String DepartmentName = null;
			String ManagerName = null;
			String OfficeLocation = null;
			String Profession = null;
			String getTitle = null;
			String AssistantTelephoneNumber = null;
			String Business2TelephoneNumber = null;
			String BusinessTelephoneNumber = null;
			String CallbackTelephoneNumber = null;
			String CarTelephoneNumber = null;
			String CompanyMainTelephoneNumber = null;
			String Home2TelephoneNumber = null;
			String HomeTelephoneNumber = null;
			String IsdnNumber = null;
			String MobileTelephoneNumber = null;
			String OtherTelephoneNumber = null;
			String PagerTelephoneNumber = null;
			String PrimaryTelephoneNumber = null;
			String RadioTelephoneNumber = null;
			String TelexNumber = null;
			String TtyTddPhoneNumber = null;
			String[] getCategories = null;
			String[] getChildren = null;
			String Subject = null;
			String Initials = null;
			String Billing_Information = null;
			String fname;
			String Phone_1_Selected = null;
			String Phone_1_Selector = "Business";
			String Phone_2_Selected = null;
			String Phone_2_Selector = "Home";
			String Phone_3_Selected = null;
			String Phone_3_Selector = "Business Fax";
			String Phone_4_Selected = null;
			String Phone_4_Selector = "Mobile";
			String Phone_5_Selected = null;
			String Phone_5_Selector = "Radio";
			String Phone_6_Selected = null;
			String Phone_6_Selector = "Car";
			String Phone_7_Selected = null;
			String Phone_7_Selector = "Other";
			String Phone_8_Selected = null;
			String Phone_8_Selector = "ISDN";
			String Message_Class = null;
			String ZIP_Postal_Code = null;
			String Mailing_Address = null;
			String Mailing_Address_Indicator = null;
			String Address_Selected = "";
			String Address_Selector = "Business";
			String Email_Selected = "";
			String Email_Selector = "Email";
			String Attachment = "NO attachment";
			String Size_on_Server = "";
			String Web_Page = null;
			String Journal = null;
			String Private = null;
			String reminder = "No Reminder";
			String Icon = "";
			String folder_Name = fOLdER;
			String Other_Fax = fOLdER;
			String User_Field_1 = null;
			String User_Field_2 = null;
			String User_Field_3 = null;
			String User_Field_4 = null;
			String Reminder_Topic = null;
			String Reminder_Time = null;
			String Outlook_Data_File = "";
			String Follow_Up_Flag = null;
			String Flag_Completed_Date = "Na";
			String Flag_Status = "Na";
			String Contacts = "";
			String Modified = "";
			String Internet_Free_Busy_Address = "";

			MapiContact con = (MapiContact) message.toMapiMessageItem();

			FollowUpOptions followUpOptions = FollowUpManager.getOptions(message);
			try {
				Follow_Up_Flag = followUpOptions.getFlagRequest();
			} catch (Exception ep) {
				Follow_Up_Flag = "";
			}
			try {
				if (Follow_Up_Flag.equalsIgnoreCase("null") || Follow_Up_Flag.contains("meta")
						|| Follow_Up_Flag.contains("aspose")) {
					Follow_Up_Flag = "NA";
				}
			} catch (Exception e1) {
				Follow_Up_Flag = "NA";
			}
			String fullday = String.valueOf(message.getDeliveryTime());
			MapiContactOtherPropertySet co = con.getOtherFields();
			try {
				User_Field_1 = co.getUserField1();
			} catch (Exception ep) {
				User_Field_1 = "";
			}
			try {
				if (User_Field_1.equalsIgnoreCase("null") || User_Field_1.contains("meta")
						|| User_Field_1.contains("aspose")) {
					User_Field_1 = "NA";
				}
			} catch (Exception e1) {
				User_Field_1 = "NA";
			}
			try {
				User_Field_2 = co.getUserField2();
			} catch (Exception ep) {
				User_Field_2 = "";
			}
			try {
				if (User_Field_2.equalsIgnoreCase("null") || User_Field_2.contains("meta")
						|| User_Field_2.contains("aspose")) {
					User_Field_2 = "NA";
				}
			} catch (Exception e1) {
				User_Field_2 = "NA";
			}
			try {
				User_Field_3 = co.getUserField3();
			} catch (Exception ep) {
				User_Field_3 = "";
			}
			try {
				if (User_Field_3.equalsIgnoreCase("null") || User_Field_3.contains("meta")
						|| User_Field_3.contains("aspose")) {
					User_Field_3 = "NA";
				}
			} catch (Exception e1) {
				User_Field_3 = "NA";
			}
			try {
				User_Field_4 = co.getUserField4();
			} catch (Exception ep) {
				User_Field_4 = "";
			}
			try {
				if (User_Field_4.equalsIgnoreCase("null") || User_Field_4.contains("meta")
						|| User_Field_4.contains("aspose")) {
					User_Field_4 = "NA";
				}
			} catch (Exception e1) {
				User_Field_4 = "NA";
			}

			try {
				Reminder_Topic = co.getReminderTopic();
			} catch (Exception ep) {
				Reminder_Topic = "";
			}
			try {
				if (Reminder_Topic.equalsIgnoreCase("null") || Reminder_Topic.contains("meta")
						|| Reminder_Topic.contains("aspose")) {
					Reminder_Topic = "NA";
				}
			} catch (Exception e1) {
				Reminder_Topic = "NA";
			}

			try {
				Reminder_Time = String.valueOf(co.getReminderTime());
			} catch (Exception ep) {
				Reminder_Time = "";
			}
			try {
				if (Reminder_Time.equalsIgnoreCase("null") || Reminder_Time.contains("meta")
						|| Reminder_Time.contains("aspose")) {
					Reminder_Time = "NA";
				}
			} catch (Exception e1) {
				Reminder_Time = "NA";
			}

			if (co.getPrivate()) {
				Private = "Yes";
			} else {
				Private = "No";
			}
			if (co.getJournal()) {
				Journal = "Yes";
			} else {
				Journal = "No";
			}
			int sen;
			String sencitivity = "null";
			try {
				sen = message.getSensitivity();
				if (sen == 0) {
					sencitivity = "NORMAL";

				}
			} catch (Exception e) {

			}

			try {
				Message_Class = message.getMessageClass();
			} catch (Exception ep) {
				Message_Class = "";
			}
			try {
				if (Message_Class.equalsIgnoreCase("null") || Message_Class.contains("meta")
						|| Message_Class.contains("aspose")) {
					Message_Class = "NA";
				}
			} catch (Exception e1) {
				Message_Class = "NA";
			}

			try {
				Subject = con.getSubject();
			} catch (Exception ep) {
				Subject = "";
			}
			try {
				if (Subject.equalsIgnoreCase("null") || Subject.contains("meta") || Subject.contains("aspose")) {
					Subject = "NA";
				}
			} catch (Exception e1) {
				Subject = "NA";
			}

			MapiContactProfessionalPropertySet ProfPropSet = null;
			try {
				ProfPropSet = con.getProfessionalInfo();
			} catch (Exception e) {

			}

			try {
				Subject = ProfPropSet.getTitle();
			} catch (Exception ep) {
				Subject = "";
			}
			try {
				if (Subject.equalsIgnoreCase("null") || Subject.contains("meta") || Subject.contains("aspose")) {
					Subject = "NA";
				}
			} catch (Exception e1) {
				Subject = "NA";
			}

			MapiContactElectronicAddress email1 = null;

			try {
				email1 = con.getElectronicAddresses().getEmail1();
			} catch (Exception e) {

			}
			try {
				getCategories = con.getCategories();
			} catch (Exception e1) {

			}
			try {
				for (int i = 0; i < getCategories.length; i++) {
					if (i == 0) {
						Categories = getCategories[i];
					} else {
						Categories = Categories + "," + getCategories[i];
					}
				}

			} catch (Exception e) {
				Categories = "";
			}
			try {
				if (Categories.equalsIgnoreCase("null") || Categories.contains("meta")
						|| Categories.contains("aspose")) {
					Categories = "NA";
				}
			} catch (Exception e1) {
				Categories = "NA";
			}
			MapiContactElectronicAddress email2 = null;
			try {
				email2 = con.getElectronicAddresses().getEmail2();
			} catch (Exception e) {

			}
			MapiContactElectronicAddress email3 = null;
			try {
				email3 = con.getElectronicAddresses().getEmail3();
			} catch (Exception e) {

			}
			MapiContactElectronicAddress homefax = null;
			try {
				homefax = con.getElectronicAddresses().getHomeFax();
			} catch (Exception e) {

			}
			MapiContactElectronicAddress primaryfax = null;
			try {
				primaryfax = con.getElectronicAddresses().getPrimaryFax();
			} catch (Exception e) {

			}
			try {
				Other_Fax = primaryfax.getFaxNumber();
			} catch (Exception e) {
				Other_Fax = "";
			}
			try {
				if (Other_Fax.equalsIgnoreCase("null") || Other_Fax.contains("meta") || Other_Fax.contains("aspose")) {
					Other_Fax = "NA";
				}
			} catch (Exception e1) {
				Other_Fax = "NA";
			}
			MapiContactElectronicAddress bussinessfax = null;
			try {
				bussinessfax = con.getElectronicAddresses().getBusinessFax();
			} catch (Exception e) {

			}
			MapiContactEventPropertySet event = null;
			try {
				event = con.getEvents();
			} catch (Exception e) {

			}
			MapiContactPersonalInfoPropertySet personfPropSet = null;
			try {
				personfPropSet = con.getPersonalInfo();
			} catch (Exception e) {

			}

			try {
				getChildren = personfPropSet.getChildren();
			} catch (Exception e1) {

			}

			try {
				for (int i = 0; i < getChildren.length; i++) {
					if (i == 0) {
						Children = getChildren[i];
					} else {
						Children = Children + "," + getChildren[i];
					}
				}

			} catch (Exception e) {
				Children = "";
			}
			try {
				if (Children.equalsIgnoreCase("null") || Children.contains("meta") || Children.contains("aspose")) {
					Children = "NA";
				}
			} catch (Exception e1) {
				Children = "NA";
			}

			try {
				Billing_Information = String.valueOf(con.getBilling());
			} catch (Exception e) {
				Billing_Information = "";
			}
			try {
				if (Billing_Information.equalsIgnoreCase("null") || Billing_Information.contains("meta")
						|| Billing_Information.contains("aspose")) {
					Billing_Information = "NA";
				}
			} catch (Exception e1) {
				Billing_Information = "NA";
			}
			try {
				mileage = String.valueOf(con.getMileage());
			} catch (Exception e) {
				mileage = "";
			}
			try {
				if (mileage.equalsIgnoreCase("null") || mileage.contains("meta") || mileage.contains("aspose")) {
					mileage = "NA";
				}
			} catch (Exception e1) {
				mileage = "NA";
			}

			try {
				account = String.valueOf(personfPropSet.getAccount());
			} catch (Exception ep) {
				account = "";
			}
			try {
				if (account.equalsIgnoreCase("null") || account.contains("meta") || account.contains("aspose")) {
					account = "NA";
				}
			} catch (Exception e1) {
				account = "NA";
			}
			try {
				BusinessHomePage = String.valueOf(personfPropSet.getBusinessHomePage());
			} catch (Exception ep) {
				BusinessHomePage = "";
			}
			try {
				if (BusinessHomePage.equalsIgnoreCase("null") || BusinessHomePage.contains("meta")
						|| BusinessHomePage.contains("aspose")) {
					BusinessHomePage = "NA";
				}
			} catch (Exception e1) {
				BusinessHomePage = "NA";
			}
			Web_Page = BusinessHomePage;
			try {
				ComputerNetworkName = String.valueOf(personfPropSet.getComputerNetworkName());
			} catch (Exception ep) {
				ComputerNetworkName = "";
			}
			try {
				if (ComputerNetworkName.equalsIgnoreCase("null") || ComputerNetworkName.contains("meta")
						|| ComputerNetworkName.contains("aspose")) {
					ComputerNetworkName = "NA";
				}
			} catch (Exception e1) {
				ComputerNetworkName = "NA";
			}
			try {

				CustomerId = String.valueOf(personfPropSet.getCustomerId());
			} catch (Exception ep) {
				CustomerId = "";
			}
			try {
				if (CustomerId.equalsIgnoreCase("null") || CustomerId.contains("meta")
						|| CustomerId.contains("aspose")) {
					CustomerId = "NA";
				}
			} catch (Exception e1) {
				CustomerId = "NA";
			}
			try {
				FreeBusyLocation = String.valueOf(personfPropSet.getFreeBusyLocation());
			} catch (Exception ep) {
				FreeBusyLocation = "";
			}
			try {
				if (FreeBusyLocation.equalsIgnoreCase("null") || FreeBusyLocation.contains("meta")
						|| FreeBusyLocation.contains("aspose")) {
					FreeBusyLocation = "NA";
				}
			} catch (Exception e1) {
				FreeBusyLocation = "NA";
			}
			try {

				FtpSite = String.valueOf(personfPropSet.getFtpSite());
			} catch (Exception ep) {
				FtpSite = "";
			}
			try {
				if (FtpSite.equalsIgnoreCase("null") || FtpSite.contains("meta") || FtpSite.contains("aspose")) {
					FtpSite = "NA";
				}
			} catch (Exception e1) {
				FtpSite = "NA";
			}
			try {
				int i = personfPropSet.getGender();
				if (i == 0) {
					Gender = "Unspecified";
				} else if (i == 1) {
					Gender = "Female";
				} else if (i == 2) {
					Gender = "Male";
				}

			} catch (Exception ep) {
				Gender = "";
			}

			try {
				GovernmentIdNumber = String.valueOf(personfPropSet.getGovernmentIdNumber());
			} catch (Exception ep) {
				GovernmentIdNumber = "";
			}
			try {
				if (GovernmentIdNumber.equalsIgnoreCase("null") || GovernmentIdNumber.contains("meta")
						|| GovernmentIdNumber.contains("aspose")) {
					GovernmentIdNumber = "NA";
				}
			} catch (Exception e1) {
				GovernmentIdNumber = "NA";
			}
			try {
				Hobbies = String.valueOf(personfPropSet.getHobbies());
			} catch (Exception ep) {
				Hobbies = "";
			}
			try {
				if (Hobbies.equalsIgnoreCase("null") || Hobbies.contains("meta") || Hobbies.contains("aspose")) {
					Hobbies = "NA";
				}
			} catch (Exception e1) {
				Hobbies = "NA";
			}
			try {

				Html = String.valueOf(personfPropSet.getHtml());
			} catch (Exception ep) {
				Html = "";
			}
			try {
				if (Html.equalsIgnoreCase("null") || Html.contains("aspose")) {
					Html = "NA";
				}
			} catch (Exception e1) {
				Html = "NA";
			}
			try {

				InstantMessagingAddress = String.valueOf(personfPropSet.getInstantMessagingAddress());
			} catch (Exception ep) {
				InstantMessagingAddress = "";
			}
			try {
				if (InstantMessagingAddress.equalsIgnoreCase("null") || InstantMessagingAddress.contains("meta")
						|| InstantMessagingAddress.contains("aspose")) {
					InstantMessagingAddress = "NA";
				}
			} catch (Exception e1) {
				InstantMessagingAddress = "NA";
			}

			try {

				Language = String.valueOf(personfPropSet.getLanguage());
			} catch (Exception ep) {
				Language = "";
			}
			try {
				if (Language.equalsIgnoreCase("null") || Language.contains("meta") || Language.contains("aspose")) {
					Language = "NA";
				}
			} catch (Exception e1) {
				Language = "NA";
			}
			try {

				Location = String.valueOf(personfPropSet.getLocation());
			} catch (Exception ep) {
				Location = "";
			}
			try {
				if (Location.equalsIgnoreCase("null") || Location.contains("meta") || Location.contains("aspose")) {
					Location = "NA";
				}
			} catch (Exception e1) {
				Location = "NA";
			}
			try {

				Notes = String.valueOf(personfPropSet.getNotes());
			} catch (Exception ep) {
				Notes = "";
			}
			try {
				if (Notes.equalsIgnoreCase("null") || Notes.contains("meta") || Notes.contains("aspose")) {
					Notes = "NA";
				}
			} catch (Exception e1) {
				Notes = "NA";
			}
			try {

				OrganizationalIdNumber = String.valueOf(personfPropSet.getOrganizationalIdNumber());
			} catch (Exception ep) {
				OrganizationalIdNumber = "";
			}
			try {
				if (OrganizationalIdNumber.equalsIgnoreCase("null") || OrganizationalIdNumber.contains("meta")
						|| OrganizationalIdNumber.contains("aspose")) {
					OrganizationalIdNumber = "NA";
				}
			} catch (Exception e1) {
				OrganizationalIdNumber = "NA";
			}
			try {

				PersonalHomePage = String.valueOf(personfPropSet.getPersonalHomePage());
			} catch (Exception ep) {
				PersonalHomePage = "";
			}
			try {
				if (PersonalHomePage.equalsIgnoreCase("null") || PersonalHomePage.contains("meta")
						|| PersonalHomePage.contains("aspose")) {
					PersonalHomePage = "NA";
				}
			} catch (Exception e1) {
				PersonalHomePage = "NA";
			}
			try {

				ReferredByName = String.valueOf(personfPropSet.getReferredByName());
			} catch (Exception ep) {
				ReferredByName = "";
			}
			try {
				if (ReferredByName.equalsIgnoreCase("null") || ReferredByName.contains("meta")
						|| ReferredByName.contains("aspose")) {
					ReferredByName = "NA";
				}
			} catch (Exception e1) {
				ReferredByName = "NA";
			}
			try {

				SpouseName = String.valueOf(personfPropSet.getSpouseName());
			} catch (Exception ep) {
				SpouseName = "";
			}
			try {
				if (SpouseName.equalsIgnoreCase("null") || SpouseName.contains("meta")
						|| SpouseName.contains("aspose")) {
					SpouseName = "NA";
				}
			} catch (Exception e1) {
				SpouseName = "NA";
			}

			MapiContactProfessionalPropertySet personPropSet = null;
			try {
				personPropSet = con.getProfessionalInfo();
			} catch (Exception e) {

			}

			MapiContactPhysicalAddressPropertySet mapipcs = null;

			try {
				mapipcs = con.getPhysicalAddresses();
			} catch (Exception e) {

			}
			MapiContactTelephonePropertySet mapitelephone = null;
			try {
				mapitelephone = con.getTelephones();
			} catch (Exception e) {

			}
			MapiContactPhysicalAddress contacthomephys = null;
			try {
				contacthomephys = mapipcs.getHomeAddress();
			} catch (Exception e) {

			}
			MapiContactPhysicalAddress contactotherphys = null;

			try {
				contactotherphys = mapipcs.getOtherAddress();
			} catch (Exception e) {

			}
			MapiContactPhysicalAddress contactworkphys = null;
			try {
				contactworkphys = mapipcs.getWorkAddress();
			} catch (Exception e) {

			}
			MapiContactNamePropertySet NamePropSet = null;
			System.out.println(con.getNameInfo().getGeneration());
			try {
				NamePropSet = con.getNameInfo();
			} catch (Exception e) {

			}

			try {

				Initials = NamePropSet.getInitials();
			} catch (Exception ep) {
				Initials = "";
			}
			try {
				if (Initials.equalsIgnoreCase("null") || Initials.contains("meta") || Initials.contains("aspose")) {
					Initials = "NA";
				}
			} catch (Exception e1) {
				Initials = "NA";
			}
			try {

				Job_title = ProfPropSet.getTitle();
			} catch (Exception ep) {
				Job_title = "";
			}
			try {
				if (Job_title.equalsIgnoreCase("null") || Job_title.contains("meta") || Job_title.contains("aspose")) {
					Job_title = "NA";
				}
			} catch (Exception e1) {
				Job_title = "NA";
			}

			try {
				firstName = NamePropSet.getGivenName();
			} catch (Exception ep) {
				firstName = "";
			}
			try {
				if (firstName.equalsIgnoreCase("null") || firstName.contains("meta") || firstName.contains("aspose")) {
					firstName = "NA";
				}
			} catch (Exception e1) {
				firstName = "NA";
			}

			try {
				nickName = NamePropSet.getNickname();
			} catch (Exception ep) {
				nickName = "";
			}
			try {
				if (nickName.equalsIgnoreCase("null") || nickName.contains("meta") || nickName.contains("aspose")) {
					nickName = "NA";
				}
			} catch (Exception e1) {
				nickName = "NA";
			}

			try {

				middleName = NamePropSet.getMiddleName();
			} catch (Exception ep) {
				middleName = "";
			}
			try {
				if (middleName.equalsIgnoreCase("null") || middleName.contains("meta")
						|| middleName.contains("aspose")) {
					middleName = "NA";
				}
			} catch (Exception e1) {
				middleName = "NA";
			}
			try {

				lastname = NamePropSet.getSurname();
			} catch (Exception ep) {
				lastname = "";
			}
			try {
				if (lastname.equalsIgnoreCase("null") || lastname.contains("meta") || lastname.contains("aspose")) {
					lastname = "NA";
				}
			} catch (Exception e1) {
				lastname = "NA";
			}
			try {

				Email1addresstype = email1.getAddressType();
			} catch (Exception ep) {
				Email1addresstype = "";
			}
			try {
				if (Email1addresstype.equalsIgnoreCase("null") || Email1addresstype.contains("meta")
						|| Email1addresstype.contains("aspose")) {
					Email1addresstype = "NA";
				}
			} catch (Exception e1) {
				Email1addresstype = "";
			}
			try {

				Email1displayname = email1.getDisplayName();
			} catch (Exception ep) {
				Email1displayname = "";
			}
			try {
				if (Email1displayname.equalsIgnoreCase("null") || Email1displayname.contains("meta")
						|| Email1displayname.contains("aspose")) {
					Email1displayname = "NA";
				}
			} catch (Exception e1) {
				Email1displayname = "";
			}

			int attsize = message.getAttachments().size();

			if (attsize > 0) {
				Attachment = "With attachment";

			}

			try {

				Email1address = email1.getEmailAddress();
			} catch (Exception ep) {
				Email1address = "";
			}
			try {
				if (Email1address.equalsIgnoreCase("null") || Email1address.contains("meta")
						|| Email1address.contains("aspose")) {

				}
			} catch (Exception e1) {
				Email1address = "NA";
			}
			try {

				Email1fax = email1.getFaxNumber();
			} catch (Exception ep) {
				Email1fax = "";
			}
			try {
				if (Email1fax.equalsIgnoreCase("null") || Email1fax.contains("meta") || Email1fax.contains("aspose")) {
					Email1fax = "NA";
				}
			} catch (Exception e1) {
				Email1fax = "NA";
			}
			try {

				Email2addresstype = email2.getAddressType();
			} catch (Exception ep) {
				Email2addresstype = "";
			}
			try {
				if (Email2addresstype.equalsIgnoreCase("null") || Email2addresstype.contains("meta")
						|| Email2addresstype.contains("aspose")) {
					Email2addresstype = "NA";
				}
			} catch (Exception e1) {
				Email2addresstype = "NA";
			}
			try {

				Email2displayname = email2.getDisplayName();
			} catch (Exception ep) {
				Email2displayname = "";
			}
			try {
				if (Email2displayname.equalsIgnoreCase("null") || Email2displayname.contains("meta")
						|| Email2displayname.contains("aspose")) {
					Email2displayname = "NA";
				}
			} catch (Exception e1) {
				Email2displayname = "NA";
			}
			try {

				Email2address = email2.getEmailAddress();
			} catch (Exception ep) {
				Email1address = "";
			}
			try {
				if (Email1address.equalsIgnoreCase("null") || Email1address.contains("meta")
						|| Email1address.contains("aspose")) {
					Email1address = "NA";
				}
			} catch (Exception e1) {
				Email1address = "NA";
			}
			try {

				Email2fax = email2.getFaxNumber();
			} catch (Exception ep) {
				Email2fax = "";
			}
			try {
				if (Email2fax.equalsIgnoreCase("null") || Email2fax.contains("meta") || Email2fax.contains("aspose")) {
					Email2fax = "NA";
				}
			} catch (Exception e1) {
				Email2fax = "NA";
			}
			try {

				Email3addresstype = email3.getAddressType();
			} catch (Exception ep) {
				Email3addresstype = "";
			}
			try {
				if (Email3addresstype.equalsIgnoreCase("null") || Email3addresstype.contains("meta")
						|| Email3addresstype.contains("aspose")) {
					Email3addresstype = "NA";
				}
			} catch (Exception e1) {
				Email3addresstype = "NA";
			}
			try {
				Email3displayname = email3.getDisplayName();
			} catch (Exception ep) {
				Email3displayname = "";
			}
			try {
				if (Email3displayname.equalsIgnoreCase("null") || Email3displayname.contains("meta")
						|| Email3displayname.contains("aspose")) {
					Email3displayname = "NA";
				}
			} catch (Exception e1) {
				Email3displayname = "NA";
			}
			try {
				Email3address = email3.getEmailAddress();
			} catch (Exception ep) {
				Email3address = "";
			}
			try {
				if (Email3address.equalsIgnoreCase("null") || Email3address.contains("meta")
						|| Email3address.contains("aspose")) {
					Email3address = "NA";
				}
			} catch (Exception e1) {
				Email3address = "NA";
			}
			try {
				Email3fax = email3.getFaxNumber();
			} catch (Exception ep) {
				Email3fax = "";
			}
			try {
				if (Email3fax.equalsIgnoreCase("null") || Email3fax.contains("meta") || Email3fax.contains("aspose")) {
					Email3fax = "NA";
				}
			} catch (Exception e1) {
				Email3fax = "NA";
			}
			try {
				homefaxaddresstype = homefax.getAddressType();
			} catch (Exception ep) {
				homefaxaddresstype = "";
			}
			try {
				if (homefaxaddresstype.equalsIgnoreCase("null") || homefaxaddresstype.contains("meta")
						|| homefaxaddresstype.contains("aspose")) {
					homefaxaddresstype = "NA";
				}
			} catch (Exception e1) {
				homefaxaddresstype = "NA";
			}
			try {
				homefaxdisplayname = homefax.getDisplayName();
			} catch (Exception ep) {
				homefaxdisplayname = "";
			}
			try {
				if (homefaxdisplayname.equalsIgnoreCase("null") || homefaxdisplayname.contains("meta")
						|| homefaxdisplayname.contains("aspose")) {
					homefaxdisplayname = "NA";
				}
			} catch (Exception e1) {
				homefaxdisplayname = "NA";
			}
			try {

				homefaxaddress = homefax.getEmailAddress();
			} catch (Exception ep) {
				homefaxaddress = "";
			}
			try {
				if (homefaxaddress.equalsIgnoreCase("null") || homefaxaddress.contains("meta")
						|| homefaxaddress.contains("aspose")) {
					homefaxaddress = "NA";
				}
			} catch (Exception e1) {
				homefaxaddress = "NA";
			}
			try {

				homefaxno = homefax.getFaxNumber();
			} catch (Exception ep) {
				homefaxno = "";
			}
			try {
				if (homefaxno.equalsIgnoreCase("null") || homefaxno.contains("meta") || homefaxno.contains("aspose")) {
					homefaxno = "NA";
				}
			} catch (Exception e1) {
				homefaxno = "NA";
			}
			try {

				primaryfaxaddresstype = primaryfax.getAddressType();
			} catch (Exception ep) {
				primaryfaxaddresstype = "";
			}
			try {
				if (primaryfaxaddresstype.equalsIgnoreCase("null") || primaryfaxaddresstype.contains("meta")
						|| primaryfaxaddresstype.contains("aspose")) {
					primaryfaxaddresstype = "NA";
				}
			} catch (Exception e1) {
				primaryfaxaddresstype = "NA";
			}
			try {
				primaryfaxdisplayname = primaryfax.getDisplayName();
			} catch (Exception ep) {
				primaryfaxdisplayname = "";
			}
			try {
				if (primaryfaxdisplayname.equalsIgnoreCase("null") || primaryfaxdisplayname.contains("meta")
						|| primaryfaxdisplayname.contains("aspose")) {
					primaryfaxdisplayname = "NA";
				}
			} catch (Exception e1) {
				primaryfaxdisplayname = "NA";
			}
			try {

				primaryfaxaddress = primaryfax.getEmailAddress();
			} catch (Exception ep) {
				primaryfaxaddress = "";
			}
			try {
				if (primaryfaxaddress.equalsIgnoreCase("null") || primaryfaxaddress.contains("meta")
						|| primaryfaxaddress.contains("aspose")) {
					primaryfaxaddress = "NA";
				}
			} catch (Exception e1) {
				primaryfaxaddress = "NA";
			}
			try {

				primaryfaxno = primaryfax.getFaxNumber();
			} catch (Exception ep) {
				primaryfaxno = "";
			}
			try {
				if (primaryfaxno.equalsIgnoreCase("null") || primaryfaxno.contains("meta")
						|| primaryfaxno.contains("aspose")) {
					primaryfaxno = "NA";
				}
			} catch (Exception e1) {
				primaryfaxno = "NA";
			}
			try {
				bussinessfaxaddresstype = bussinessfax.getFaxNumber();
			} catch (Exception ep) {
				bussinessfaxaddresstype = "";
			}
			try {
				if (bussinessfaxaddresstype.equalsIgnoreCase("null") || bussinessfaxaddresstype.contains("meta")
						|| bussinessfaxaddresstype.contains("aspose")) {
					bussinessfaxaddresstype = "NA";
				}
			} catch (Exception e1) {
				bussinessfaxaddresstype = "NA";
			}
			try {

				bussinessfaxdisplayname = bussinessfax.getDisplayName();
			} catch (Exception ep) {
				bussinessfaxdisplayname = "";
			}
			try {
				if (bussinessfaxdisplayname.equalsIgnoreCase("null") || bussinessfaxdisplayname.contains("meta")
						|| bussinessfaxdisplayname.contains("aspose")) {
					bussinessfaxdisplayname = "NA";
				}
			} catch (Exception e1) {
				bussinessfaxdisplayname = "NA";
			}
			try {

				bussinessfaxaddress = bussinessfax.getEmailAddress();
			} catch (Exception ep) {
				bussinessfaxaddress = "";
			}
			try {
				if (bussinessfaxaddress.equalsIgnoreCase("null") || bussinessfaxaddress.contains("meta")
						|| bussinessfaxaddress.contains("aspose")) {
					bussinessfaxaddress = "NA";
				}
			} catch (Exception e1) {
				bussinessfaxaddress = "NA";
			}
			try {

				bussinessfaxno = bussinessfax.getFaxNumber();
			} catch (Exception ep) {
				bussinessfaxno = "";
			}
			try {
				if (bussinessfaxno.equalsIgnoreCase("null") || bussinessfaxno.contains("meta")
						|| bussinessfaxno.contains("aspose")) {
					bussinessfaxno = "NA";
				}
			} catch (Exception e1) {
				bussinessfaxno = "NA";
			}
			Phone_3_Selected = bussinessfaxno;
			try {

				birthday = String.valueOf(event.getBirthday());
			} catch (Exception ep) {
				birthday = "";
			}
			try {
				if (birthday.equalsIgnoreCase("null") || birthday.contains("meta") || birthday.contains("aspose")) {
					birthday = "NA";
				}
			} catch (Exception e1) {
				birthday = "NA";
			}
			try {

				WeddingAnniversary = String.valueOf(event.getWeddingAnniversary());
			} catch (Exception ep) {
				WeddingAnniversary = "";
			}
			try {
				if (WeddingAnniversary.equalsIgnoreCase("null") || WeddingAnniversary.contains("meta")
						|| WeddingAnniversary.contains("aspose")) {
					WeddingAnniversary = "NA";
				}
			} catch (Exception e1) {
				WeddingAnniversary = "NA";
			}
			try {

				Email1addresstype = email1.getAddressType();
			} catch (Exception ep) {
				Email1addresstype = "";
			}
			try {
				if (Email1addresstype.equalsIgnoreCase("null") || Email1addresstype.contains("meta")
						|| firstName.contains("aspose")) {
					Email1addresstype = "NA";
				}
			} catch (Exception e1) {
				Email1addresstype = "NA";
			}
			try {

				title = NamePropSet.getDisplayNamePrefix();
			} catch (Exception ep) {
				title = "";
			}
			try {
				if (title.equalsIgnoreCase("null") || title.contains("meta") || title.contains("aspose")) {
					title = "NA";
				}
			} catch (Exception e1) {
				title = "NA";
			}
			try {
				fileunder = NamePropSet.getFileUnder();
			} catch (Exception ep) {
				fileunder = "";
			}
			try {
				if (fileunder.equalsIgnoreCase("null") || fileunder.contains("meta") || fileunder.contains("aspose")) {
					fileunder = "NA";
				}
			} catch (Exception e1) {
				fileunder = "NA";
			}
			try {
				fileunderid = String.valueOf(NamePropSet.getFileUnderID());
			} catch (Exception ep) {
				fileunderid = "";
			}
			try {
				if (fileunderid.equalsIgnoreCase("null") || fileunderid.contains("meta")
						|| fileunderid.contains("aspose")) {
					fileunderid = "NA";
				}
			} catch (Exception e1) {
				fileunderid = "NA";
			}
			try {
				Suffix = String.valueOf(NamePropSet.getGeneration());
			} catch (Exception ep) {
				Suffix = "";
			}
			try {
				if (Suffix.equalsIgnoreCase("null") || Suffix.contains("meta") || Suffix.contains("aspose")) {
					Suffix = "NA";
				}
			} catch (Exception e1) {
				Suffix = "NA";
			}

			try {

				homegetStreet = String.valueOf(contacthomephys.getStreet());
			} catch (Exception ep) {
				homegetStreet = "";
			}
			try {
				if (homegetStreet.equalsIgnoreCase("null") || homegetStreet.contains("meta")
						|| homegetStreet.contains("aspose")) {
					homegetStreet = "NA";
				}
			} catch (Exception e1) {
				homegetStreet = "NA";
			}
			try {

				homeAddress = String.valueOf(contacthomephys.getAddress());
			} catch (Exception ep) {
				homeAddress = "";
			}
			try {
				if (homeAddress.equalsIgnoreCase("null") || homeAddress.contains("meta")
						|| homeAddress.contains("aspose")) {
					homeAddress = "NA";
				}
			} catch (Exception e1) {
				homeAddress = "NA";
			}
			try {

				homeCountry = String.valueOf(contacthomephys.getCountry());
			} catch (Exception ep) {
				homeCountry = "";
			}
			try {
				if (homeCountry.equalsIgnoreCase("null") || homeCountry.contains("meta")
						|| homeCountry.contains("aspose")) {
					homeCountry = "NA";
				}
			} catch (Exception e1) {
				homeCountry = "NA";
			}
			try {

				homeCountryCode = String.valueOf(contacthomephys.getCountryCode());
			} catch (Exception ep) {
				homeCountryCode = "";
			}
			try {
				if (homeCountryCode.equalsIgnoreCase("null") || homeCountryCode.contains("meta")
						|| homeCountryCode.contains("aspose")) {
					homeCountryCode = "NA";
				}
			} catch (Exception e1) {
				homeCountryCode = "NA";
			}
			try {

				homePostalCode = String.valueOf(contacthomephys.getPostalCode());
			} catch (Exception ep) {
				homePostalCode = "";
			}
			try {
				if (homePostalCode.equalsIgnoreCase("null") || homePostalCode.contains("meta")
						|| homePostalCode.contains("aspose")) {
					homePostalCode = "NA";
				}
			} catch (Exception e1) {
				homePostalCode = "NA";
			}
			try {

				homegetPostOfficeBox = String.valueOf(contacthomephys.getPostOfficeBox());
			} catch (Exception ep) {
				homegetPostOfficeBox = "";
			}
			try {
				if (homegetPostOfficeBox.equalsIgnoreCase("null") || homegetPostOfficeBox.contains("meta")
						|| homegetPostOfficeBox.contains("aspose")) {
					homegetPostOfficeBox = "NA";
				}
			} catch (Exception e1) {
				homegetPostOfficeBox = "NA";
			}
			try {

				homeStateOrProvince = String.valueOf(contacthomephys.getStateOrProvince());
			} catch (Exception ep) {
				homeStateOrProvince = "";
			}
			try {
				if (homeStateOrProvince.equalsIgnoreCase("null") || homeStateOrProvince.contains("meta")
						|| homeStateOrProvince.contains("aspose")) {
					homeStateOrProvince = "NA";
				}
			} catch (Exception e1) {
				homeStateOrProvince = "NA";
			}
			try {

				othergetStreet = String.valueOf(contactotherphys.getStreet());
			} catch (Exception ep) {
				othergetStreet = "";
			}
			try {
				if (othergetStreet.equalsIgnoreCase("null") || othergetStreet.contains("meta")
						|| othergetStreet.contains("aspose")) {
					othergetStreet = "NA";
				}
			} catch (Exception e1) {
				othergetStreet = "NA";
			}
			try {
				otherAddress = String.valueOf(contactotherphys.getAddress());
			} catch (Exception ep) {
				otherAddress = "";
			}
			try {
				if (otherAddress.equalsIgnoreCase("null") || otherAddress.contains("meta")
						|| otherAddress.contains("aspose")) {
					otherAddress = "NA";
				}
			} catch (Exception e1) {
				otherAddress = "NA";
			}
			try {

				otherCity = String.valueOf(contactotherphys.getCity());
			} catch (Exception ep) {
				otherCity = "";
			}
			try {
				if (otherCity.equalsIgnoreCase("null") || otherCity.contains("meta") || otherCity.contains("aspose")) {
					otherCity = "NA";
				}
			} catch (Exception e1) {
				otherCity = "NA";
			}
			try {

				otherCountry = String.valueOf(contactotherphys.getCountry());
			} catch (Exception ep) {
				otherCountry = "";
			}
			try {
				if (otherCountry.equalsIgnoreCase("null") || otherCountry.contains("meta")
						|| otherCountry.contains("aspose")) {
					otherCountry = "NA";
				}
			} catch (Exception e1) {
				otherCountry = "NA";
			}
			try {

				otherCountryCode = String.valueOf(contactotherphys.getCountryCode());
			} catch (Exception ep) {
				otherCountryCode = "";
			}
			try {
				if (otherCountryCode.equalsIgnoreCase("null") || otherCountryCode.contains("meta")
						|| otherCountryCode.contains("aspose")) {
					otherCountryCode = "NA";
				}
			} catch (Exception e1) {
				otherCountryCode = "NA";
			}
			try {

				otherPostalCode = String.valueOf(contactotherphys.getPostalCode());
			} catch (Exception ep) {
				otherPostalCode = "";
			}
			try {
				if (otherPostalCode.equalsIgnoreCase("null") || otherPostalCode.contains("meta")
						|| otherPostalCode.contains("aspose")) {
					otherPostalCode = "NA";
				}
			} catch (Exception e1) {
				otherPostalCode = "NA";
			}
			try {

				othergetPostOfficeBox = String.valueOf(contactotherphys.getPostOfficeBox());
			} catch (Exception ep) {
				othergetPostOfficeBox = "";
			}
			try {
				if (othergetPostOfficeBox.equalsIgnoreCase("null") || othergetPostOfficeBox.contains("meta")
						|| othergetPostOfficeBox.contains("aspose")) {
					othergetPostOfficeBox = "NA";
				}
			} catch (Exception e1) {
				othergetPostOfficeBox = "NA";
			}
			try {

				otherStateOrProvince = String.valueOf(contactotherphys.getStateOrProvince());
			} catch (Exception ep) {
				otherStateOrProvince = "";
			}
			try {
				if (otherStateOrProvince.equalsIgnoreCase("null") || otherStateOrProvince.contains("meta")
						|| otherStateOrProvince.contains("aspose")) {
					otherStateOrProvince = "NA";
				}
			} catch (Exception e1) {
				otherStateOrProvince = "NA";
			}
			try {

				workgetStreet = String.valueOf(contactworkphys.getStreet());
			} catch (Exception ep) {
				workgetStreet = "";
			}
			try {
				if (workgetStreet.equalsIgnoreCase("null") || workgetStreet.contains("meta")
						|| workgetStreet.contains("aspose")) {
					workgetStreet = "NA";
				}
			} catch (Exception e1) {
				workgetStreet = "NA";
			}

			try {

				workAddress = String.valueOf(contactworkphys.getAddress());
			} catch (Exception ep) {
				workAddress = "";
			}
			try {
				if (workAddress.equalsIgnoreCase("null") || workAddress.contains("meta")
						|| workAddress.contains("aspose")) {
					workAddress = "NA";
				}
			} catch (Exception e1) {
				workAddress = "NA";
			}
			Mailing_Address = workAddress;
			Address_Selected = workAddress;
			try {
				if (contactworkphys.isMailingAddress()) {
					Mailing_Address_Indicator = "Yes";
				} else {
					Mailing_Address_Indicator = "No";
				}
			} catch (Exception e1) {

			}
			try {

				workCity = String.valueOf(contactworkphys.getCity());
			} catch (Exception ep) {
				workCity = "";
			}
			try {
				if (workCity.equalsIgnoreCase("null") || workCity.contains("meta") || workCity.contains("aspose")) {
					workCity = "NA";
				}
			} catch (Exception e1) {
				workCity = "NA";
			}

			try {
				workCountry = String.valueOf(contactworkphys.getCountry());
			} catch (Exception ep) {
				workCountry = "";
			}
			try {
				if (workCountry.equalsIgnoreCase("null") || workCountry.contains("meta")
						|| workCountry.contains("aspose")) {
					workCountry = "NA";
				}
			} catch (Exception e1) {
				workCountry = "NA";
			}
			try {

				workCountryCode = String.valueOf(contactworkphys.getCountryCode());
			} catch (Exception ep) {
				workCountryCode = "";
			}

			try {
				if (workCountryCode.equalsIgnoreCase("null") || workCountryCode.contains("meta")
						|| workCountryCode.contains("aspose")) {
					workCountryCode = "NA";
				}
			} catch (Exception e1) {
				workCountryCode = "NA";
			}
			try {

				workPostalCode = String.valueOf(contactworkphys.getPostalCode());
			} catch (Exception ep) {
				workPostalCode = "";
			}
			try {
				if (workPostalCode.equalsIgnoreCase("null") || workPostalCode.contains("meta")
						|| workPostalCode.contains("aspose")) {
					workPostalCode = "NA";
				}
			} catch (Exception e1) {
				workPostalCode = "NA";
			}
			ZIP_Postal_Code = workPostalCode;
			try {

				workgetPostOfficeBox = String.valueOf(contactworkphys.getPostOfficeBox());
			} catch (Exception ep) {
				workgetPostOfficeBox = "";
			}
			try {
				if (workgetPostOfficeBox.equalsIgnoreCase("null") || workgetPostOfficeBox.contains("meta")
						|| workgetPostOfficeBox.contains("aspose")) {
					workgetPostOfficeBox = "NA";
				}
			} catch (Exception e1) {
				workgetPostOfficeBox = "NA";
			}
			try {

				workStateOrProvince = String.valueOf(contactworkphys.getStateOrProvince());
			} catch (Exception ep) {
				workStateOrProvince = "";
			}
			try {
				if (workStateOrProvince.equalsIgnoreCase("null") || workStateOrProvince.contains("meta")
						|| workStateOrProvince.contains("aspose")) {
					workStateOrProvince = "NA";
				}
			} catch (Exception e1) {
				workStateOrProvince = "NA";
			}
			try {

				Assistant = String.valueOf(personPropSet.getAssistant());
			} catch (Exception ep) {
				Assistant = "";
			}
			try {
				if (Assistant.equalsIgnoreCase("null") || Assistant.contains("meta") || Assistant.contains("aspose")) {
					Assistant = "NA";
				}
			} catch (Exception e1) {
				Assistant = "NA";
			}
			try {

				CompanyName = String.valueOf(personPropSet.getCompanyName());
			} catch (Exception ep) {
				CompanyName = "";
			}
			try {
				if (CompanyName.equalsIgnoreCase("null") || CompanyName.contains("meta")
						|| CompanyName.contains("aspose")) {
					CompanyName = "NA";
				}
			} catch (Exception e1) {
				CompanyName = "NA";
			}
			try {

				DepartmentName = String.valueOf(personPropSet.getDepartmentName());
			} catch (Exception ep) {
				DepartmentName = "";
			}
			try {
				if (DepartmentName.equalsIgnoreCase("null") || DepartmentName.contains("meta")
						|| DepartmentName.contains("aspose")) {
					DepartmentName = "NA";
				}
			} catch (Exception e1) {
				DepartmentName = "NA";
			}

			try {

				ManagerName = String.valueOf(personPropSet.getManagerName());
			} catch (Exception ep) {
				ManagerName = "";
			}
			try {
				if (ManagerName.equalsIgnoreCase("null") || ManagerName.contains("meta")
						|| ManagerName.contains("aspose")) {
					ManagerName = "NA";
				}
			} catch (Exception e1) {
				ManagerName = "NA";
			}
			try {

				OfficeLocation = String.valueOf(personPropSet.getOfficeLocation());
			} catch (Exception ep) {
				OfficeLocation = "";
			}
			try {
				if (OfficeLocation.equalsIgnoreCase("null") || OfficeLocation.contains("meta")
						|| OfficeLocation.contains("aspose")) {
					OfficeLocation = "NA";
				}
			} catch (Exception e1) {
				OfficeLocation = "NA";
			}
			try {

				Profession = String.valueOf(personPropSet.getProfession());
			} catch (Exception ep) {
				Profession = "";
			}
			try {
				if (Profession.equalsIgnoreCase("null") || Profession.contains("meta")
						|| Profession.contains("aspose")) {
					Profession = "NA";
				}
			} catch (Exception e1) {
				Profession = "";
			}
			try {

				getTitle = String.valueOf(personPropSet.getTitle());
			} catch (Exception ep) {
				getTitle = "";
			}
			try {
				if (getTitle.equalsIgnoreCase("null") || getTitle.contains("meta") || getTitle.contains("aspose")) {
					getTitle = "NA";
				}
			} catch (Exception e1) {
				getTitle = "NA";
			}
			try {

				AssistantTelephoneNumber = String.valueOf(mapitelephone.getAssistantTelephoneNumber());
			} catch (Exception ep) {
				AssistantTelephoneNumber = "";
			}
			try {
				if (AssistantTelephoneNumber.equalsIgnoreCase("null") || homeCountry.contains("meta")
						|| AssistantTelephoneNumber.contains("meta") || AssistantTelephoneNumber.contains("aspose")) {
					AssistantTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				AssistantTelephoneNumber = "NA";
			}
			try {

				AssistantTelephoneNumber = String.valueOf(mapitelephone.getAssistantTelephoneNumber());
			} catch (Exception ep) {
				AssistantTelephoneNumber = "";
			}
			try {
				if (AssistantTelephoneNumber.equalsIgnoreCase("null") || AssistantTelephoneNumber.contains("meta")
						|| AssistantTelephoneNumber.contains("aspose")) {
					AssistantTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				AssistantTelephoneNumber = "NA";
			}
			try {

				Business2TelephoneNumber = String.valueOf(mapitelephone.getBusiness2TelephoneNumber());
			} catch (Exception ep) {
				Business2TelephoneNumber = "";
			}
			try {
				if (Business2TelephoneNumber.equalsIgnoreCase("null") || Business2TelephoneNumber.contains("meta")
						|| Business2TelephoneNumber.contains("aspose")) {
					Business2TelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				Business2TelephoneNumber = "NA";
			}
			try {

				BusinessTelephoneNumber = String.valueOf(mapitelephone.getBusinessTelephoneNumber());
			} catch (Exception ep) {
				BusinessTelephoneNumber = "";
			}
			try {
				if (BusinessTelephoneNumber.equalsIgnoreCase("null") || BusinessTelephoneNumber.contains("meta")
						|| BusinessTelephoneNumber.contains("aspose")) {
					BusinessTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				BusinessTelephoneNumber = "NA";
			}
			Phone_1_Selected = BusinessTelephoneNumber;
			try {

				CallbackTelephoneNumber = String.valueOf(mapitelephone.getCallbackTelephoneNumber());
			} catch (Exception ep) {
				CallbackTelephoneNumber = "";
			}
			try {
				if (CallbackTelephoneNumber.equalsIgnoreCase("null") || CallbackTelephoneNumber.contains("meta")
						|| CallbackTelephoneNumber.contains("aspose")) {
					CallbackTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				CallbackTelephoneNumber = "NA";
			}
			try {

				CarTelephoneNumber = String.valueOf(mapitelephone.getCarTelephoneNumber());
			} catch (Exception ep) {
				CarTelephoneNumber = "";
			}
			try {
				if (CarTelephoneNumber.equalsIgnoreCase("null") || CarTelephoneNumber.contains("meta")
						|| CarTelephoneNumber.contains("aspose")) {
					CarTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				CarTelephoneNumber = "NA";
			}
			Phone_6_Selected = CarTelephoneNumber;
			try {

				CompanyMainTelephoneNumber = String.valueOf(mapitelephone.getCompanyMainTelephoneNumber());
			} catch (Exception ep) {
				CompanyMainTelephoneNumber = "";
			}
			try {
				if (CompanyMainTelephoneNumber.equalsIgnoreCase("null") || CompanyMainTelephoneNumber.contains("meta")
						|| CompanyMainTelephoneNumber.contains("aspose")) {
					CompanyMainTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				CompanyMainTelephoneNumber = "NA";
			}
			try {

				Home2TelephoneNumber = String.valueOf(mapitelephone.getHome2TelephoneNumber());
			} catch (Exception ep) {
				Home2TelephoneNumber = "";
			}
			try {
				if (Home2TelephoneNumber.equalsIgnoreCase("null") || homeCountry.contains("meta")
						|| Home2TelephoneNumber.contains("meta") || Home2TelephoneNumber.contains("aspose")) {
					Home2TelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				Home2TelephoneNumber = "NA";
			}

			try {

				HomeTelephoneNumber = String.valueOf(mapitelephone.getHomeTelephoneNumber());
			} catch (Exception ep) {
				HomeTelephoneNumber = "";
			}
			try {
				if (HomeTelephoneNumber.equalsIgnoreCase("null") || HomeTelephoneNumber.contains("meta")
						|| HomeTelephoneNumber.contains("aspose")) {
					HomeTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				HomeTelephoneNumber = "NA";
			}
			Phone_2_Selected = HomeTelephoneNumber;
			try {

				IsdnNumber = String.valueOf(mapitelephone.getIsdnNumber());
			} catch (Exception ep) {
				IsdnNumber = "";
			}
			try {
				if (IsdnNumber.equalsIgnoreCase("null") || IsdnNumber.contains("meta")
						|| IsdnNumber.contains("aspose")) {
					IsdnNumber = "NA";
				}
			} catch (Exception e1) {
				IsdnNumber = "NA";
			}
			Phone_8_Selected = IsdnNumber;
			try {

				MobileTelephoneNumber = String.valueOf(mapitelephone.getMobileTelephoneNumber());
			} catch (Exception ep) {
				MobileTelephoneNumber = "";
			}
			try {
				if (MobileTelephoneNumber.equalsIgnoreCase("null") || MobileTelephoneNumber.contains("meta")
						|| MobileTelephoneNumber.contains("aspose")) {
					MobileTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				MobileTelephoneNumber = "NA";
			}
			Phone_4_Selected = MobileTelephoneNumber;

			try {

				OtherTelephoneNumber = String.valueOf(mapitelephone.getOtherTelephoneNumber());
			} catch (Exception ep) {
				OtherTelephoneNumber = "";
			}
			try {
				if (OtherTelephoneNumber.equalsIgnoreCase("null") || OtherTelephoneNumber.contains("meta")
						|| OtherTelephoneNumber.contains("aspose")) {
					OtherTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				OtherTelephoneNumber = "NA";
			}
			Phone_7_Selected = OtherTelephoneNumber;
			try {

				PagerTelephoneNumber = String.valueOf(mapitelephone.getPagerTelephoneNumber());
			} catch (Exception ep) {
				PagerTelephoneNumber = "";
			}
			try {
				if (PagerTelephoneNumber.equalsIgnoreCase("null") || PagerTelephoneNumber.contains("meta")
						|| PagerTelephoneNumber.contains("aspose")) {
					PagerTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				PagerTelephoneNumber = "NA";
			}
			try {

				PrimaryTelephoneNumber = String.valueOf(mapitelephone.getPrimaryTelephoneNumber());
			} catch (Exception ep) {
				PrimaryTelephoneNumber = "";
			}
			try {
				if (PrimaryTelephoneNumber.equalsIgnoreCase("null") || PrimaryTelephoneNumber.contains("meta")
						|| PrimaryTelephoneNumber.contains("aspose")) {
					PrimaryTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				PrimaryTelephoneNumber = "NA";
			}
			try {

				RadioTelephoneNumber = String.valueOf(mapitelephone.getRadioTelephoneNumber());
			} catch (Exception ep) {
				RadioTelephoneNumber = "";
			}
			try {
				if (RadioTelephoneNumber.equalsIgnoreCase("null") || homeCountry.contains("meta")
						|| RadioTelephoneNumber.contains("meta") || RadioTelephoneNumber.contains("aspose")) {
					RadioTelephoneNumber = "NA";
				}
			} catch (Exception e1) {
				RadioTelephoneNumber = "NA";
			}
			Phone_5_Selected = RadioTelephoneNumber;
			try {

				TelexNumber = String.valueOf(mapitelephone.getTelexNumber());
			} catch (Exception ep) {
				TelexNumber = "";
			}
			try {
				if (TelexNumber.equalsIgnoreCase("null") || TelexNumber.contains("meta")
						|| TelexNumber.contains("aspose")) {
					TelexNumber = "NA";
				}
			} catch (Exception e1) {
				TelexNumber = "NA";
			}

			try {

				TtyTddPhoneNumber = String.valueOf(mapitelephone.getTtyTddPhoneNumber());
			} catch (Exception ep) {
				TtyTddPhoneNumber = "";
			}
			try {
				if (TtyTddPhoneNumber.equalsIgnoreCase("null") || TtyTddPhoneNumber.contains("meta")
						|| TtyTddPhoneNumber.contains("aspose")) {
					TtyTddPhoneNumber = "NA";
				}
			} catch (Exception e1) {
				TtyTddPhoneNumber = "NA";
			}
			//
			fname = null;
			try {
				if (firstName != null & middleName != null & lastname != null) {
					fname = firstName.concat(middleName).concat(lastname);
					if (fname.contains("NA")) {
						fname = fname.replace("NA", " ");

					}
				} else if (firstName != null & middleName != null) {
					fname = firstName.concat(middleName);
				} else if (firstName != null & lastname != null) {
					fname = firstName.concat(lastname);
				} else {

					fname = firstName;
				}
			} catch (Exception e1) {
				fname = "null";
			}

			try {
				String[] data1 = { account, Address_Selected, Address_Selector, WeddingAnniversary, Assistant,
						AssistantTelephoneNumber, Attachment, Billing_Information, birthday, workAddress, workCity,
						workCountry, workgetPostOfficeBox, workPostalCode, workStateOrProvince, workgetStreet,
						bussinessfaxno, BusinessHomePage, BusinessTelephoneNumber, Business2TelephoneNumber,
						CallbackTelephoneNumber, CarTelephoneNumber, Categories, Children, workCity, CompanyName,
						CompanyMainTelephoneNumber, ComputerNetworkName, Contacts, homeCountry, fullday, CustomerId,
						DepartmentName, Email1address, Email2address, Email3address, Email1addresstype,
						Email1displayname, Email_Selected, Email_Selector, Email2addresstype, Email2displayname,
						Email3addresstype, Email3displayname, fileunder, firstName, Flag_Completed_Date, Flag_Status,
						Follow_Up_Flag, FtpSite, fname, Gender, GovernmentIdNumber, Hobbies, homeAddress, homeCity,
						homeCountry, homegetPostOfficeBox, homePostalCode, homeStateOrProvince, homegetStreet,
						homefaxno, HomeTelephoneNumber, Home2TelephoneNumber, Icon, InstantMessagingAddress,
						folder_Name, Initials, Internet_Free_Busy_Address, IsdnNumber, Job_title, Journal, Language,
						lastname, Location, Mailing_Address, Mailing_Address_Indicator, ManagerName, Message_Class,
						middleName, mileage, MobileTelephoneNumber, Modified, nickName, OfficeLocation,
						OrganizationalIdNumber, otherAddress, otherCity, otherCountry, othergetPostOfficeBox,
						otherPostalCode, otherStateOrProvince, othergetStreet, Other_Fax, OtherTelephoneNumber,
						Outlook_Data_File, PersonalHomePage, PersonalHomePage, PagerTelephoneNumber, PersonalHomePage,
						Phone_1_Selected, Phone_1_Selector, Phone_2_Selected, Phone_2_Selector, Phone_3_Selected,
						Phone_3_Selector, Phone_4_Selected, Phone_4_Selector, Phone_5_Selected, Phone_5_Selector,
						Phone_6_Selected, Phone_6_Selector, Phone_7_Selected, Phone_7_Selector, Phone_8_Selected,
						Phone_8_Selector, othergetPostOfficeBox, PrimaryTelephoneNumber, Private, Profession,
						RadioTelephoneNumber, PersonalHomePage, ReferredByName, reminder, Reminder_Time, Reminder_Topic,
						sencitivity, PersonalHomePage, Size_on_Server, SpouseName, workStateOrProvince, workgetStreet,
						Subject, Suffix, TelexNumber, title, TtyTddPhoneNumber, User_Field_1, User_Field_2,
						User_Field_3, User_Field_4, Web_Page, ZIP_Postal_Code };
				writer.writeNext(data1);

			} catch (Error e) {
				mf.logger.warning("ERROR : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

			catch (Exception e) {
				mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

		} else if (message.getMessageClass().equals("IPM.Appointment")
				|| message.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {
			MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

			String subject = null;
			String startdate = null;
			String enddate = null;
			String alldayevent = null;
			String reminder = null;
			String remindertime = null;
			String requiredattend = null;
			String categories = null;
			String location = "";
			String mileage = "";
			String getSenderEmailAddress = null;
			String getReplyTo = null;
			String getDisplayCc = null;
			String body = null;
			String getDisplayBcc = null;
			String Messageclass = null;
			try {
				subject = cal.getSubject();
			} catch (Exception e) {
				subject = "";
			}
			try {
				if (subject.equalsIgnoreCase("null") || subject.contains("meta") || subject.contains("aspose")) {
					subject = "NA";
				}
			} catch (Exception e1) {
				subject = "NA";
			}
			try {
				startdate = cal.getStartDate().toString();
			} catch (Exception e) {
				startdate = "";
			}
			try {
				if (startdate.equalsIgnoreCase("null") || startdate.contains("meta") || startdate.contains("aspose")) {
					startdate = "NA";
				}
			} catch (Exception e1) {
				startdate = "NA";
			}
			try {
				enddate = cal.getEndDate().toString();
			} catch (Exception e) {
				enddate = "";
			}
			try {
				if (enddate.equalsIgnoreCase("null") || enddate.contains("meta") || enddate.contains("aspose")) {
					enddate = "NA";
				}
			} catch (Exception e1) {
				enddate = "NA";
			}
			try {
				alldayevent = String.valueOf(cal.isAllDay());
			} catch (Exception e) {
				alldayevent = "";
			}
			try {
				if (alldayevent.equalsIgnoreCase("null") || alldayevent.contains("meta")
						|| alldayevent.contains("aspose")) {
					alldayevent = "NA";
				}
			} catch (Exception e1) {
				alldayevent = "NA";
			}
			try {
				reminder = String.valueOf(cal.getReminderSet());
			} catch (Exception e) {
				reminder = "";
			}
			try {
				if (reminder.equalsIgnoreCase("null") || reminder.contains("meta") || reminder.contains("aspose")) {
					reminder = "NA";
				}
			} catch (Exception e1) {
				reminder = "NA";
			}
			try {
				remindertime = String.valueOf(cal.getReminderDelta());
			} catch (Exception e) {
				remindertime = "";
			}
			try {
				if (remindertime.equalsIgnoreCase("null") || remindertime.contains("meta")
						|| remindertime.contains("aspose")) {
					remindertime = "NA";
				}
			} catch (Exception e1) {
				remindertime = "NA";
			}
			try {
				requiredattend = String.valueOf(cal.getAttendees());
			} catch (Exception e) {
				requiredattend = "";
			}
			try {
				if (requiredattend.equalsIgnoreCase("null") || requiredattend.contains("meta")
						|| requiredattend.contains("aspose")) {
					requiredattend = "NA";
				}
			} catch (Exception e1) {
				requiredattend = "NA";
			}
			try {
				categories = String.valueOf(cal.getCategories());
			} catch (Exception e) {
				categories = "";
			}
			try {
				if (categories.equalsIgnoreCase("null") || categories.contains("meta")
						|| categories.contains("aspose")) {
					categories = "NA";
				}
			} catch (Exception e1) {
				categories = "NA";
			}
			try {
				location = String.valueOf(cal.getLocation());
			} catch (Exception e) {
				location = "";
			}
			try {
				if (location.equalsIgnoreCase("null") || location.contains("meta") || location.contains("aspose")) {
					location = "NA";
				}
			} catch (Exception e1) {
				location = "NA";
			}
			try {
				mileage = String.valueOf(cal.getMileage());
			} catch (Exception e) {
				mileage = "";
			}
			try {
				if (mileage.equalsIgnoreCase("null") || mileage.contains("meta") || mileage.contains("aspose")) {
					mileage = "NA";
				}
			} catch (Exception e1) {
				mileage = "NA";
			}

			try {
				body = String.valueOf(cal.getBody());
			} catch (Exception e) {
				body = "";
			}
			try {
				if (body.equalsIgnoreCase("null") || body.contains("meta") || body.contains("aspose")) {
					body = "NA";
				}
			} catch (Exception e1) {
				body = "NA";
			}

			try {
				getSenderEmailAddress = message.getSenderEmailAddress();
			} catch (Exception e) {
				getSenderEmailAddress = "NA";
			}
			try {
				if (getSenderEmailAddress.equalsIgnoreCase("null") || getSenderEmailAddress.contains("meta")
						|| getSenderEmailAddress.contains("aspose")) {
					getSenderEmailAddress = "NA";
				}
			} catch (Exception e1) {
				getSenderEmailAddress = "NA";
			}

			try {
//				getReplyTo = message.getRecipients().get_Item(1).getEmailAddress();
				getReplyTo = message.getDisplayTo().toString();

			} catch (Exception e) {
				getReplyTo = "NA";
			}
			try {
				if (getReplyTo.equalsIgnoreCase("null") || getReplyTo.contains("meta")
						|| getReplyTo.contains("aspose")) {
					getReplyTo = "NA";
				}
			} catch (Exception e1) {
				getReplyTo = "NA";
			}

			try {
				getDisplayCc = message.getDisplayCc();
			} catch (Exception e) {
				getDisplayCc = "NA";
			}
			try {
				if (getDisplayCc.equalsIgnoreCase("null") || getDisplayCc.contains("meta")
						|| getDisplayCc.contains("aspose")) {
					getDisplayCc = "NA";
				}
			} catch (Exception e1) {
				getDisplayCc = "NA";
			}

			try {
				getDisplayBcc = message.getDisplayBcc();
			} catch (Exception e) {
				getDisplayBcc = "NA";
			}
			try {
				if (getDisplayBcc.equalsIgnoreCase("null") || getDisplayBcc.contains("meta")
						|| getDisplayBcc.contains("aspose")) {
					getDisplayBcc = "NA";
				}
			} catch (Exception e1) {
				getDisplayBcc = "NA";
			}

			try {
				Messageclass = message.getMessageClass();
			} catch (Exception e) {
				Messageclass = "NA";
			}
			try {
				if (Messageclass.equalsIgnoreCase("null") || Messageclass.contains("meta")
						|| Messageclass.contains("aspose")) {
					Messageclass = "NA";
				}
			} catch (Exception e1) {
				Messageclass = "NA";
			}
			if (message.getAttachments().size() > 0) {
				File fd = new File(destination_path + File.separator + path + File.separator + "Attachment"
						+ File.separator + subname);
				fd.mkdirs();
				try {
					String[] data1 = { subject, body, getSenderEmailAddress, getReplyTo, getDisplayCc, getDisplayBcc,
							startdate, enddate, alldayevent, reminder, requiredattend, remindertime, categories,
							location, fd.getAbsolutePath() };
					writer.writeNext(data1);
				} catch (Error e) {
					mf.logger.warning("ERROR : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				} catch (Exception e) {
					mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}
			} else {
				try {
					String[] data1 = { subject, body, getSenderEmailAddress, getReplyTo, getDisplayCc, getDisplayBcc,
							startdate, enddate, alldayevent, reminder, requiredattend, remindertime, categories,
							location, "No Attachment" };
					writer.writeNext(data1);
				} catch (Error e) {
					mf.logger.warning("ERROR : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				} catch (Exception e) {
					mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
							+ System.lineSeparator());
				}
			}
		} else if (message.getMessageClass().equals("IPM.Task")) {

			MapiTask task = (MapiTask) message.toMapiMessageItem();

			String subject = "";
			String startdate = "";
			String getDueDate = "";
			String getPercentComplete = "";
			String getEstimatedEffort = "";
			String getActualEffort = "";
			String getOwner = "";
			String getLastUser = "";
			String getLastDelegate = "";
			String getAttendeessizesize = "";
			String getOriginalDisplayName = "";
			String getDisplayName = "";
			String getEmailAddress = "";
			String getFaxNumber = "";
			String getAddressType = "";
			String comapanies = "";
			String Categories = "";
			String getMileage = "";
			String getBilling = "";
			String getSensitivity = "";
			String getStatus = "";
			String getHistory = "";

			MapiTaskUsers user = null;
			MapiElectronicAddress address = null;
			String[] company = null;
			String[] getCategories = null;
			try {
				user = task.getUsers();
			} catch (Exception e1) {

			}
			try {
				address = user.getAssigner();
			} catch (Exception e1) {

			}

			try {
				company = task.getCompanies();
			} catch (Exception e1) {

			}

			try {
				getCategories = task.getCategories();
			} catch (Exception e1) {

			}

			try {
				subject = task.getSubject();
			} catch (Exception e) {
				subject = "";
			}
			try {
				if (subject.equalsIgnoreCase("null") || subject.contains("meta") || subject.contains("aspose")) {
					subject = "NA";
				}
			} catch (Exception e1) {
				subject = "NA";
			}
			try {
				startdate = task.getStartDate().toString();
			} catch (Exception e) {
				startdate = "";
			}
			try {
				if (startdate.equalsIgnoreCase("null") || startdate.contains("meta") || startdate.contains("aspose")) {
					startdate = "NA";
				}
			} catch (Exception e1) {
				startdate = "NA";
			}
			try {
				getDueDate = task.getDueDate().toString();
			} catch (Exception e) {
				getDueDate = "";
			}
			try {
				if (getDueDate.equalsIgnoreCase("null") || getDueDate.contains("meta")
						|| getDueDate.contains("aspose")) {
					getDueDate = "NA";
				}
			} catch (Exception e1) {
				getDueDate = "NA";
			}
			try {
				getPercentComplete = String.valueOf(task.getPercentComplete());
			} catch (Exception e) {
				getPercentComplete = "";
			}
			try {
				if (getPercentComplete.equalsIgnoreCase("null") || getPercentComplete.contains("meta")
						|| getPercentComplete.contains("aspose")) {
					getPercentComplete = "NA";
				}
			} catch (Exception e1) {
				getPercentComplete = "NA";
			}

			try {
				getEstimatedEffort = String.valueOf(task.getEstimatedEffort());
			} catch (Exception e) {
				getEstimatedEffort = "";
			}
			try {
				if (getEstimatedEffort.equalsIgnoreCase("null") || getEstimatedEffort.contains("meta")
						|| getEstimatedEffort.contains("aspose")) {
					getEstimatedEffort = "NA";
				}
			} catch (Exception e1) {
				getEstimatedEffort = "NA";
			}
			try {
				getActualEffort = String.valueOf(task.getActualEffort());
			} catch (Exception e) {
				getActualEffort = "";
			}
			try {
				if (getActualEffort.equalsIgnoreCase("null") || getActualEffort.contains("meta")
						|| getActualEffort.contains("aspose")) {
					getActualEffort = "NA";
				}
			} catch (Exception e1) {
				getActualEffort = "";
			}
			try {
				getOwner = String.valueOf(user.getOwner());
			} catch (Exception e) {
				getOwner = "";
			}
			try {
				if (getOwner.equalsIgnoreCase("null") || getOwner.contains("meta") || getOwner.contains("aspose")) {
					getOwner = "NA";
				}
			} catch (Exception e1) {
				getOwner = "NA";
			}
			try {
				getLastUser = String.valueOf(user.getLastUser());
			} catch (Exception e) {
				getLastUser = "";
			}
			try {
				if (getLastUser.equalsIgnoreCase("null") || getLastUser.contains("meta")
						|| getLastUser.contains("aspose")) {
					getLastUser = "NA";
				}
			} catch (Exception e1) {
				getLastUser = "NA";
			}
			try {
				getLastDelegate = String.valueOf(user.getLastDelegate());
			} catch (Exception e) {
				getLastDelegate = "";
			}
			try {
				if (getLastDelegate.equalsIgnoreCase("null") || getLastDelegate.contains("meta")
						|| getLastDelegate.contains("aspose")) {
					getLastDelegate = "NA";
				}
			} catch (Exception e1) {
				getLastDelegate = "NA";
			}
			try {
				getAttendeessizesize = String.valueOf(user.getAttendees().size());
			} catch (Exception e) {
				getAttendeessizesize = "";
			}
			try {
				if (getAttendeessizesize.equalsIgnoreCase("null") || getAttendeessizesize.contains("meta")
						|| getAttendeessizesize.contains("aspose")) {
					getAttendeessizesize = "NA";
				}
			} catch (Exception e2) {
				getAttendeessizesize = "NA";
			}
			try {
				getOriginalDisplayName = String.valueOf(address.getOriginalDisplayName());

			} catch (Exception e) {
				getOriginalDisplayName = "";
			}
			try {
				if (getOriginalDisplayName.equalsIgnoreCase("null") || getOriginalDisplayName.contains("meta")
						|| getOriginalDisplayName.contains("aspose")) {
					getOriginalDisplayName = "NA";
				}
			} catch (Exception e1) {
				getOriginalDisplayName = "NA";
			}
			try {
				getDisplayName = String.valueOf(address.getDisplayName());

			} catch (Exception e) {
				getDisplayName = "";
			}
			if (getDisplayName.equalsIgnoreCase("null") || getDisplayName.contains("meta")
					|| getDisplayName.contains("aspose")) {
				getDisplayName = "NA";
			}
			try {
				getEmailAddress = String.valueOf(address.getEmailAddress());

			} catch (Exception e) {
				getEmailAddress = "";
			}
			try {
				if (getEmailAddress.equalsIgnoreCase("null") || getEmailAddress.contains("meta")
						|| getEmailAddress.contains("aspose")) {
					getEmailAddress = "NA";
				}
			} catch (Exception e1) {
				getEmailAddress = "NA";
			}
			try {
				getFaxNumber = String.valueOf(address.getFaxNumber());

			} catch (Exception e) {
				getFaxNumber = "";
			}
			try {
				if (getFaxNumber.equalsIgnoreCase("null") || getFaxNumber.contains("meta")
						|| getFaxNumber.contains("aspose")) {
					getFaxNumber = "NA";
				}
			} catch (Exception e1) {
				getFaxNumber = "NA";
			}
			try {
				getAddressType = String.valueOf(address.getAddressType());

			} catch (Exception e) {
				getAddressType = "";
			}
			try {
				if (getAddressType.equalsIgnoreCase("null") || getAddressType.contains("meta")
						|| getAddressType.contains("aspose")) {
					getAddressType = "NA";
				}
			} catch (Exception e1) {
				getAddressType = "NA";
			}
			try {
				for (int i = 0; i < company.length; i++) {
					if (i == 0) {
						comapanies = company[i];
					} else {
						comapanies = comapanies + "," + company[i];
					}
				}

			} catch (Exception e) {
				comapanies = "";
			}
			if (comapanies.equalsIgnoreCase("null") || comapanies.contains("meta") || comapanies.contains("aspose")) {
				comapanies = "NA";
			}

			try {
				for (int i = 0; i < getCategories.length; i++) {
					if (i == 0) {
						Categories = getCategories[i];
					} else {
						Categories = Categories + "," + getCategories[i];
					}
				}

			} catch (Exception e) {
				Categories = "";
			}
			try {
				if (Categories.equalsIgnoreCase("null") || Categories.contains("meta")
						|| Categories.contains("aspose")) {
					Categories = "NA";
				}
			} catch (Exception e1) {
				Categories = "NA";
			}
			try {
				getMileage = String.valueOf(task.getMileage());
			} catch (Exception e) {
				getMileage = "";
			}
			try {
				if (getMileage.equalsIgnoreCase("null") || getMileage.contains("meta")
						|| getMileage.contains("aspose")) {
					getMileage = "NA";
				}
			} catch (Exception e1) {
				getMileage = "NA";
			}
			try {
				getBilling = String.valueOf(task.getBilling());
			} catch (Exception e) {
				getBilling = "";
			}
			try {
				if (getBilling.equalsIgnoreCase("null") || getBilling.contains("meta")
						|| getBilling.contains("aspose")) {
					getBilling = "NA";
				}
			} catch (Exception e1) {
				getBilling = "NA";
			}
			try {
				int i = task.getSensitivity();
				if (i == 0) {
					getSensitivity = "None";
				} else if (i == 1) {
					getSensitivity = "Personal";

				} else if (i == 2) {
					getSensitivity = "Private";
				} else if (i == 3) {
					getSensitivity = "Company Confidential";
				}

			} catch (Exception e) {
				getSensitivity = "";
			}
			try {
				if (getSensitivity.equalsIgnoreCase("null") || getSensitivity.contains("meta")
						|| getSensitivity.contains("aspose")) {
					getSensitivity = "NA";
				}
			} catch (Exception e2) {
				getSensitivity = "NA";
			}
			try {
				int i = task.getStatus();
				if (i == 0) {
					getStatus = "Not Started";
				} else if (i == 1) {
					getStatus = "In Progress";

				} else if (i == 2) {
					getStatus = "Complete";
				} else if (i == 3) {
					getStatus = "Waiting";
				} else if (i == 4) {
					getStatus = "Deferred";
				}

			} catch (Exception e) {
				getStatus = "";
			}
			try {
				if (getStatus.equalsIgnoreCase("null") || getStatus.contains("meta") || getStatus.contains("aspose")) {
					getStatus = "NA";
				}
			} catch (Exception e1) {
				getStatus = "NA";
			}
			try {
				int i = task.getHistory();
				if (i == 0) {
					getHistory = "No Changes";
				} else if (i == 1) {
					getHistory = "Accepted";

				} else if (i == 2) {
					getHistory = "Rejected";
				} else if (i == 3) {
					getHistory = "Another Property Changed";
				} else if (i == 4) {
					getHistory = "Due Date Changed";
				} else if (i == 5) {
					getHistory = "Assigned";
				}

			} catch (Exception e) {
				getStatus = "";
			}
			try {
				if (getHistory.equalsIgnoreCase("null") || getHistory.contains("meta")
						|| getHistory.contains("aspose")) {
					getHistory = "NA";
				}
			} catch (Exception e) {
				getHistory = "NA";
			}

			String[] data1 = { subject, startdate, getDueDate, getPercentComplete, getEstimatedEffort, getActualEffort,
					getOwner, getLastUser, getLastDelegate, getAttendeessizesize, getOriginalDisplayName,
					getDisplayName, getEmailAddress, getFaxNumber, getAddressType, comapanies, Categories, getMileage,
					getBilling, getSensitivity, getStatus, getHistory };

			writer.writeNext(data1);
		} else {
			MapiConversionOptions d = MapiConversionOptions.getUnicodeFormat();
			MailConversionOptions de = new MailConversionOptions();
			MailMessage mess = message.toMailMessage(de);
			msg = MapiMessage.fromMailMessage(mess, d);
			try {

				String date = null;
				String subject = null;
				String getBody = null;
				String getSenderEmailAddress = null;
				String getReplyTo = null;
				System.out.println(message.getBody() + "  3878 :");
				try {
					date = message.getDeliveryTime().toString();
				} catch (Exception e) {
					date = "NA";
				}

				try {
					if (date.equalsIgnoreCase("null") || date.contains("meta") || date.contains("aspose")) {
						date = "NA";
					}
				} catch (Exception e1) {
					date = "NA";
				}

				try {
					subject = message.getSubject();
				} catch (Exception e) {

					subject = "NA";
				}

				try {
					if (subject.equalsIgnoreCase("null") || subject.contains("meta") || subject.contains("aspose")) {
						subject = "NA";
					}
				} catch (Exception e1) {
					subject = "NA";
				}

				try {
					getBody = message.getBody().toString();
				} catch (Exception e) {
					getBody = msg.getBody().toString();
				}
				try {
					if (getBody.equalsIgnoreCase("null") || getBody.contains("meta") || getBody.contains("aspose")) {
						getBody = "NA";
					}

				} catch (Exception e1) {
					getBody = "NA";
				}

				try {
					getSenderEmailAddress = message.getSenderEmailAddress();
				} catch (Exception e) {
					getSenderEmailAddress = "NA";
				}

				try {
					if (getSenderEmailAddress.equalsIgnoreCase("null") || getSenderEmailAddress.contains("meta")
							|| getSenderEmailAddress.contains("aspose")) {
						getSenderEmailAddress = "NA";
					}
				} catch (Exception e1) {
					getSenderEmailAddress = "NA";
				}

				try {
					getReplyTo = message.getDisplayTo().toString();
				} catch (Exception e) {
					getReplyTo = "NA";
				}

				try {
					if (getReplyTo.equalsIgnoreCase("null") || getReplyTo.contains("meta")
							|| getReplyTo.contains("aspose")) {
						getReplyTo = "NA";
					}
				} catch (Exception e1) {
					getReplyTo = "NA";
				}
				try {
					if (getReplyTo.equalsIgnoreCase("null") || getReplyTo.contains("meta")
							|| getReplyTo.contains("aspose")) {
						getReplyTo = "NA";
					}
				} catch (Exception e1) {
					getReplyTo = "NA";
				}

				String getDisplayCc = null;
				try {
					getDisplayCc = message.getDisplayCc();
				} catch (Exception e) {
					getDisplayCc = "NA";
				}

				try {
					if (getDisplayCc.equalsIgnoreCase("null") || getDisplayCc.contains("meta")
							|| getDisplayCc.contains("aspose")) {
						getDisplayCc = "NA";
					}
				} catch (Exception e1) {
					getDisplayCc = "NA";
				}

				String getDisplayBcc = null;
				try {
					getDisplayBcc = message.getDisplayBcc();
				} catch (Exception e) {
					getDisplayBcc = "NA";
				}

				try {
					if (getDisplayBcc.equalsIgnoreCase("null") || getDisplayBcc.contains("meta")
							|| getDisplayBcc.contains("aspose")) {
						getDisplayBcc = "NA";
					}
				} catch (Exception e1) {
					getDisplayBcc = "NA";
				}

				if (message.getAttachments().size() > 0) {
					File fd = new File(destination_path + File.separator + path + File.separator + "Attachment"
							+ File.separator + subname.trim());

					fd.mkdirs();

					String[] data1 = { date, subject, getBody, getSenderEmailAddress, getReplyTo, getDisplayCc,
							getDisplayBcc, fd.getAbsolutePath() };

					writer.writeNext(data1);

				} else {
					String[] data1 = { date, subject, getBody, getSenderEmailAddress, getReplyTo, getDisplayCc,
							getDisplayBcc };

					writer.writeNext(data1);
				}

			} catch (Error e) {
				e.printStackTrace();
				mf.logger.warning("ERROR : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

			catch (Exception e) {
				e.printStackTrace();
				mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

		}
		try {
			if (message.getAttachments().size() != 0) {
				f3 = new File(destination_path + File.separator + path + File.separator + "Attachment" + File.separator
						+ subname.trim());
				f3.mkdirs();
			}

			if (message.getAttachments().size() > 0) {

//				for (MapiAttachment attachment : message.getAttachments()) {
//					try {
//						attachment.save(f3.getAbsolutePath() + File.separator + attachment.getDisplayName());
//					} catch (Exception e) {
//						attachment.save(f3.getAbsolutePath() + File.separator + attachment.getLongFileName());
//					}
//				}

				for (int j = 0; j < message.getAttachments().size(); j++) {

					MapiAttachment att = message.getAttachments().get_Item(j);
					if (att.getMimeTag() != null) {
						mime_tag = att.getMimeTag();
					}
					System.out.println(mime_tag + "  mime_tag");
					String s = "";
					if (att.getDisplayName() != null || att.getLongFileName() != null) {
						try {
							s = att.getDisplayName().replaceAll("[\\[\\]]", "").trim();
						} catch (Exception e1) {
							s = att.getLongFileName().replaceAll("[\\[\\]]", "").trim();
						}
					} else {
						s = "NA";
					}
					byte[] bytes = s.getBytes(StandardCharsets.US_ASCII);
					String str = new String(bytes, StandardCharsets.US_ASCII);
					if (mime_tag != null) {
						if (mime_tag.contains("message/rfc822")) {
							att.save(f3.getAbsolutePath() + File.separator
									+ Main_Frame.getRidOfIllegalFileNameCharacters(str) + ".msg");
						} else if (mime_tag.contains("text/calendar")) {
							att.save(f3.getAbsolutePath() + File.separator
									+ Main_Frame.getRidOfIllegalFileNameCharacters(str) + ".ics");
						} else {
							att.save(f3.getAbsolutePath() + File.separator
									+ Main_Frame.getRidOfIllegalFileNameCharacters(str));
						}
					}

				}
			}
		} catch (Exception e) {
			e.printStackTrace();
			mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
					+ System.lineSeparator());
		}

	}
}
