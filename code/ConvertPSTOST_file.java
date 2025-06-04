package email.code;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.aspose.email.AppointmentSaveFormat;
import com.aspose.email.BodyContentType;
import com.aspose.email.ContactSaveFormat;
import com.aspose.email.FolderInfo;
import com.aspose.email.FolderInfoCollection;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiAttachment;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiContactNamePropertySet;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiMessage;
import com.aspose.email.MessageInfo;
import com.aspose.email.MessageInfoCollection;
import com.aspose.email.PersonalStorage;

public class ConvertPSTOST_file implements Runnable {

	File faa, f1 = null;
	boolean checkDate = false;
	static Date fromdate;
	static Date todate;
	ArrayList<Date> fromList = new ArrayList<Date>();
	ArrayList<Date> toList = new ArrayList<Date>();
	List<String> listduplicacy = new ArrayList<String>();
	List<String> listdupliccal = new ArrayList<String>();
	List<String> listdupliccontact = new ArrayList<String>();
	String from;
	String to;
	String first = null, middle = null, last = null;
	private Main_Frame mf;
	private main_multiplefile mm;
	private String filetype = "";
	private String filepath = "";
	private String destination_path = "";
	private String path = "";
	static long count_destination;
	static PersonalStorage pst;
	String Folder;
	List<String> pstfolderlist;
	static boolean datevalidflag = false;
	static long foldermessagecount = 0;

	public ConvertPSTOST_file(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList) {
		this.mf = mf;
		this.mm = mm;
		this.filetype = filetype;
		this.filepath = filepath;
		this.destination_path = destination_path;
		ConvertPSTOST_file.count_destination = count_destination;
		this.pstfolderlist = pstfolderlist;
		this.fromList = fromList;
		this.toList = toList;
	}

	public void run() {
		convertPSTOST_file(mf, filetype, destination_path, count_destination, filepath, mm, pstfolderlist, fromList,
				toList);
		System.out.println("  Count is @:" + count_destination);
		main_multiplefile.count_destination = ConvertPSTOST_file.count_destination;
	}

	void convertPSTOST_file(Main_Frame mf, String filetype, String destination_path, long count_destination,
			String filepath, main_multiplefile mm, List<String> pstfolderlist, ArrayList<Date> fromList,
			ArrayList<Date> toList) {

		ConvertPSTOST_file.count_destination = 0;
		System.out.println(" File NAME  " + filepath);
		pst = PersonalStorage.fromFile(filepath);
		String path1 = "";
		FolderInfo folderInfo2 = pst.getRootFolder();
		Folder = folderInfo2.getDisplayName();
		Folder = Folder.replace(",", "").replace(".", "");
		Folder = Main_Frame.getRidOfIllegalFileNameCharacters(Folder);
		Folder = Folder.replaceAll("[\\[\\]]", "");
		Folder = Folder.trim();
		if (Folder.equalsIgnoreCase("")) {
			Folder = "Root Folder";
		}
		String clonepath = path;
		path = path + File.separator + Folder;
		path1 = Folder;
		if (mm.chckbxMaintainFolderStructure.isSelected()) {
			new File(destination_path + File.separator + path).mkdirs();
			clonepath = path;
		}
		MessageInfoCollection messageInfoCollection1 = folderInfo2.getContents();

		int messagesize1;
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
				int countr = 0;
				MessageInfo messageInfo = (MessageInfo) messageInfoCollection1.get_Item(i);

				MapiMessage message1 = pst.extractMessage(messageInfo);
				MapiConversionOptions d = MapiConversionOptions.getUnicodeFormat();
				MailConversionOptions de = new MailConversionOptions();
				MailMessage mess1 = message1.toMailMessage(de);
				if (mm.chckbxMigrateOrBackup.isSelected()) {
					mess1.getAttachments().clear();
				}
				if (!mm.chckbxMigrateOrBackup.isSelected() && message1.getAttachments().size() > 0
						&& mm.chckbxSavePdfAttachment.isSelected()) {
					f1 = new File(destination_path + File.separator + clonepath + File.separator + "Attachment");
					f1.mkdirs();
				}
				MapiMessage message = MapiMessage.fromMailMessage(mess1, d);
				if (main_multiplefile.datefilter.isSelected()) {
					datevalidflag = mm.checkdate(message, mess1);
				}
				if (message.getMessageClass().equals("IPM.Contact")) {
					try {
						MapiContact con = (MapiContact) message.toMapiMessageItem();
//						System.out.println("3..."+mess1.getBody());
						con.setBody(mess1.getBody());
						MapiContactNamePropertySet NamePropSet = con.getNameInfo();
						if (NamePropSet.getGivenName() != null) {
							first = NamePropSet.getGivenName();
						} else {
							first = "";
						}
						if (NamePropSet.getMiddleName() != null) {
							middle = NamePropSet.getMiddleName();
						} else {
							middle = "";
						}
						if (NamePropSet.getSurname() != null) {
							last = NamePropSet.getSurname();
						} else {
							last = "";
						}
						MapiContactNamePropertySet NameProp = new MapiContactNamePropertySet();
						NameProp.setDisplayName(first + " " + middle + " " + last);
						con.setNameInfo(NameProp);
						if (mm.chckbxRemoveDuplicacy.isSelected()) {
							String input = mm.duplicacymapiContact(con);
							input = input.replaceAll("\\s", "");
							input = input.trim();
							if (!listdupliccontact.contains(input)) {
								listdupliccontact.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {
										if (!mm.chckbxMigrateOrBackup.isSelected()
												&& message1.getAttachments().size() > 0
												&& mm.chckbxSavePdfAttachment.isSelected()) {
											saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
											message1.getAttachments().clear();
										}
										con.save(
												destination_path + File.separator + clonepath + File.separator
														+ mm.namingconventionmapi(message1) + "_" + i + ".vcf",
												ContactSaveFormat.VCard);
										ConvertPSTOST_file.count_destination++;
									}
								} else {

									if (!mm.chckbxMigrateOrBackup.isSelected() && message1.getAttachments().size() > 0
											&& mm.chckbxSavePdfAttachment.isSelected()) {
										saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
										message1.getAttachments().clear();
									}
									con.save(
											destination_path + File.separator + clonepath + File.separator
													+ mm.namingconventionmapi(message1) + "_" + i + ".vcf",
											ContactSaveFormat.VCard);
									ConvertPSTOST_file.count_destination++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {

									if (!mm.chckbxMigrateOrBackup.isSelected() && message1.getAttachments().size() > 0
											&& mm.chckbxSavePdfAttachment.isSelected()) {
										saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
										message1.getAttachments().clear();
									}
									con.save(
											destination_path + File.separator + clonepath + File.separator
													+ mm.namingconventionmapi(message1) + "_" + i + ".vcf",
											ContactSaveFormat.VCard);
									ConvertPSTOST_file.count_destination++;
								}
							} else {

								if (!mm.chckbxMigrateOrBackup.isSelected() && message1.getAttachments().size() > 0
										&& mm.chckbxSavePdfAttachment.isSelected()) {
									saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
									message1.getAttachments().clear();
								}
								con.save(
										destination_path + File.separator + clonepath + File.separator
												+ mm.namingconventionmapi(message1) + "_" + i + ".vcf",
										ContactSaveFormat.VCard);
								ConvertPSTOST_file.count_destination++;
							}
						}

						countr++;
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning(
								"Exception : " + e.getMessage() + "Contact" + " " + countr + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else if ((message.getMessageClass().equals("IPM.Appointment")
						|| message.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR"))) {

					try {
						MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();
						if (mm.chckbxRemoveDuplicacy.isSelected()) {
							String input = mm.duplicacymapiCal(cal);
							input = input.replaceAll("\\s", "");
							input = input.trim();

							if (!listdupliccal.contains(input)) {
								listdupliccal.add(input);
								if (main_multiplefile.datefilter.isSelected()) {
									if (datevalidflag) {

										if (!mm.chckbxMigrateOrBackup.isSelected()
												&& message1.getAttachments().size() > 0
												&& mm.chckbxSavePdfAttachment.isSelected()) {
											saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
											message1.getAttachments().clear();
										}
										cal.save(
												destination_path + File.separator + clonepath + File.separator
														+ mm.namingconventionmapi(message) + "_" + i + ".ics",
												AppointmentSaveFormat.Ics);
										ConvertPSTOST_file.count_destination++;
									}
								} else {

									if (!mm.chckbxMigrateOrBackup.isSelected() && message1.getAttachments().size() > 0
											&& mm.chckbxSavePdfAttachment.isSelected()) {
										saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
										message1.getAttachments().clear();
									}
									cal.save(
											destination_path + File.separator + clonepath + File.separator
													+ mm.namingconventionmapi(message) + "_" + i + ".ics",
											AppointmentSaveFormat.Ics);
									ConvertPSTOST_file.count_destination++;
								}
							}
						} else {
							if (main_multiplefile.datefilter.isSelected()) {
								if (datevalidflag) {

									if (!mm.chckbxMigrateOrBackup.isSelected() && message1.getAttachments().size() > 0
											&& mm.chckbxSavePdfAttachment.isSelected()) {
										saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
										message1.getAttachments().clear();
									}
									cal.save(
											destination_path + File.separator + clonepath + File.separator
													+ mm.namingconventionmapi(message) + "_" + i + ".ics",
											AppointmentSaveFormat.Ics);
									ConvertPSTOST_file.count_destination++;
								}
							} else {

								if (!mm.chckbxMigrateOrBackup.isSelected() && message1.getAttachments().size() > 0
										&& mm.chckbxSavePdfAttachment.isSelected()) {
									saveContactAttachment(mess1, message1, destination_path, clonepath, f1);
									message1.getAttachments().clear();
								}
								cal.save(
										destination_path + File.separator + clonepath + File.separator
												+ mm.namingconventionmapi(message) + "_" + i + ".ics",
										AppointmentSaveFormat.Ics);
								ConvertPSTOST_file.count_destination++;
							}

						}
						countr++;
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " " + i
								+ mf.namingconventionmapi(message) + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else if (message.getMessageClass().equals("IPM.Task")) {
					try {

						mapifile mf1 = new mapifile(mm, mf, filetype, destination_path, clonepath,
								ConvertPSTOST_file.count_destination, message, filepath, ConvertPSTOST_file.this,
								mess1);
						Thread saveTh = new Thread(mf1);
						saveTh.start();
						saveTh.join();

					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning(
								"Exception : " + e.getMessage() + "Task" + " " + countr + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				} else {
					try {

						mapifile mf1 = new mapifile(mm, mf, filetype, destination_path, clonepath,
								ConvertPSTOST_file.count_destination, message, filepath, ConvertPSTOST_file.this,
								mess1);
						Thread saveTh = new Thread(mf1);
						saveTh.start();
						saveTh.join();
					} catch (OutOfMemoryError ep) {
						mf.logger.info(
								"Out of memory error:" + ep.getMessage() + "  " + mf.namingconventionmapi(message));
					} catch (Exception e) {
						mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + i
								+ mf.namingconventionmapi(message) + System.lineSeparator());
						e.printStackTrace();
						continue;
					}

				}

				mm.lbl_progressreport.setText("Total Message Saved : " + ConvertPSTOST_file.count_destination + "    "
						+ Folder + " Extracting messsage " + message.getSubject());

			} catch (Exception e) {
				e.printStackTrace();
				continue;
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
				Folder = Main_Frame.getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();
				path = path + File.separator + Folder;
				String path3 = path1 + File.separator + Folder;

				for (int l = 0; l < pstfolderlist.size(); l++) {

					if (mm.stop) {
						break;
					}
					if (pstfolderlist.get(l).trim().equalsIgnoreCase(path3)) {
						mm.lbl_progressreport.setText(" Getting Folder " + Folder);
						int countr = 1;

						if (mm.chckbxMaintainFolderStructure.isSelected()) {
							new File(destination_path + File.separator + path).mkdirs();
							clonepath = path;
						}

						MessageInfoCollection messageInfoCollection = folderInfo.getContents();
						listduplicacy.clear();
						listdupliccal.clear();
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
								if (main_multiplefile.datefilter.isSelected()) {
									datevalidflag = mm.checkdate(message1, mess);
								}
								if (!mm.chckbxMigrateOrBackup.isSelected() && message1.getAttachments().size() > 0
										&& mm.chckbxSavePdfAttachment.isSelected()) {
									f1 = new File(destination_path + File.separator + clonepath + File.separator
											+ "Attachment");
									f1.mkdirs();
								}
								MapiMessage message = MapiMessage.fromMailMessage(mess, d);
								if (message.getMessageClass().equals("IPM.Contact")) {
									try {
										MapiContact con = (MapiContact) message.toMapiMessageItem();
//										System.out.println("2..."+mess.getBody());
										
										con.setBodyContent(mess.getBody(), mess.getBodyType());
//										con
//										System.out.println(con.getBody());
										MapiContactNamePropertySet NamePropSet = con.getNameInfo();

										if (NamePropSet.getGivenName() != null) {
											first = NamePropSet.getGivenName();
										} else {
											first = "";
										}
										if (NamePropSet.getMiddleName() != null) {
											middle = NamePropSet.getMiddleName();
										} else {
											middle = "";
										}
										if (NamePropSet.getSurname() != null) {
											last = NamePropSet.getSurname();
										} else {
											last = "";
										}
										MapiContactNamePropertySet NameProp = new MapiContactNamePropertySet();
										NameProp.setDisplayName(first + " " + middle + " " + last);
										con.setNameInfo(NameProp);

										if (mm.chckbxRemoveDuplicacy.isSelected()) {

											String input = mm.duplicacymapiContact(con);
											input = input.replaceAll("\\s", "");
											input = input.trim();

											if (!listdupliccontact.contains(input)) {
												listdupliccontact.add(input);

												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {

														if (!mm.chckbxMigrateOrBackup.isSelected()
																&& message1.getAttachments().size() > 0
																&& mm.chckbxSavePdfAttachment.isSelected()) {
															saveContactAttachment(mess, message1, destination_path,
																	clonepath, f1);
															message1.getAttachments().clear();
														}
														con.save(destination_path + File.separator + clonepath
																+ File.separator + mm.namingconventionmapi(message1)
																+ "_" + i + ".vcf", ContactSaveFormat.VCard);
														ConvertPSTOST_file.count_destination++;
													}
												} else {

													if (!mm.chckbxMigrateOrBackup.isSelected()
															&& message1.getAttachments().size() > 0
															&& mm.chckbxSavePdfAttachment.isSelected()) {
														saveContactAttachment(mess, message1, destination_path,
																clonepath, f1);
														message1.getAttachments().clear();
													}
													con.save(destination_path + File.separator + clonepath
															+ File.separator + mm.namingconventionmapi(message1) + "_"
															+ i + ".vcf", ContactSaveFormat.VCard);
													ConvertPSTOST_file.count_destination++;
												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {

													if (!mm.chckbxMigrateOrBackup.isSelected()
															&& message1.getAttachments().size() > 0
															&& mm.chckbxSavePdfAttachment.isSelected()) {
														saveContactAttachment(mess, message1, destination_path,
																clonepath, f1);
														message1.getAttachments().clear();
													}
													con.save(destination_path + File.separator + clonepath
															+ File.separator + mm.namingconventionmapi(message1) + "_"
															+ i + ".vcf", ContactSaveFormat.VCard);
													ConvertPSTOST_file.count_destination++;
												}
											} else {

												if (!mm.chckbxMigrateOrBackup.isSelected()
														&& message1.getAttachments().size() > 0
														&& mm.chckbxSavePdfAttachment.isSelected()) {
													saveContactAttachment(mess, message1, destination_path, clonepath,
															f1);
													message1.getAttachments().clear();
												}
												con.save(
														destination_path + File.separator + clonepath + File.separator
																+ mm.namingconventionmapi(message1) + "_" + i + ".vcf",
														ContactSaveFormat.VCard);
												ConvertPSTOST_file.count_destination++;
											}
										}
										countr++;
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message));
									} catch (Exception e) {
										mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " " + countr
												+ System.lineSeparator());
										e.printStackTrace();
										continue;
									}

								} else if ((message.getMessageClass().equals("IPM.Appointment")
										|| message.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR"))) {

									try {

										MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();
										if (mm.chckbxRemoveDuplicacy.isSelected()) {

											String input = mm.duplicacymapiCal(cal);
											input = input.replaceAll("\\s", "");
											input = input.trim();

											if (!listdupliccal.contains(input)) {
												listdupliccal.add(input);
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {

														if (!mm.chckbxMigrateOrBackup.isSelected()
																&& message1.getAttachments().size() > 0
																&& mm.chckbxSavePdfAttachment.isSelected()) {
															saveContactAttachment(mess, message1, destination_path,
																	clonepath, f1);
															message1.getAttachments().clear();
														}
														cal.save(destination_path + File.separator + clonepath
																+ File.separator
																+ mm.namingconventionmapi(message).trim() + "_" + i
																+ ".ics", AppointmentSaveFormat.Ics);
														ConvertPSTOST_file.count_destination++;
													}
												} else {

													if (!mm.chckbxMigrateOrBackup.isSelected()
															&& message1.getAttachments().size() > 0
															&& mm.chckbxSavePdfAttachment.isSelected()) {
														saveContactAttachment(mess, message1, destination_path,
																clonepath, f1);
														message1.getAttachments().clear();
													}
													cal.save(destination_path + File.separator + clonepath
															+ File.separator + mm.namingconventionmapi(message).trim()
															+ "_" + i + ".ics", AppointmentSaveFormat.Ics);
													ConvertPSTOST_file.count_destination++;

												}
											}
										} else {
											if (main_multiplefile.datefilter.isSelected()) {
												if (datevalidflag) {

													if (!mm.chckbxMigrateOrBackup.isSelected()
															&& message1.getAttachments().size() > 0
															&& mm.chckbxSavePdfAttachment.isSelected()) {
														saveContactAttachment(mess, message1, destination_path,
																clonepath, f1);
														message1.getAttachments().clear();
													}
													cal.save(destination_path + File.separator + clonepath
															+ File.separator + mm.namingconventionmapi(message).trim()
															+ "_" + i + ".ics", AppointmentSaveFormat.Ics);
													ConvertPSTOST_file.count_destination++;
												}
											} else {

												if (!mm.chckbxMigrateOrBackup.isSelected()
														&& message1.getAttachments().size() > 0
														&& mm.chckbxSavePdfAttachment.isSelected()) {
													saveContactAttachment(mess, message1, destination_path, clonepath,
															f1);
													message1.getAttachments().clear();
												}
												cal.save(destination_path + File.separator + clonepath + File.separator
														+ mm.namingconventionmapi(message).trim() + "_" + i + ".ics",
														AppointmentSaveFormat.Ics);
												ConvertPSTOST_file.count_destination++;

											}

										}
										countr++;
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message));
									} catch (Exception e) {
										mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " " + i
												+ mf.namingconventionmapi(message) + System.lineSeparator());
										e.printStackTrace();
										continue;
									}

								} else if (message.getMessageClass().equals("IPM.Task")) {
									try {

										mapifile mf1 = new mapifile(mm, mf, filetype, destination_path, clonepath,
												ConvertPSTOST_file.count_destination, message, filepath,
												ConvertPSTOST_file.this, mess);
										Thread saveTh = new Thread(mf1);
										saveTh.start();
										saveTh.join();
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message));
									} catch (Exception e) {
										mf.logger.warning("Exception : " + e.getMessage() + "Task" + " " + countr
												+ System.lineSeparator());
										e.printStackTrace();
										continue;
									}

								} else {
									try {

										mapifile mf1 = new mapifile(mm, mf, filetype, destination_path, clonepath,
												ConvertPSTOST_file.count_destination, message, filepath,
												ConvertPSTOST_file.this, mess);
										Thread saveTh = new Thread(mf1);
										saveTh.start();
										saveTh.join();
									} catch (OutOfMemoryError ep) {
										mf.logger.info("Out of memory error:" + ep.getMessage() + "  "
												+ mf.namingconventionmapi(message));
									} catch (Exception e) {
										mf.logger.warning("Exception : " + e.getMessage() + "Message" + " " + i
												+ mf.namingconventionmapi(message) + System.lineSeparator());
										e.printStackTrace();
										continue;
									}

								}

								mm.lbl_progressreport
										.setText("Total Message Saved : " + ConvertPSTOST_file.count_destination
												+ "    " + Folder + " Extracting messsage " + message.getSubject());
							} catch (Exception e) {
								continue;
							}
						}
						System.out.println("  Total Message Saved Count  " + ConvertPSTOST_file.count_destination + "  "
								+ Folder + " 693 " + messagesize);
					}
				}
				if (folderInfo.hasSubFolders()) {
					getsubfolderforPSTOSTfile(folderInfo, path3, clonepath);
				}

				path = mf.removefolder(path);
			} catch (Exception e) {
				continue;
			}

		}

	}

	void getsubfolderforPSTOSTfile(FolderInfo f, String path3, String clonepath) {

		FolderInfoCollection subfolder = f.getSubFolders();
		for (int k = 0; k < subfolder.size(); k++) {
			try {
				if (mm.stop) {
					break;
				}
				FolderInfo folderf = subfolder.get_Item(k);

				String Folder = folderf.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = Main_Frame.getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();

				path = path + File.separator + Folder;
				path3 = path3 + File.separator + Folder;

				for (int l = 0; l < pstfolderlist.size(); l++) {
					if (mm.stop) {
						break;
					}

					if (pstfolderlist.get(l).equalsIgnoreCase(path3)) {
						mm.lbl_progressreport.setText("Getting : " + Folder);
						if (mm.chckbxMaintainFolderStructure.isSelected()) {
							new File(destination_path + File.separator + path).mkdirs();
							clonepath = path;
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
										datevalidflag = mm.checkdate(message, mess);
									}
									if (!mm.chckbxMigrateOrBackup.isSelected() && message1.getAttachments().size() > 0
											&& mm.chckbxSavePdfAttachment.isSelected()) {
										f1 = new File(destination_path + File.separator + clonepath + File.separator
												+ "Attachment");
										f1.mkdirs();
									}
									if (message.getMessageClass().equals("IPM.Contact")) {
										try {
											MapiContact con = (MapiContact) message.toMapiMessageItem();
										System.out.println("1..."+mess.getBody());
											con.setBody(mess.getBody());
										
											MapiContactNamePropertySet NamePropSet = con.getNameInfo();

											if (NamePropSet.getGivenName() != null) {
												first = NamePropSet.getGivenName();
											} else {
												first = "";
											}
											if (NamePropSet.getMiddleName() != null) {
												middle = NamePropSet.getMiddleName();
											} else {
												middle = "";
											}
											if (NamePropSet.getSurname() != null) {
												last = NamePropSet.getSurname();
											} else {
												last = "";
											}
											MapiContactNamePropertySet NameProp = new MapiContactNamePropertySet();
											NameProp.setDisplayName(first + " " + middle + " " + last);
											con.setNameInfo(NameProp);
											if (mm.chckbxRemoveDuplicacy.isSelected()) {
												String input = mm.duplicacymapiContact(con);
												input = input.replaceAll("\\s", "");
												input = input.trim();
												if (!listdupliccontact.contains(input)) {
													listdupliccontact.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {

															if (!mm.chckbxMigrateOrBackup.isSelected()
																	&& message1.getAttachments().size() > 0
																	&& mm.chckbxSavePdfAttachment.isSelected()) {
																saveContactAttachment(mess, message1, destination_path,
																		clonepath, f1);
																con.getAttachments().clear();

															}
															con.save(destination_path + File.separator + clonepath
																	+ File.separator + mm.namingconventionmapi(message)
																	+ "_" + i + ".vcf", ContactSaveFormat.VCard);
															ConvertPSTOST_file.count_destination++;
														}
													} else {

														if (!mm.chckbxMigrateOrBackup.isSelected()
																&& message1.getAttachments().size() > 0
																&& mm.chckbxSavePdfAttachment.isSelected()) {
															saveContactAttachment(mess, message1, destination_path,
																	clonepath, f1);
															con.getAttachments().clear();
														}
														con.save(destination_path + File.separator + clonepath
																+ File.separator + mm.namingconventionmapi(message)
																+ "_" + i + ".vcf", ContactSaveFormat.VCard);
														ConvertPSTOST_file.count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {

														if (!mm.chckbxMigrateOrBackup.isSelected()
																&& message1.getAttachments().size() > 0
																&& mm.chckbxSavePdfAttachment.isSelected()) {
															saveContactAttachment(mess, message1, destination_path,
																	clonepath, f1);
															con.getAttachments().clear();
														}
														con.save(destination_path + File.separator + clonepath
																+ File.separator + mm.namingconventionmapi(message)
																+ "_" + i + ".vcf", ContactSaveFormat.VCard);
														ConvertPSTOST_file.count_destination++;
													}
												} else {

													if (!mm.chckbxMigrateOrBackup.isSelected()
															&& message1.getAttachments().size() > 0
															&& mm.chckbxSavePdfAttachment.isSelected()) {
														saveContactAttachment(mess, message1, destination_path,
																clonepath, f1);
														con.getAttachments().clear();
													}
													con.save(destination_path + File.separator + clonepath
															+ File.separator + mm.namingconventionmapi(message) + "_"
															+ i + ".vcf", ContactSaveFormat.VCard);
													ConvertPSTOST_file.count_destination++;
												}
											}
										} catch (Exception e) {
											e.printStackTrace();
											mf.logger.warning("Exception : " + e.getMessage() + "Contact" + " "
													+ System.lineSeparator());
											continue;
										}

									} else if (message.getMessageClass().equals("IPM.Appointment")
											|| message.getMessageClass().contains("IPM.Schedule.Meeting.Request")&&!message.getMessageClass().contains("IPM.Schedule.Meeting.Request.NDR")) {

										try {

											MapiCalendar cal = null;
											try {
												cal = (MapiCalendar) message.toMapiMessageItem();
											} catch (Exception e) {
											}
											if (mm.chckbxRemoveDuplicacy.isSelected()) {

												String input = mm.duplicacymapiCal(cal);
												input = input.replaceAll("\\s", "");
												input = input.trim();

												if (!listdupliccal.contains(input)) {
													listdupliccal.add(input);
													if (main_multiplefile.datefilter.isSelected()) {
														if (datevalidflag) {

															if (!mm.chckbxMigrateOrBackup.isSelected()
																	&& message1.getAttachments().size() > 0
																	&& mm.chckbxSavePdfAttachment.isSelected()) {
																saveContactAttachment(mess, message1, destination_path,
																		clonepath, f1);
																cal.getAttachments().clear();
															}
															cal.save(destination_path + File.separator + clonepath
																	+ File.separator + mm.namingconventionmapi(message)
																	+ "_" + i + ".ics", AppointmentSaveFormat.Ics);
															ConvertPSTOST_file.count_destination++;
														}
													} else {

														if (!mm.chckbxMigrateOrBackup.isSelected()
																&& message1.getAttachments().size() > 0
																&& mm.chckbxSavePdfAttachment.isSelected()) {
															saveContactAttachment(mess, message1, destination_path,
																	clonepath, f1);
															cal.getAttachments().clear();
														}
														cal.save(destination_path + File.separator + clonepath
																+ File.separator + mm.namingconventionmapi(message)
																+ "_" + i + ".ics", AppointmentSaveFormat.Ics);
														ConvertPSTOST_file.count_destination++;
													}
												}
											} else {
												if (main_multiplefile.datefilter.isSelected()) {
													if (datevalidflag) {

														if (!mm.chckbxMigrateOrBackup.isSelected()
																&& message.getAttachments().size() > 0
																&& mm.chckbxSavePdfAttachment.isSelected()) {
															saveContactAttachment(mess, message1, destination_path,
																	clonepath, f1);
															cal.getAttachments().clear();
														}
														cal.save(destination_path + File.separator + clonepath
																+ File.separator + mm.namingconventionmapi(message)
																+ "_" + i + ".ics", AppointmentSaveFormat.Ics);
														ConvertPSTOST_file.count_destination++;
													}
												} else {

													if (!mm.chckbxMigrateOrBackup.isSelected()
															&& message1.getAttachments().size() > 0
															&& mm.chckbxSavePdfAttachment.isSelected()) {
														saveContactAttachment(mess, message1, destination_path,
																clonepath, f1);
														cal.getAttachments().clear();
													}
													cal.save(destination_path + File.separator + clonepath
															+ File.separator + mm.namingconventionmapi(message) + "_"
															+ i + ".ics", AppointmentSaveFormat.Ics);
													ConvertPSTOST_file.count_destination++;
												}

											}

										} catch (Exception e) {
											e.printStackTrace();
											mf.logger.warning("Exception : " + e.getMessage() + "Calendar" + " "
													+ System.lineSeparator());
											continue;
										}

									} else if (message.getMessageClass().equals("IPM.Task")) {
										try {

											mapifile mf1 = new mapifile(mm, mf, filetype, destination_path, clonepath,
													ConvertPSTOST_file.count_destination, message, filepath,
													ConvertPSTOST_file.this, mess);
											Thread saveTh = new Thread(mf1);
											saveTh.start();
											saveTh.join();
										} catch (Exception e) {
											mf.logger.warning("Exception : " + e.getMessage() + "Message" + " "
													+ System.lineSeparator());
											continue;
										}

									} else {
										try {

											mapifile mf1 = new mapifile(mm, mf, filetype, destination_path, clonepath,
													ConvertPSTOST_file.count_destination, message, filepath,
													ConvertPSTOST_file.this, mess);
											Thread saveTh = new Thread(mf1);
											saveTh.start();
											saveTh.join();
										} catch (Exception e) {
											mf.logger.warning("Exception : " + e.getMessage() + "Message" + " "
													+ System.lineSeparator());
											continue;
										}
									}
									mm.lbl_progressreport.setText(
											"  Total Message Saved Count  " + ConvertPSTOST_file.count_destination
													+ "  " + Folder + "   Extracting messsage " + message.getSubject());
								} catch (Exception e) {
									continue;
								}
							}
							System.out.println("  Total Message Saved Count  " + ConvertPSTOST_file.count_destination
									+ "  " + Folder + " 1022 " + messagesize);
						}
					}
				}
				if (folderf.hasSubFolders()) {
					getsubfolderforPSTOSTfile(folderf, path3, clonepath);
				}

				path = mf.removefolder(path);
				path3 = mf.removefolder(path3);

			} catch (Exception e) {
				continue;
			}
		}

	}

	void saveContactAttachment(MailMessage mess, MapiMessage message, String destination_path, String clonepath,
			File f1) {
		File f2 = new File(f1.getAbsolutePath() + File.separator + mm.namingconventionmapi(message).trim());
		if (f2.exists()) {
			f2 = new File(f1.getAbsolutePath() + File.separator + mm.namingconventionmapi(message).trim() + "_"
					+ count_destination);
		}
		f2.mkdir();
		for (MapiAttachment attachment : message.getAttachments()) {
			try {
				attachment.save(f2.getAbsolutePath() + File.separator + attachment.getDisplayName().trim());
			} catch (Exception e) {
				attachment.save(f2.getAbsolutePath() + File.separator + attachment.getLongFileName().trim());
			}
		}
	}
}
