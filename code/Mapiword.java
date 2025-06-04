package email.code;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.aspose.cells.Workbook;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiAttachment;
import com.aspose.email.MapiAttachmentCollection;
import com.aspose.email.MapiConversionOptions;
import com.aspose.email.MapiMessage;
import com.aspose.email.MapiTask;
import com.aspose.email.SaveOptions;
import com.aspose.pdf.Image;
import com.aspose.pdf.facades.PdfContentEditor;
import com.aspose.slides.Presentation;
import com.aspose.words.Document;
import com.aspose.words.ImageSaveOptions;
import com.aspose.words.LoadFormat;
import com.aspose.words.LoadOptions;
import com.aspose.words.PageSet;
import com.aspose.words.SaveFormat;

public class Mapiword implements Runnable {
	ArrayList<String> all;
	File m;
	File f;
	int number = 1;
	private List<String> filelist = new ArrayList<String>();
	private Main_Frame mf;
	private main_multiplefile mm;
	private String filetype = "";
	private String destination_path = "";
	private String path = "";
	ConvertPSTOST_word cf;
	long count_destination;
	private MapiMessage message = null;
	private MailMessage mess = null;
	private Date reciveddate = null;
	long k;
	private String filepath = "";
	static boolean datevalidflag = false;
	List<String> listduplicacy = new ArrayList<String>();
	List<String> listduplictask = new ArrayList<String>();
	File f3 = null;
	String mime_tag = "";

	public Mapiword(Main_Frame mf, String filetype, String destination_path, String path, long count_destination,
			MapiMessage message, String filepath, ConvertPSTOST_word cf, main_multiplefile mm,
			List<String> listduplicacy, MailMessage mess, List<String> listduplictask) {
		this.mf = mf;
		this.mm = mm;
		this.filetype = filetype;
		this.listduplicacy = listduplicacy;
		this.listduplictask = listduplictask;
		this.cf = cf;
		this.destination_path = destination_path;
		this.path = path;
		this.filepath = filepath;
		this.count_destination = count_destination;
		this.message = message;
		this.mess = mess;
		mf.stop = false;
	}

	public void run() {
		if (message.getMessageClass().equals("IPM.Task")) {
			mapidupword_Task(message, reciveddate, cf, mess);
		} else {
			mapidupword(message, reciveddate, cf, mess);
		}

		cf.listduplicacy = listduplicacy;
		cf.listduplictask = listduplictask;
		cf.count_destination = count_destination;
	}

	void mapidupword(MapiMessage message, Date reciveddate, ConvertPSTOST_word cf, MailMessage mess) {
		if (main_multiplefile.datefilter.isSelected()) {
			datevalidflag = mm.checkdate(message, mess);
		}
		if (mm.chckbxRemoveDuplicacy.isSelected()) {

			String input = mm.duplicacymapi(message);

			if (!listduplicacy.contains(input)) {
				listduplicacy.add(input);
				if (main_multiplefile.datefilter.isSelected()) {
					if (ConvertPSTOST_word.datevalidflag) {
						Mapimessage_word(message);
						count_destination++;
					}
				} else {
					Mapimessage_word(message);
					count_destination++;
				}
			}
		} else {
			if (main_multiplefile.datefilter.isSelected()) {
				if (ConvertPSTOST_word.datevalidflag) {
					Mapimessage_word(message);
					count_destination++;
				}
			} else {
				Mapimessage_word(message);
				count_destination++;
			}
		}

	}

	// Task :
	void mapidupword_Task(MapiMessage message, Date reciveddate, ConvertPSTOST_word cf, MailMessage mess) {
		MapiTask task = null;
		if (message.getMessageClass().equals("IPM.Task")) {
			task = (MapiTask) message.toMapiMessageItem();
		}
		if (main_multiplefile.datefilter.isSelected()) {
			datevalidflag = mm.checkdate(message, mess);
		}
		if (mm.chckbxRemoveDuplicacy.isSelected()) {

			String input = mm.duplicacymapiTask(task);
			if (!listduplictask.contains(input)) {
				listduplictask.add(input);
				if (main_multiplefile.datefilter.isSelected()) {
					if (ConvertPSTOST_word.datevalidflag) {
						Mapimessage_word(message);
						count_destination++;
					}
				} else {
					Mapimessage_word(message);
					count_destination++;
				}
			}
		} else {
			if (main_multiplefile.datefilter.isSelected()) {
				if (ConvertPSTOST_word.datevalidflag) {
					Mapimessage_word(message);
					count_destination++;
				}
			} else {
				Mapimessage_word(message);
				count_destination++;
			}
		}

	}

	@SuppressWarnings({})
	public void Mapimessage_word(MapiMessage message) {
		String path5 = "";
		String naming_convention = "";
		if (mm.fileoption.equals("EML File (.eml)") || mm.fileoption.equals("EMLX File (.emlx)")
				|| mm.fileoption.equals("Message File (.msg)") || mm.fileoption.equals("Maildir")) {
			naming_convention = mm.namingconventionmapi(message, new File(filepath));
			path5 = destination_path + File.separator + path + File.separator + naming_convention + "_"
					+ Main_Frame.count_destination;

		} else {
			naming_convention = mm.namingconventionmapi(message).trim();
			path5 = destination_path + File.separator + path + File.separator + naming_convention + "_"
					+ Main_Frame.count_destination;
		}

		ByteArrayOutputStream emlStream = new ByteArrayOutputStream();

		message.save(emlStream, SaveOptions.getDefaultMhtml());
		LoadOptions lo = new LoadOptions();
		lo.setLoadFormat(LoadFormat.MHTML);
		MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
		MapiMessage message1 = MapiMessage.fromMailMessage(mess, d);
		try {
			File fu = null, f2 = null;
			if (mm.chckbxSavePdfAttachment.isSelected()) {
				fu = new File(destination_path + File.separator + path + File.separator + "Attachment");
				fu.mkdirs();
			}
			Document doc = new Document(new ByteArrayInputStream(emlStream.toByteArray()), lo);
			if (filetype.equalsIgnoreCase("PDF")) {

				path5 = path5.replaceAll("\\p{C}", "-") + ".pdf";
				doc.save(path5, SaveFormat.PDF);

				if (message.getAttachments().size() > 0) {
					PdfContentEditor editor = new PdfContentEditor();
					editor.bindPdf(path5);
					if (filelist.contains(naming_convention)) {
						naming_convention = naming_convention + "_" + number;
					} else {
						filelist.add(naming_convention);
					}

					if (message.getAttachments().size() > 0) {
//						if (mm.chckbx_convert_pdf_to_pdf.isSelected()) {
//							f2 = new File(fu.getAbsolutePath() + File.separator + naming_convention.trim());
//							if (f2.exists()) {
//								f2 = new File(fu.getAbsolutePath() + File.separator + naming_convention.trim() + "_"
//										+ count_destination);
//							}
//							f2.mkdir();
//						}
						for (int j = 0; j < message.getAttachments().size(); j++) {
							MapiAttachment attachment = message.getAttachments().get_Item(j);
							if (attachment.getMimeTag() != null) {
								mime_tag = attachment.getMimeTag();
							} else {
								MapiAttachment att1 = message1.getAttachments().get_Item(j);
								mime_tag = att1.getMimeTag();
							}
							if (mm.chckbxSavePdfAttachment.isSelected()) {
								save_Attachment_Word(message, naming_convention, fu);

							} else if (mm.chckbx_convert_pdf_to_pdf.isSelected()) {

								File f21 = new File(destination_path + File.separator + "Attachment" + File.separator
										+ naming_convention);

								f21.mkdirs();
								String str1 = null;
								String str = null;

//								if (!(attachment.getExtension() == null)) {

								File f1 = new File(System.getenv("APPDATA") + File.separator + "Test" + File.separator
										+ File.separator + path.trim() + File.separator + "Attachment" + File.separator
										+ naming_convention);
								f1.mkdirs();

								String s11;
								try {
									s11 = attachment.getLongFileName().replaceAll("[\\[\\]]", "");
								} catch (Exception e1) {
									s11 = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
								}

								byte[] bytes1 = s11.getBytes(StandardCharsets.US_ASCII);
								str1 = new String(bytes1, StandardCharsets.US_ASCII);
								f1.isAbsolute();
								f1.exists();

								attachment.save(f1.getAbsolutePath().trim() + File.separator
										+ Main_Frame.getRidOfIllegalFileNameCharacters(str1));
								f1.getPath();
								String str4 = f1.getPath().trim() + File.separator + str1;
								String s = attachment.getDisplayName().replaceAll("[\\[\\]]", "");
								byte[] bytes11 = s.getBytes(StandardCharsets.UTF_16LE);
								str = new String(bytes11, StandardCharsets.UTF_16LE);

								try {

									if (str.endsWith("txt")) {

										Document doc1 = new Document(str4);
										doc1.save(f21.getAbsolutePath().trim() + File.separator + str.replace("txt", "")
												+ "pdf", SaveFormat.PDF);
										editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("txt", "") + "pdf", "");
										m = new File(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("txt", "") + "pdf");
										m.delete();
									} else if (str.endsWith("docx")) {

										Document doc1 = new Document(str4);
										doc1.save(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("docx", "") + "pdf", SaveFormat.PDF);
										editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("docx", "") + "pdf", "");
										m = new File(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("docx", "") + "pdf");
										m.delete();
									} else if (str.endsWith("docm")) {
										Document doc1 = new Document(str4);
										doc1.save(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("docm", "") + "pdf", SaveFormat.PDF);
										editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("docm", "") + "pdf", "");
										m = new File(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("docm", "") + "pdf");
										m.delete();
									} else if (str.endsWith("doc")) {
										Document doc1 = new Document(str4);
										doc1.save(f21.getAbsolutePath().trim() + File.separator + str.replace("doc", "")
												+ "pdf", SaveFormat.PDF);
										editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("doc", "") + "pdf", "");
										m = new File(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("doc", "") + "pdf");
										m.delete();
									} else if (str.endsWith("png")) {
										java.nio.file.Path dataDir = Paths.get(str4);
										com.aspose.pdf.Document d1 = new com.aspose.pdf.Document();
										com.aspose.pdf.Page p = d1.getPages().add();
										Image img = new Image();
										img.setFile(Paths.get(dataDir.toString()).toString());

										p.getPageInfo().getMargin().setBottom(0);
										p.getPageInfo().getMargin().setTop(0);
										p.getPageInfo().getMargin().setRight(0);
										p.getPageInfo().getMargin().setLeft(0);
										p.getParagraphs().add(img);
										d1.save(Paths.get(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("png", "") + "pdf").toString());
										editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("png", "") + "pdf", "");
										m = new File(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("png", "") + "pdf");
										m.delete();
									} else if (str.endsWith("gif")) {
										java.nio.file.Path dataDir = Paths.get(str4);
										com.aspose.pdf.Document d1 = new com.aspose.pdf.Document();
										com.aspose.pdf.Page p = d1.getPages().add();
										Image img = new Image();
										img.setFile(Paths.get(dataDir.toString()).toString());
										p.getPageInfo().getMargin().setBottom(0);
										p.getPageInfo().getMargin().setTop(0);
										p.getPageInfo().getMargin().setRight(0);
										p.getPageInfo().getMargin().setLeft(0);
										p.getParagraphs().add(img);
										d1.save(Paths.get(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("gif", "") + "pdf").toString());
										editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("gif", "") + "pdf", "");
										m = new File(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("gif", "") + "pdf");
										m.delete();
									} else if (str.endsWith("jfif")) {
										java.nio.file.Path dataDir = Paths.get(str4);
										com.aspose.pdf.Document d1 = new com.aspose.pdf.Document();
										com.aspose.pdf.Page p = d1.getPages().add();
										Image img = new Image();
										img.setFile(Paths.get(dataDir.toString()).toString());
										p.getPageInfo().getMargin().setBottom(0);
										p.getPageInfo().getMargin().setTop(0);
										p.getPageInfo().getMargin().setRight(0);
										p.getPageInfo().getMargin().setLeft(0);
										p.getParagraphs().add(img);
										d1.save(Paths.get(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("jfif", "") + "pdf").toString());
										editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("jfif", "") + "pdf", "");
										m = new File(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("jfif", "") + "pdf");
										m.delete();
									} else if (str.endsWith("jpg")) {
										java.nio.file.Path dataDir = Paths.get(str4);
										com.aspose.pdf.Document d1 = new com.aspose.pdf.Document();
										com.aspose.pdf.Page p = d1.getPages().add();
										Image img = new Image();
										img.setFile(Paths.get(dataDir.toString()).toString());
										p.getParagraphs().add(img);
										d1.save(Paths.get(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("jpg", "") + "pdf").toString());
										editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("jpg", "") + "pdf", "");
										m = new File(f21.getAbsolutePath().trim() + File.separator
												+ str.replace("jpg", "") + "pdf");
										m.delete();
									} else if (str.endsWith("xlsx") || str.endsWith("xls") || str.endsWith("xlsm")
											|| str.endsWith("xlsb") || str.endsWith("xltx") || str.endsWith("xltm")
											|| str.endsWith("xlt") || str.endsWith("xml") || str.endsWith("xlam")
											|| str.endsWith("xla") || str.endsWith("xlw") || str.endsWith("xlr")) {

										Workbook book = new Workbook(str4);
										if (str.endsWith("xlsx")) {
											book.save(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlsx", "") + "pdf");
											editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlsx", "") + "pdf", "");
											m = new File(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlsx", "") + "pdf");
											m.delete();
										} else if (str.endsWith("xls")) {
											book.save(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xls", "") + "pdf");
											editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xls", "") + "pdf", "");
											m = new File(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xls", "") + "pdf");
											m.delete();
										} else if (str.endsWith("xlsm")) {
											book.save(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlsm", "") + "pdf");
											editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlsm", "") + "pdf", "");
											m = new File(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlsm", "") + "pdf");
											m.delete();
										} else if (str.endsWith("xlsb")) {
											book.save(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlsb", "") + "pdf");
											editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlsb", "") + "pdf", "");
											m = new File(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlsb", "") + "pdf");
											m.delete();
										} else if (str.endsWith("xltx")) {
											book.save(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xltx", "") + "pdf");
											editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xltx", "") + "pdf", "");
											m = new File(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xltx", "") + "pdf");
											m.delete();
										} else if (str.endsWith("xltm")) {
											book.save(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xltm", "") + "pdf");
											editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xltm", "") + "pdf", "");
											m = new File(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xltm", "") + "pdf");
											m.delete();
										} else if (str.endsWith("xlt")) {
											book.save(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlt", "") + "pdf");
											editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlt", "") + "pdf", "");
											m = new File(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlt", "") + "pdf");
											m.delete();
										} else if (str.endsWith("xml")) {
											book.save(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xml", "") + "pdf");
											editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xml", "") + "pdf", "");
											m = new File(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xml", "") + "pdf");
											m.delete();
										} else if (str.endsWith("xlam")) {
											book.save(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlam", "") + "pdf");
											editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlam", "") + "pdf", "");
											m = new File(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlam", "") + "pdf");
											m.delete();
										} else if (str.endsWith("xla")) {
											book.save(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xla", "") + "pdf");
											editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xla", "") + "pdf", "");
											m = new File(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xla", "") + "pdf");
											m.delete();
										} else if (str.endsWith("xlw")) {
											book.save(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlw", "") + "pdf");
											editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlw", "") + "pdf", "");
											m = new File(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlw", "") + "pdf");
											m.delete();
										} else if (str.endsWith("xlr")) {
											book.save(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlr", "") + "pdf");
											editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlr", "") + "pdf", "");
											m = new File(f21.getAbsolutePath().trim() + File.separator
													+ str.replace("xlr", "") + "pdf");
											m.delete();
										}

									} else if (str.endsWith("pptx")) {
										File f4 = new File(str4);
										FileInputStream fis = new FileInputStream(f4);
										Presentation pres = new Presentation(fis);
										String str9 = f21.getAbsolutePath().trim() + File.separator
												+ str.replace("pptx", "") + "pdf";
										File f7 = new File(str9);
										FileOutputStream fos = new FileOutputStream(f7);
										pres.save(fos, com.aspose.slides.SaveFormat.Pdf);
										editor.addDocumentAttachment(str9, "");
										fos.close();
										f7.delete();

									} else if (str.endsWith("csv")) {
										File f = Csv_Pdf.csv(str4);
										editor.addDocumentAttachment(f.getAbsolutePath(), "");

									} else if (str.endsWith("eml")) {
										MailMessage mes = MailMessage.load(str4);
										mes.save(
												f21.getAbsolutePath().trim() + File.separator + Main_Frame
														.getRidOfIllegalFileNameCharacters(str1).replace("eml", "html"),
												SaveOptions.getDefaultHtml());
										Document document = new Document(f21.getAbsolutePath().trim() + File.separator
												+ Main_Frame.getRidOfIllegalFileNameCharacters(str1).replace("eml",
														"html"));
										document.save(
												f21.getAbsolutePath().trim() + File.separator + Main_Frame
														.getRidOfIllegalFileNameCharacters(str1).replace("html", "pdf"),
												SaveFormat.PDF);
										editor.addDocumentAttachment(
												f21.getAbsolutePath().trim() + File.separator + Main_Frame
														.getRidOfIllegalFileNameCharacters(str1).replace("html", "pdf"),
												"");
										m = new File(f21.getAbsolutePath().trim() + File.separator + Main_Frame
												.getRidOfIllegalFileNameCharacters(str1).replace("html", "pdf"));
										m.delete();
									}

									else if (str.endsWith("pdf") || str.endsWith("ics") || str.endsWith("vcf")) {
										attachment.save(f21.getAbsolutePath().trim() + File.separator
												+ Main_Frame.getRidOfIllegalFileNameCharacters(str1));
										editor.addDocumentAttachment(f21.getAbsolutePath().trim() + File.separator
												+ Main_Frame.getRidOfIllegalFileNameCharacters(str1), "");
										m = new File(f21.getAbsolutePath().trim() + File.separator
												+ Main_Frame.getRidOfIllegalFileNameCharacters(str1));
										m.delete();
									}

									else {
										attachment.save(f21.getAbsolutePath().trim() + File.separator
												+ Main_Frame.getRidOfIllegalFileNameCharacters(str1) + ".msg");
										editor.addDocumentAttachment(
												f21.getAbsolutePath().trim() + File.separator
														+ Main_Frame.getRidOfIllegalFileNameCharacters(str1) + ".msg",
												"");
										m = new File(f21.getAbsolutePath().trim() + File.separator
												+ Main_Frame.getRidOfIllegalFileNameCharacters(str1) + ".msg");
										m.delete();
									}

								}

								catch (Exception e) {
									e.printStackTrace();
									ByteArrayOutputStream eml = new ByteArrayOutputStream();

									attachment.save(eml);

									editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
											attachment.getDisplayName(), "");

								}
								deleteDir(f1);
								deleteDir(f21);
								String filePath = destination_path + File.separator + path + File.separator
										+ "Attachment";
								File file = new File(filePath);
								deleteFolder(file);

							} else {
								ByteArrayOutputStream eml = new ByteArrayOutputStream();
								attachment.save(eml);

								try {
									if (mime_tag.contains("message")) {
										editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
												attachment.getDisplayName() + ".msg", "");
									} else {
										editor.addDocumentAttachment(new ByteArrayInputStream(eml.toByteArray()),
												attachment.getDisplayName(), "");
									}
								} catch (Exception e) {
									e.printStackTrace();
								}

							}

						}
						message.getAttachments().clear();
					}
					editor.save(path5);
				}
			} else if (filetype.equalsIgnoreCase("DOC")) {
				save_Attachment_Word(message, naming_convention, fu);
				doc.save(path5 + ".doc", SaveFormat.DOC);
			} else if (filetype.equalsIgnoreCase("DOCX")) {
				save_Attachment_Word(message, naming_convention, fu);
				doc.save(path5 + ".docx", SaveFormat.DOCX);

			} else if (filetype.equalsIgnoreCase("DOCM")) {
				save_Attachment_Word(message, naming_convention, fu);
				doc.save(path5 + ".docm", SaveFormat.DOCM);

			} else if (filetype.equalsIgnoreCase("TIFF")) {
				save_Attachment_Word(message, naming_convention, fu);
				doc.save(path5 + ".tiff", SaveFormat.TIFF);

			} else if (filetype.equalsIgnoreCase("TXT")) {
				save_Attachment_Word(message, naming_convention, fu);
				doc.save(path5 + ".txt", SaveFormat.TEXT);
			} else if (filetype.equalsIgnoreCase("BMP")) {
				save_Attachment_Word(message, naming_convention, fu);
				doc.save(path5 + ".bmp", SaveFormat.BMP);
			} else if (filetype.equalsIgnoreCase("GIF")) {
				save_Attachment_Word(message, naming_convention, fu);
				for (int pageNumber = 0; pageNumber < doc.getPageCount(); pageNumber++) {
					ImageSaveOptions options11 = new ImageSaveOptions(SaveFormat.GIF);
					PageSet p = new PageSet(pageNumber);
					options11.setPageSet(p);
					doc.save(path5 + "_" + pageNumber + ".gif", options11);
				}

			} else if (filetype.equalsIgnoreCase("JPG")) {
				save_Attachment_Word(message, naming_convention, fu);
				for (int pageNumber = 0; pageNumber < doc.getPageCount(); pageNumber++) {
					ImageSaveOptions options11 = new ImageSaveOptions(SaveFormat.JPEG);
					PageSet p = new PageSet(pageNumber);
					options11.setPageSet(p);
					doc.save(path5 + "_" + pageNumber + ".jpg", options11);
				}

			} else if (filetype.equalsIgnoreCase("PNG")) {
				save_Attachment_Word(message, naming_convention, fu);
				for (int pageNumber = 0; pageNumber < doc.getPageCount(); pageNumber++) {
					ImageSaveOptions options11 = new ImageSaveOptions(SaveFormat.PNG);
					PageSet p = new PageSet(pageNumber);
					options11.setPageSet(p);
					doc.save(path5 + "_" + pageNumber + ".png", options11);
				}

			}
		} catch (Exception e) {

			e.printStackTrace();
		}

		k = Main_Frame.count_destination++;

	}

	void save_Attachment_Word(MapiMessage message, String naming_convention, File f) {
		if (mm.chckbxSavePdfAttachment.isSelected()) {
			MapiAttachmentCollection attachments = message.getAttachments();
			File f1 = null;
			if (attachments.size() > 0) {
				f1 = new File(f.getAbsolutePath() + File.separator + naming_convention);
				if (f1.exists()) {
					f1 = new File(f.getAbsolutePath() + File.separator + naming_convention + "_" + count_destination);
				}
				f1.mkdirs();
				for (MapiAttachment attachment : message.getAttachments()) {

					if (attachment.getMimeTag() != null) {
						mime_tag = attachment.getMimeTag();
					}
					if (mime_tag.contains("message/rfc822")) {
						try {
							attachment
									.save(f1.getAbsolutePath() + File.separator + attachment.getDisplayName() + ".msg");
						} catch (Exception e) {
							attachment.save(f1.getAbsolutePath() + File.separator + attachment.getLongFileName());
						}
					} else {
						try {
							attachment.save(f1.getAbsolutePath() + File.separator + attachment.getDisplayName());
						} catch (Exception e) {
							attachment.save(f1.getAbsolutePath() + File.separator + attachment.getLongFileName());
						}
					}

				}
				message.getAttachments().clear();
			}
		}

	}

	void deleteDir(File file) {
		File[] contents = file.listFiles();
		if (contents != null) {
			for (File f : contents) {
				deleteDir(f);
			}
		}
		if (file.delete()) {
		} else {
			all = new ArrayList<String>();
			all.add(file.getAbsolutePath());
			all.clear();
		}

	}

	static void deleteFolder(File file) {
		if (file.listFiles() != null) {
			for (File subFile : file.listFiles()) {
				if (subFile.isDirectory()) {
					deleteFolder(subFile);
				} else {
					subFile.delete();
				}
			}
		}
		file.delete();
	}

}
