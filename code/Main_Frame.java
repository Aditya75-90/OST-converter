package email.code;

import java.awt.AWTException;
import java.awt.CardLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Cursor;
import java.awt.Desktop;
import java.awt.Font;
import java.awt.Label;
import java.awt.MenuItem;
import java.awt.PopupMenu;
import java.awt.SystemColor;
import java.awt.SystemTray;
import java.awt.Toolkit;
import java.awt.TrayIcon;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.math.BigInteger;
import java.net.InetAddress;
import java.net.URI;
import java.net.URISyntaxException;
import java.net.URL;
import java.net.UnknownHostException;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.TimeUnit;
import java.util.logging.FileHandler;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;
import java.util.regex.Pattern;
import javax.swing.Action;
import javax.swing.ButtonGroup;
import javax.swing.DefaultComboBoxModel;
import javax.swing.DefaultListCellRenderer;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JEditorPane;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JMenu;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPasswordField;
import javax.swing.JPopupMenu;
import javax.swing.JProgressBar;
import javax.swing.JRadioButton;
import javax.swing.JScrollPane;
import javax.swing.JSpinner;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.KeyStroke;
import javax.swing.SpinnerDateModel;
import javax.swing.SpinnerModel;
import javax.swing.SpinnerNumberModel;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.WindowConstants;
import javax.swing.border.EmptyBorder;
import javax.swing.border.LineBorder;
import javax.swing.border.TitledBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.text.DefaultEditorKit;
import javax.swing.text.html.HTMLEditorKit;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.TreeNode;
import javax.swing.tree.TreePath;
import com.aspose.email.Appointment;
import com.aspose.email.Attachment;
import com.aspose.email.Contact;
import com.aspose.email.EWSClient;
import com.aspose.email.EmailClient;
import com.aspose.email.EmlSaveOptions;
import com.aspose.email.ExchangeFolderInfo;
import com.aspose.email.ExchangeFolderInfoCollection;
import com.aspose.email.ExchangeMessageInfo;
import com.aspose.email.ExchangeMessageInfoCollection;
import com.aspose.email.FolderInfo;
import com.aspose.email.FolderInfoCollection;
import com.aspose.email.IConnection;
import com.aspose.email.IEWSClient;
import com.aspose.email.ImapClient;
import com.aspose.email.ImapFolderInfo;
import com.aspose.email.ImapFolderInfoCollection;
import com.aspose.email.ImapMessageInfo;
import com.aspose.email.ImapMessageInfoCollection;
import com.aspose.email.MailAddressCollection;
import com.aspose.email.MailConversionOptions;
import com.aspose.email.MailMessage;
import com.aspose.email.MapiAttachment;
import com.aspose.email.MapiCalendar;
import com.aspose.email.MapiContact;
import com.aspose.email.MapiContactElectronicAddress;
import com.aspose.email.MapiContactEventPropertySet;
import com.aspose.email.MapiContactNamePropertySet;
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
import com.aspose.email.MboxLoadOptions;
import com.aspose.email.MboxrdStorageReader;
import com.aspose.email.MboxrdStorageWriter;
import com.aspose.email.MessageInfo;
import com.aspose.email.MessageInfoCollection;
import com.aspose.email.OlmFolder;
import com.aspose.email.OlmStorage;
import com.aspose.email.PersonalStorage;
import com.aspose.email.SecurityOptions;
import com.aspose.email.TgzReader;
import com.aspose.email.system.io.FileAccess;
import com.aspose.email.system.io.FileMode;
import com.aspose.email.system.io.FileStream;
import com.opencsv.CSVWriter;
import com.toedter.calendar.JDateChooser;
import com.toedter.calendar.JTextFieldDateEditor;

import email.activation.ActivationFrame;
import email.activation.OnlineActivation;
import email.activation.Starting_Frame;
import email.activation.Uninstall;
import email.design.CustomTreeNode;
import it.cnr.imaa.essi.lablib.gui.checkboxtree.CheckboxTree;
import it.cnr.imaa.essi.lablib.gui.checkboxtree.DefaultCheckboxTreeCellRenderer;
import net.Hash;

import javax.swing.JMenuItem;
import javax.swing.JMenuBar;
import javax.swing.border.EtchedBorder;
import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.LayoutStyle.ComponentPlacement;

@SuppressWarnings({ "rawtypes", "deprecation" })

public class Main_Frame extends JFrame {

	JTextFieldDateEditor fromdateeditor;
	JTextFieldDateEditor todateeditor;
	Main_Frame mf;
	static Date fromdate;
	static Date todate;
	ArrayList<Date> fromList = new ArrayList<Date>();
	ArrayList<Date> toList = new ArrayList<Date>();
	static boolean datevalidflag = false;
	String from;
	String to;
	String first = null, middle = null, last = null;
	public static JMenu tools;
	JMenuBar menuBar;
	static File licFileon;
	static File licFileonoff;
	JButton btnActivate;
	JButton btnNewButton;
	JLabel label_backupschedul;
	long maxsize = 0;
	int pstindex = 0;
	JLabel lbl_splitpst;
	public static String projectTitle;
	JCheckBox chckbxMigrateOrBackup;
	JLabel label_convert_pdf_attch_pdf;
	JCheckBox chckbx_convert_pdf_to_pdf;
	JLabel lblNextMigrationStart;
	JButton btnStartTheMigration;
	Map<String, ImageIcon> listimage;
	private Map<String, ImageIcon> imageMap;
	public Map<String, ImageIcon> imageMap_output;
	Cursor cursor = new Cursor(Cursor.HAND_CURSOR);
	DefaultComboBoxModel<String> l_output;
	static List<String> listst = new ArrayList<String>();
	static List<DefaultMutableTreeNode> lists = new ArrayList<DefaultMutableTreeNode>();
	private static final long serialVersionUID = 1L;
	private JPanel contentPane;
	JCheckBox chckbxSavePdfAttachment;
	JButton btnbackup;
	JButton btnStopMigration;
	JSpinner spinner_sizespinner;
	JComboBox comboBox_setsize;
	String destination = "";
	static JFileChooser jFileChooser;
	JButton btn_ques;
	public static JButton	updateBtn;
	static String[] filesfd_img = null;
	static String[] emailsfd_img = null;
	JLabel lblSaveintheSameFolder;
	JCheckBox chckbxSaveInSame;
	JLabel lblTotalMessageCount;

	int accoutcount;
	JLabel lblThisOptionAllows;
	static JFileChooser jFileChooser_destination;
	static PersonalStorage pst;
	Date c1;
	JCheckBox chckbxAutoIncrementBackup;
	JLabel lblNewLabel_5;
	JLabel lblThisOptionAllows_1;
	JLabel lblthirdPartyPassword_1;
	static String OS = System.getProperty("os.name").toLowerCase();
	static String destination_path;
	static JRadioButton rdbtn_MultipleFile;
	int kl = 0;
	Date pdf_date;
	JButton btnI;
	JCheckBox chckbxDeleteEmailFrom;
	JLabel label_Calendarenddate;
	loadingThreadclassformailbox obTh;
	JLabel label_11;
	JLabel label_BR;
	JTextArea textArea_contact;
	String thunderbirdpath = "";
	JLabel lblSubject;
	JLabel label_Calendarsubject;
	JCheckBox chckbxCustomFolderName;
	JPanel panel_progress;
	public static JButton btn_SignIn;
	JLabel lblEnableImap_p3;
	Boolean contatcheck;
	Boolean calendarcheck;
	JLabel lblTurnOffTwo_p3;
	JPanel panel_taskfilter;
	JCheckBox chckbxShowPassword_p2;
	ArrayList<String> pstfolderlist;
	static File file;
	int portmo;
	JLabel lblMakeSureYou;
	static String[] file_sfd=null;
	static String[] email_sfd=null;
	JLabel lblLive_Chat_p3;
	JButton button_stop;
	JCheckBox chckbxRemoveDuplicacy;
	JLabel lblPortNo;
	int portnofiletype;
	int ret;
	int doubleclickcount = 0;
	Date mailfilterstartdate;
	Date mailfilterenddate;
	Date Calenderfilterstartdate;
	Date Calenderfilterenddate;
	Date taskfilterstartdate;
	Date taskfilterenddate;
	JButton btn_Cancel_pane2;
	String date;
	static String filefoldername;
	String foldername6 = "";
	String fileaddress;
	JLabel lblNewLabel_7;
	private JButton btn_next_pane2;
	JLabel lblthirdPartyPassword;
	JScrollPane scrollPane_fortable_p2;
	String filepath;
	static FolderInfo folderInfo;
	static String[] si;
	static String foldername = "";
	static String reportpath;
	static String fname;
	JPanel panel_1;
	JLabel Progressbar;
	JPanel panel_2;
	JPanel panel_2_2_1;
	JPanel panel_3;
	JPanel panel_3_;
	JPanel panel_3_1_2;
	JPanel panel_3_1_1;
	JPanel panel_3_1_1_1;
	JPanel panel_3_2;
	JPanel panel_Loading;
	JLabel label_7;
	JLabel label_8;
	int countcheck = 0;
	JPanel panel_computername;
	String Status = "Completed";
	JPanel panel_4;
	JButton btn_buy;
	JPanel panel_mailfilter;
	Boolean checkdestination = true;
	Boolean checkconvertagain = false;
	public static String fileoption;
	String filetype = "Not selected";
	static CheckboxTree tree;
	static DefaultTableModel modeli;
	static DefaultTreeModel model;
	static DefaultMutableTreeNode root;
	static DefaultTableModel mode;
	static DefaultMutableTreeNode lastNode;
	JDateChooser dateChooser_mail_fromdate;
	JDateChooser dateChooser_mail_tilldate;
	static JComboBox comboBox_FiletypeChooser;
	static JComboBox comboBox_fileDestination_type;
	static JComboBox comboBox;
	private static JTable table_fileinformation;
	static JTextField tf_Destination_Location;
	public static JPanel Cardlayout;
	private static JTable table_fileConvertionreport_panel4;
	static IEWSClient clientforexchange_output;
	static IEWSClient clientforexchange_input;
	static IConnection iconnforimap_input;
	static ImapClient clientforimap_input;
	static IConnection iconnforimap_output;
	static ImapClient clientforimap_output;
	static String mailboxUri = "https://outlook.office365.com/EWS/Exchange.asmx";
	public static String username_p2 = "";
	public static String domain_p2 = "";
	public static String password_p2 = "";
	public static String username_p3 = "";
	public static String domain_p3 = "";
	public static String password_p3 = "";
	static long count_destination;
	String demoversion = "Demo";
	String Fullversion = "Full";
	String filetemp = "";
	ImapFolderInfo imapFolderInfo;
	Logger logger;

	String buyurl = "";
	String infourl = "";
	String helpurl = "";
	static int versiontype;
	String enableimapgmail = "";
	String allowlesssecureappgmail = "";
	String turnofftwostepverificationyahoo = "";
	String generatethirdpartypassyahoo = "";
	String createnewpasswordfor365 = "";
	String multifatcorauthicationfor365 = "";
	String createnewpasswordforhotmail = "";
	String createapppasswordforaol = "";
	String turnofftwostepverificationgmail = "";
	String turnofftwostepverificationZohoMail = "";
	String turnofftwostepverificationYandexMail = "";
	All_Data ad = new All_Data();
	boolean stop = false;
	boolean stop_tree = false;
	static Main_Frame frame;
	public static String calendertime;
	JLabel lbl_connecting_p2;
	JLabel lbl_DomainName_computername;
	static JLabel lbl_progressreport;
	public static JTextField textField_username_p2;
	public static JPasswordField passwordField_p2;
	private JTextField textField_Domainname_p2;
	static JCheckBox chckbx_Mail_Filter;
	static JCheckBox chckbx_calender_box;
	public static Calendar cal;
	static String parentname = "";
	static String path = "";
	String path1 = "";
	static long foldermessagecount = 0;
	static HashMap<String, Long> olmpathmap = new HashMap<String, Long>();
	static OlmStorage storage;
	private JButton btnDowloadReport;
	private JTextField textField_username_p3;
	private JPasswordField passwordField_p3;
	static Iterator<MapiMessage> it;
	private JTextField textField_domain_name_p3;
	static String Folder;
	static String hostName = "";
	JButton btn_Destination;
	JButton btn_Sign_p3;
	JButton btn_Previous_pane2;
	JButton btn_converter;
	JButton btn_previous_p3;
	JButton btn_signout_p3;
	JButton btnStop;
	JButton btnTempPath;
	File f;
	Thread th;
	JPanel panel_Calender;
	static File index;
	private JCheckBox chckbxShowPassword_p3;
	JEditorPane editorPane;
	private JLabel lbl_connecting_p3;
	static String mess;
	FolderInfo info1;
	static JProgressBar progressBar_message_p3;
	private JPanel panel_loginpanel;
	private JPanel panel_selectfile;
	JPanel panel_mailfilt7;
	JLabel lbl_Domain;
	JDateChooser dateChooser_task_end_date;
	JLabel lblPleaseWatTable;
	private JPanel inner_cardlayout;
	JLabel lbl_Email;
	JLabel lbl_subject;
	JLabel lblNew_setemail;
	FolderInfo info = new FolderInfo();
	PersonalStorage ost;
	JRadioButton rdbtnSingleFile;
	ExchangeFolderInfo exchangeFolderInfo;
	OlmFolder folderi;
	JCheckBox task_box;
	JDateChooser dateChooser_task_start_date;
	JLabel lblNew_setsubject;
	private final ButtonGroup buttonGroup_file = new ButtonGroup();
	private JLabel lbl_Date;
	private JLabel label_date;
	private JButton btnViewer;
	private JButton btnAttachment;
	private JPanel innerCardlayout;
	private JPanel viewer;
	private JPanel attachment;
	private JScrollPane scrollPane;
	private JTable table_1;
	private JScrollPane scrollPane_1;
	private JLabel lblNewLabel_2;
	private JLabel lblReadingFoldersPlease;
	private JLabel lblNewLabel_3;
	private JLabel lblNewLabel_4;
	Boolean input = false;
	Boolean output = false;
	JCheckBox chckbxRestoreToDefault;
	public static Boolean demo = true;
	JButton btnSavingLog;
	private JButton btnSavingLog_1;
	private JButton btnConvertAgain;
	private JButton btn_next_pane1;
	private JLabel label_6;
	private JLabel label_9;
	private JLabel label_10;
	private List<MailMessage> listmail = new ArrayList<MailMessage>();
	private List<MapiMessage> listmapi = new ArrayList<MapiMessage>();
	private List<MessageInfo> listPSTOSTgemesingo = new ArrayList<MessageInfo>();
	private List<ImapFolderInfo> listFolderinfo = new ArrayList<ImapFolderInfo>();
	private List<ImapFolderInfo> listFolderinfofinal = new ArrayList<ImapFolderInfo>();
	private List<String> listFolderinfostring = new ArrayList<String>();
	private List<ExchangeFolderInfo> listExchangemesingo = new ArrayList<ExchangeFolderInfo>();
	private List<ExchangeFolderInfo> listExchangemesingos = new ArrayList<ExchangeFolderInfo>();
	private List<ExchangeFolderInfo> listExchangdinal = new ArrayList<ExchangeFolderInfo>();
	List<String> listduplicacy = new ArrayList<String>();
	private List<String> listdupliccal = new ArrayList<String>();
	private List<String> listduplictask = new ArrayList<String>();
	private List<String> listdupliccontact = new ArrayList<String>();
	private JLabel label_13;
	Boolean Stoppreview = false;
	String version;
	JPopupMenu menu;
	private JTextField textField_hi;
	public String logpath = "";
	public String temppath = "";
	public JTextField textField_1;
	private JTextField tf_portNo_p2;
	private JTextField tf_portNo_p3;
	private JLabel lblLiveChat;
	private JLabel lblPleaseMakeSure;
	private JLabel lblEnabledImap;
	private JLabel lblTurnOffTwo;
	private JPanel panel;
	JLabel label_12;
	private JPanel panel_Callendar;
	private JPanel panel_Contact;
	private JLabel label_calendarstartdate;
	private JLabel lblFullName;
	private JLabel label_contactfullname;
	private JLabel lblEMailAd;
	private JLabel lblCompany;
	private JLabel label_contactemail;
	private JLabel label_contactcompany;
	private JLabel lblPhoneNo;
	private JLabel label_contactphonenumber;
	private JLabel label_contacticon;
	private JLabel label_Calendaricon;
	private JPanel panel_office365BR;
	private JTextField textField_customfolder;
	private JPanel panel_9;
	private JCheckBox chckbxMaintainFolderHeirachy;
	private JLabel lblemailAddress_1;
	private JLabel lblSavesbackupmigrateAs;
	private JLabel lblthisOptionIs;
	private JLabel lblthisOptionIs_1;
	private JLabel label_Remove_Duplicate;
	private JLabel label_Maintain_Folder_Hierarchy;
	private JLabel label_Save_PDF_Attachments_Separately;
	private JLabel label_Free_up_Server_Space;
	private JLabel lblSkip_Previously_Migrated_Items;

	public static ResultSet rs;
	public static Connection schsqlconnection = null;
	public static Statement schsqlstmt = null;
	JRadioButton rdbtnEveryday;
	JRadioButton rdbtnEveryWeek;
	JRadioButton rdbtnOnce;
	Boolean nexttime = true;
	JComboBox comboBox_weekdays;
	JComboBox comboBox_MonthDay;
	public static Long starttime;
	public static Long nextTime = null;
	public static Long nextcount = null;
	public static Date nextDate = null;
	public TrayIcon trayIcon;
	public static String once;
	public static String everyday;
	public static String OnWeekday;
	public static String everyweek;
	public static String OnMonthday;
	public static String nextPAth;
	public static String everymonth;
	public static String nextfiletype;
	public static String removeduplica;
	public static String maintainfolderh;
	public static String savepdfattac;
	public static String freeupserverspace;
	public static List<String> listnextduration = new ArrayList<String>();
	public static List<Long> listnextstarttime = new ArrayList<Long>();
	public static List<Long> listnextendtime = new ArrayList<Long>();
	public static List<Long> listnextcount = new ArrayList<Long>();
	JCheckBox chckbxSetBackupSchedule;
	JDateChooser dateChooserNextSchedular;
	JSpinner spinner;
	Label label_14;
	Calendar caltime = null;
	JRadioButton rdbtnOnWeekDay;
	private final ButtonGroup buttonGroup_Schedulling = new ButtonGroup();
	private JRadioButton rdbtnOnmonthDay;
	private JRadioButton rdbtnEveryMonth;
	Boolean nextstart = false;
	private JCheckBox chckbx_splitpst;
	private JDateChooser dateChooser_newFrom;
	private JDateChooser dateChooser_newTo;
	private JTable table;
	private JScrollPane scrollPane_2;
	JCheckBox DateFilter;
	private JButton addButton;
	private JButton removeButton;
	public static String messageboxtitle = "";
	private JLabel lblNewLabel_12;
	private JButton btnNewButton_1;
	static String strSerialNumber;
	static String hashKey;
	public static main_multiplefile multi =null;
	public static Main_Frame mf1=null;
	

	static {
		System.out.println("stsrt from here");
		
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		//	frame = new Starting_Frame();
		} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
				| UnsupportedLookAndFeelException e1) {

			e1.printStackTrace();
		}
		
		if (System.getProperty("os.name").toLowerCase().contains("windows")) {
			FileHandler fh;
			try {
				fh = new FileHandler(System.getProperty("java.io.tmpdir") + File.separator + "chilkat.log");
				// logger.addHandler(fh);
				SimpleFormatter formatter = new SimpleFormatter();
				fh.setFormatter(formatter);
				// logger.info("My Log File");
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
			try {
				InputStream in = Main_Frame.class.getResourceAsStream("/chilkat.dll");
				byte[] buffer = new byte[1024];
				int read = -1;
				File temp = new File(new File(System.getProperty("java.io.tmpdir")), "chilkat.dll");
				int i = 0;
				while (temp.exists()) {
					temp = new File(new File(System.getProperty("java.io.tmpdir")), "chilkat" + i + ".dll");
					i++;
				}
				FileOutputStream fos = null;
				try {
					fos = new FileOutputStream(temp);
				} catch (FileNotFoundException e) {
					e.printStackTrace();
					// logger.warning(e.getMessage());
				}
				try {
					while ((read = in.read(buffer)) != -1) {
						fos.write(buffer, 0, read);
					}
				} catch (IOException e) {
					// logger.warning(e.getMessage());
				}
				try {
					fos.close();
				} catch (IOException e) {
					// logger.warning(e.getMessage());
				}
				try {
					in.close();
				} catch (IOException e) {
					// logger.warning(e.getMessage());
				}
				try {
					System.load(temp.getAbsolutePath());
				} catch (Error er) {
					in = Main_Frame.class.getResourceAsStream("/chilkatX64.dll");
					buffer = new byte[1024];
					read = -1;
					temp = new File(new File(System.getProperty("java.io.tmpdir")), "chilkatX64.dll");
					i = 0;
					while (temp.exists()) {
						temp = new File(new File(System.getProperty("java.io.tmpdir")), "chilkatX64" + i + ".dll");
						i++;
					}
					fos = null;
					try {
						fos = new FileOutputStream(temp);
					} catch (FileNotFoundException e) {
						e.printStackTrace();
						// logger.warning(e.getMessage());
					}
					try {
						while ((read = in.read(buffer)) != -1) {
							fos.write(buffer, 0, read);
						}
					} catch (IOException e) {
						// logger.warning(e.getMessage());
					}
					try {
						fos.close();
					} catch (IOException e) {
						// logger.warning(e.getMessage());
					}
					try {
						in.close();
					} catch (IOException e) {
						// logger.warning(e.getMessage());
					}
					System.load(temp.getAbsolutePath());
				}
			} catch (UnsatisfiedLinkError | Exception e) {
				e.printStackTrace();
				// logger.warning(e.getMessage());
				// System.loadLibrary("chilkat64");
				// logger.warning(e.getMessage());
				// System.loadLibrary("chilkat");
				// System.load("‪D:\\EXE\\chilkat-9.5.0-jdk11-x64\\chilkat.dll");
			}
		} else {
			// Logger logger = Logger.getLogger("MyLog1");
			FileHandler fh;
			try {

				fh = new FileHandler(System.getProperty("java.io.tmpdir") + File.separator + "chilkat.log");
				// logger.addHandler(fh);
				SimpleFormatter formatter = new SimpleFormatter();
				fh.setFormatter(formatter);
			} catch (SecurityException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
			try {
				System.out.println("605..");
				InputStream in = Main_Frame.class.getResourceAsStream("/libchilkat.jnilib");
				byte[] buffer = new byte[1024];
				int read = -1;
				File temp = new File(new File(System.getProperty("java.io.tmpdir")), "libchilkat.jnilib");
				System.out.println("610");
				int i = 0;
				while (temp.exists()) {
					temp.delete();
					temp = new File(new File(System.getProperty("java.io.tmpdir")), "libchilkat" + i + ".jnilib");
					i++;
				}
				FileOutputStream fos = null;
				try {
					fos = new FileOutputStream(temp);
				} catch (FileNotFoundException e) {
					e.printStackTrace();
					// logger.warning(e.getMessage());
				}
				try {
					while ((read = in.read(buffer)) != -1) {
						fos.write(buffer, 0, read);
					}
				} catch (IOException e) {
					// logger.warning(e.getMessage());
				}
				try {
					fos.close();
				} catch (IOException e) {
					// logger.warning(e.getMessage());
				}
				try {
					in.close();
				} catch (IOException e) {

					// logger.warning(e.getMessage());
				}
				System.out.println("640.." + temp.getAbsolutePath());
				System.load(temp.getAbsolutePath());

				temp.delete();
			} catch (UnsatisfiedLinkError | Exception e) {
				e.printStackTrace();
				// logger.warning(e.getMessage());
				// System.loadLibrary("chilkat64");
				// logger.warning(e.getMessage());
				// System.loadLibrary("chilkat");
				// System.load("â€ªD:\\EXE\\chilkat-9.5.0-jdk11-x64\\chilkat.dll");
			}

		}
	}

	@SuppressWarnings({ "unchecked" })
	public Main_Frame(Boolean demo, int versiontype) {
		Main_Frame.demo = demo;
		if (demo)
			accoutcount = 3;
		messageboxtitle = All_Data.messageboxtitle;
		projectTitle = All_Data.messageboxtitle;

		buyurl = All_Data.buyurl;
		infourl = All_Data.infourl;
		helpurl = All_Data.helpurl;
		Main_Frame.versiontype = versiontype;
		enableimapgmail = ad.enableimapgmail;
		allowlesssecureappgmail = ad.allowlesssecureappgmail;
		turnofftwostepverificationyahoo = ad.turnofftwostepverificationyahoo;
		generatethirdpartypassyahoo = ad.generatethirdpartypassyahoo;
		createnewpasswordfor365 = ad.createnewpasswordfor365;
		multifatcorauthicationfor365 = All_Data.multifatcorauthicationfor365;
		createnewpasswordforhotmail = ad.createnewpasswordforhotmail;
		createapppasswordforaol = ad.createapppasswordforaol;
		turnofftwostepverificationgmail = ad.turnofftwostepverificationgmail;
		turnofftwostepverificationZohoMail = ad.turnofftwostepverificationZohoMail;
		turnofftwostepverificationYandexMail = ad.turnofftwostepverificationYandexMail;
		version = All_Data.version;
		if (versiontype == 1) {
			messageboxtitle = messageboxtitle + "(Single License)";
			accoutcount = 50;
		} else if (versiontype == 2) {
			accoutcount = 200;
			messageboxtitle = messageboxtitle + "(Admin License)";
		} else if (versiontype == 3) {
			accoutcount = 500;
			messageboxtitle = messageboxtitle + "(Technical License)";
		} else if (versiontype == 4) {
			accoutcount = 500;
			messageboxtitle = messageboxtitle + "(Enterprise License)";
		}
		//
		if (demo) {
			String title = messageboxtitle.replace("(Enterprise License)", " ").replace("(Technical License)", " ");
			String center = title + "(" + demoversion + ")";
			messageboxtitle = center;
			setTitle(center);
			System.out.println("bbbbbbbbbbb");
		} else {

			String center = messageboxtitle;
			setTitle(center);
		}
		try {
			InputStream is = Main_Frame.class.getResourceAsStream("/data.txt");
			byte[] buff = new byte[is.available()];
			is.read(buff);
			ency td = new ency();
			String decrpt = td.decrypt(buff);
			OutputStream outStream = new FileOutputStream(
					System.getProperty("java.io.tmpdir") + File.separator + "Email.txt");
			outStream.write(decrpt.getBytes());
			is.close();
			outStream.flush();
			outStream.close();
			File file = new File(System.getProperty("java.io.tmpdir") + File.separator + "Email.txt");
			if (file.exists()) {

				try {

					com.aspose.email.License lic_email = new com.aspose.email.License();
					com.aspose.words.License lic_word = new com.aspose.words.License();
					com.aspose.pdf.License lic_pdf = new com.aspose.pdf.License();
					lic_email.setLicense(file.getPath());
					lic_word.setLicense(file.getPath());
					lic_pdf.setLicense(file.getPath());
					file.delete();
				} catch (Exception e1) {

					if (e1.getMessage().contains("Culture Name:")) {
						Locale.setDefault(new Locale("en", "US"));
						com.aspose.email.License lic_email = new com.aspose.email.License();
						com.aspose.words.License lic_word = new com.aspose.words.License();
						com.aspose.pdf.License lic_pdf = new com.aspose.pdf.License();
						com.aspose.cells.License lic_cells = new com.aspose.cells.License();
						lic_email.setLicense(file.getPath());
						lic_word.setLicense(file.getPath());
						lic_pdf.setLicense(file.getPath());
						lic_cells.setLicense(file.getPath());
						file.delete();
					} else {
						JOptionPane.showMessageDialog(null,
								"Getting error while loading license. Please connect with live chat support.");
					}

				}
			}
			file = new File(System.getProperty("java.io.tmpdir") + File.separator + "Email.txt");
			if (file.exists()) {
				file.delete();
			}
		} catch (Exception e1) {

		}

		addWindowListener(new WindowAdapter() {

//			public void windowClosing(WindowEvent arg0) {
//
//				if (!SystemTray.isSupported()) {
//					String warn = "Do you want to close the Application?";
//					int ans = JOptionPane.showConfirmDialog(Main_Frame.this, warn, messageboxtitle,
//							JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
//							new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
//					if (ans == JOptionPane.YES_OPTION) {
//
//						openBrowser(infourl);
//						System.exit(0);
//					}
//
//				}
//
//				SystemTray systemTray = SystemTray.getSystemTray();
//
//				PopupMenu trayPopupMenu = new PopupMenu();
//				MenuItem action = new MenuItem("Show");
//				action.addActionListener(new ActionListener() {
//					@Override
//					public void actionPerformed(ActionEvent e) {
//
//						setVisible(true);
//					}
//				});
//				trayPopupMenu.add(action);
//
//				MenuItem close = new MenuItem("Exit");
//				close.addActionListener(new ActionListener() {
//					@Override
//					public void actionPerformed(ActionEvent e) {
//
//						String warn = "Do you want to close the Application?";
//						int ans = JOptionPane.showConfirmDialog(Main_Frame.this, warn, messageboxtitle,
//								JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
//								new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
//						if (ans == JOptionPane.YES_OPTION) {
//
//							if (demo) {
//
//								openBrowser(infourl);
//
//							}
//
//							System.exit(0);
//						}
//
//					}
//				});
//				trayPopupMenu.add(close);
//
//				trayIcon = new TrayIcon(
//						Toolkit.getDefaultToolkit().getImage(Main_Frame.class.getResource("/128x128.png")),
//						messageboxtitle, trayPopupMenu);
//				trayIcon.setImageAutoSize(true);
//
//				try {
//
//					TrayIcon[] icons = (TrayIcon[]) SystemTray.getSystemTray().getTrayIcons();
//
//					boolean check = false;
//					for (int i = 0; i < icons.length; i++) {
//
//						if (icons[i].getImage().equals(trayIcon.getImage())) {
//							check = true;
//							break;
//						}
//					}
//
//					if (!check) {
//						systemTray.add(trayIcon);
//						trayIcon.displayMessage("Tool Added in Tray ", " ", TrayIcon.MessageType.NONE);
//					} else {
//						System.out.println("tool already in tray");
//					}
//				} catch (AWTException awtException) {
//					awtException.printStackTrace();
//				}
//				System.out.println("end of main");
//				setVisible(false);
//
//			}
			
			public void windowClosing(WindowEvent arg0) {
				String warn = "Do you want to close the Application?";
				int ans = JOptionPane.showConfirmDialog(Main_Frame.this, warn, messageboxtitle,
						JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {
					setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
					System.exit(0);
					Main_Frame.this.dispose();
				}
			}
		});

		setIconImage(Toolkit.getDefaultToolkit().getImage(Main_Frame.class.getResource("/128x128.png")));

		setLocationRelativeTo(null);
		setResizable(true);

		setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
		setBounds(100, 100, 1081, 718);
		contentPane = new JPanel();
		contentPane.setBackground(Color.WHITE);
		contentPane.setBorder(null);
		setContentPane(contentPane);
		
				JButton btn_help = new JButton("");
				btn_help.setToolTipText("Click here for software guide.");
				btn_help.addActionListener(new ActionListener() {
					public void actionPerformed(ActionEvent e) {
						openBrowser(helpurl);
					}
				});
				btn_help.addMouseListener(new MouseAdapter() {

					public void mouseEntered(MouseEvent arg0) {
						btn_help.setIcon(new ImageIcon(Main_Frame.class.getResource("/about-hvr-btn.png")));
					}

					public void mouseExited(MouseEvent e) {
						btn_help.setIcon(new ImageIcon(Main_Frame.class.getResource("/about-btn.png")));
					}
				});
								//		lblNewLabel_12.setIcon(new ImageIcon(AboutDialog.class.getResource("/arrow.png")));
								
										JButton btn_info = new JButton("");
										btn_info.setToolTipText("Click here for software information.");
										btn_info.addActionListener(new ActionListener() {

											public void actionPerformed(ActionEvent e) {
												AboutDialog ab;
												if (demo) {
													ab = new AboutDialog(frame, true, "Demo");

												} else {
													String aboutlic = "";

													if (versiontype == 1) {
														aboutlic = "Single License";
													} else if (versiontype == 2) {
														aboutlic = "Admin License";
													} else if (versiontype == 3) {
														aboutlic = "Technical License";
													} else if (versiontype == 4) {
														aboutlic = "Enterprise License";
													}

													ab = new AboutDialog(frame, true, aboutlic);
												}
												ab.setLocationRelativeTo(frame);
												ab.setVisible(true);

											}
										});
										btn_info.addMouseListener(new MouseAdapter() {

											public void mouseEntered(MouseEvent arg0) {
												btn_info.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-hvr-btn.png")));
											}

											public void mouseExited(MouseEvent e) {
												btn_info.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-btn.png")));
											}
										});
										
												btn_info.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-btn.png")));
												btn_info.setRolloverEnabled(false);
												btn_info.setRequestFocusEnabled(false);
												btn_info.setOpaque(false);
												btn_info.setFocusable(false);
												btn_info.setFocusTraversalKeysEnabled(false);
												btn_info.setFocusPainted(false);
												btn_info.setDefaultCapable(false);
												btn_info.setContentAreaFilled(false);
												btn_info.setBorderPainted(false);
				
						btn_help.setIcon(new ImageIcon(Main_Frame.class.getResource("/about-btn.png")));
						
								btn_help.setRolloverEnabled(false);
								btn_help.setRequestFocusEnabled(false);
								btn_help.setOpaque(false);
								btn_help.setFocusable(false);
								btn_help.setFocusTraversalKeysEnabled(false);
								btn_help.setFocusPainted(false);
								btn_help.setDefaultCapable(false);
								btn_help.setContentAreaFilled(false);
								btn_help.setBorderPainted(false);

		Cardlayout = new JPanel();
		Cardlayout.setBackground(Color.LIGHT_GRAY);
		Cardlayout.setLayout(new CardLayout(0, 0));

		panel_1 = new JPanel();
		panel_1.setBackground(Color.WHITE);
		Cardlayout.add(panel_1, "panel_1");

		menu = new JPopupMenu();
		menu.setBackground(Color.WHITE);
		Action cut = new DefaultEditorKit.CutAction();
		cut.putValue(Action.NAME, "Cut");
		cut.putValue(Action.ACCELERATOR_KEY, KeyStroke.getKeyStroke("control X"));
		menu.add(cut);

		Action copy = new DefaultEditorKit.CopyAction();
		copy.putValue(Action.NAME, "Copy");
		copy.putValue(Action.ACCELERATOR_KEY, KeyStroke.getKeyStroke("control C"));
		menu.add(copy);

		Action paste = new DefaultEditorKit.PasteAction();
		paste.putValue(Action.NAME, "Paste");
		paste.putValue(Action.ACCELERATOR_KEY, KeyStroke.getKeyStroke("control V"));

		menu.add(paste);

		Action selectAll = new SelectAll();
		menu.add(selectAll);
		

		DefaultComboBoxModel<String> l1 = new DefaultComboBoxModel<>();
		for (int i1 = 0; i1 < ad.sop.length; i1++) {

			l1.addElement(ad.sop[i1]);

		}

		comboBox_FiletypeChooser = new JComboBox(ad.sop);

		if (ad.sop.length == 1) {
			comboBox_FiletypeChooser.setEnabled(false);
		}

		imageMap = createImageMap_input(l1);
		comboBox_FiletypeChooser.setRenderer(new ListRenderer());

		SwingUtilities.invokeLater(new Runnable() {

			public void run() {
				comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
			}
		});
		comboBox_FiletypeChooser.setFont(new Font("Tahoma", Font.BOLD, 15));
		 btn_buy = new JButton("");
		 rdbtnSingleFile = new JRadioButton("Single File/Multiple Files");
		comboBox_FiletypeChooser.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent args) {

				if (demo) {
					btn_buy.setVisible(true);
				}
				
				buttonGroup_file.clearSelection();
				rdbtnSingleFile.setSelected(true);
				btn_next_pane1.setEnabled(true);
				chckbxSaveInSame.setVisible(true);
				lblSaveintheSameFolder.setVisible(true);
				btn_signout_p3.setVisible(false);
				panel_computername.setVisible(false);
				task_box.setVisible(false);
				panel_taskfilter.setVisible(false);
				lblNewLabel_7.setVisible(false);
				tf_portNo_p2.setVisible(false);
				comboBox.removeItem("Original File Name");
				textField_Domainname_p2.setText("");
				textField_username_p2.setText("");
				passwordField_p2.setText("");
				temppath = textField_1.getText();
				lblPleaseMakeSure.setVisible(true);
				chckbxMaintainFolderHeirachy.setVisible(false);
				label_Maintain_Folder_Hierarchy.setVisible(false);
				chckbxSetBackupSchedule.setVisible(false);
				label_backupschedul.setVisible(false);
				chckbxSetBackupSchedule.setSelected(false);
				chckbxDeleteEmailFrom.setVisible(false);

				label_Free_up_Server_Space.setVisible(false);

				chckbxAutoIncrementBackup.setVisible(false);
				lblSkip_Previously_Migrated_Items.setVisible(false);
				lblEnabledImap.setVisible(true);
				lblTurnOffTwo.setVisible(true);

				if (args.getSource() == comboBox_FiletypeChooser) {
					
					System.out.println("here we riched 1177");
					JComboBox cbo = (JComboBox) args.getSource();
					fileoption = (String) cbo.getSelectedItem();
					System.out.println("this is line no " + fileoption);
				}
				textField_hi.setText(System.getProperty("user.home") + File.separator + "Documents");
				button_stop.setVisible(true);
				btn_next_pane1.setEnabled(true);
				if (fileoption.equalsIgnoreCase("OLM File (.olm)")
						|| fileoption.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
						|| fileoption.equalsIgnoreCase("Exchange Offline Storage (.ost)")
						|| fileoption.equalsIgnoreCase("Nodes Storage (.nsf)")) {

					if (fileoption.equalsIgnoreCase("OLM File (.olm)")) {
						chckbxMaintainFolderHeirachy.setVisible(true);
						label_Maintain_Folder_Hierarchy.setVisible(true);

						label_7.setIcon(new ImageIcon(Main_Frame.class.getResource("/olm.png")));

					} else if (fileoption.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
						chckbxMaintainFolderHeirachy.setVisible(true);
						label_Maintain_Folder_Hierarchy.setVisible(true);
						label_7.setIcon(new ImageIcon(Main_Frame.class.getResource("/outlook.png")));

					} else if (fileoption.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {
						chckbxMaintainFolderHeirachy.setVisible(true);
						label_Maintain_Folder_Hierarchy.setVisible(true);
						label_7.setIcon(new ImageIcon(Main_Frame.class.getResource("/ost.png")));

					}

					input = false;
					CardLayout card = (CardLayout) inner_cardlayout.getLayout();
					card.show(inner_cardlayout, "panel_selectfile");
					panel_mailfilter.setVisible(true);

					chckbx_Mail_Filter.setVisible(true);

					task_box.setVisible(true);

					panel_taskfilter.setVisible(true);

					dateChooser_task_end_date.setVisible(true);
					dateChooser_task_start_date.setVisible(true);

				} else if (fileoption.equalsIgnoreCase("MBOX")) {
					label_7.setIcon(new ImageIcon(Main_Frame.class.getResource("/mbox.png")));
					panel_mailfilter.setVisible(true);
					input = false;
					CardLayout card = (CardLayout) inner_cardlayout.getLayout();

					card.show(inner_cardlayout, "panel_selectfile");
					chckbx_Mail_Filter.setVisible(true);
				} else if (fileoption.equalsIgnoreCase("OFFICE 365 Backup & Restore")
						|| fileoption.equalsIgnoreCase("PST to Office 365")
						|| fileoption.equalsIgnoreCase("Exchange Backup & Restore")
						|| fileoption.equalsIgnoreCase("Mbox to Office 365")) {

					if (fileoption.equalsIgnoreCase("Exchange Backup & Restore")) {
						label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/exchange-topbar.png")));
						label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/exchange-topbar.png")));
					} else {
						label_11.setIcon(new ImageIcon(Main_Frame.class.getResource("/office365-mian-topabr.png")));
						label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/office365-mian-topabr.png")));
					}
					btn_SignIn.setToolTipText("Click here to Sign in to " + fileoption + ".");
					panel_mailfilter.setVisible(true);
					input = false;
					lblThisOptionAllows_1.setText(
							"This option allows the data to be read from PST/OST File and store it in office 365");
					lblThisOptionAllows.setText(
							"This option allows us to take backup from office 365. It can take backup in 25+ format");
					label_BR.setText("Office 365");
					btnbackup.setVisible(true);
					lblThisOptionAllows.setVisible(true);
					btnI.setVisible(true);
					if (fileoption.equalsIgnoreCase("Exchange Backup & Restore")) {
						label_BR.setText("Live Exchange");
						lblThisOptionAllows_1.setText(
								"This option allows the data to be read from PST/OST File and store it in Exchange");
						lblThisOptionAllows.setText(
								"This option allows us to take backup from Exchange. It can take backup in 25+ format");
					} else if (fileoption.equalsIgnoreCase("Mbox to Office 365")) {
						logpath = textField_hi.getText();
						temppath = textField_1.getText();
						tools.setVisible(false);
						restore restore = new restore(Main_Frame.this, demo, messageboxtitle);

						if (demo) {
							if (Starting_Frame.check) {
								restore.setVisible(true);
							}
						} else {
							restore.setVisible(true);
						}
						restore.setLocationRelativeTo(null);

						restore.setResizable(false);
						dispose();
					} else if (fileoption.equalsIgnoreCase("PST to Office 365")) {
						tools.setVisible(false);
						logpath = textField_hi.getText();
						temppath = textField_1.getText();
						fileoption = "MICROSOFT OUTLOOK (.pst)";
						restore restore = new restore(Main_Frame.this, demo, messageboxtitle);
						if (demo) {
							if (Starting_Frame.check) {
								restore.setVisible(true);
							}
						} else {
							restore.setVisible(true);
						}
						restore.setLocationRelativeTo(null);

						restore.setResizable(false);
						dispose();
					}
					CardLayout card = (CardLayout) inner_cardlayout.getLayout();
					card.show(inner_cardlayout, "panel_office365BR");
					chckbx_Mail_Filter.setVisible(true);
				} else if (fileoption.equalsIgnoreCase("Thunderbird")) {

					String str = null;
					if (OS.contains("windows")) {
						str = System.getenv("APPDATA") + File.separator + "Thunderbird" + File.separator + "Profiles";
					} else {
						str = System.getProperty("user.home") + File.separator + "Library" + File.separator
								+ "Thunderbird" + File.separator + "Profiles";
					}

					if (new File(str).exists()) {

						panel_mailfilter.setVisible(true);
						input = false;
						chckbx_Mail_Filter.setVisible(true);
						logger = logFile();
						cal = Calendar.getInstance();
						calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());
						logger = logFile();
						logger.info("Start Time : " + calendertime + System.lineSeparator() + "File Type : "
								+ fileoption + "                         " + System.lineSeparator()
								+ "======================================================================");

						File[] f = new File(str).listFiles();
						for (File fl : f) {
							if (fl != null) {
								if (fl.isDirectory()) {
									String filename = fl.getName();
									String extension = filename.substring(filename.lastIndexOf(".") + 1,
											filename.length());
									String ext = "default-release";

									if (!ext.equals(extension)) {
										ext = "default";
									}

									if (ext.equals(extension)) {

										thunderbirdpath = str;
										String defaultfolder = fl.getName();

										str = str + File.separator + defaultfolder;
										try {
											InetAddress addr = InetAddress.getLocalHost();
											hostName = addr.getHostName();
											CardLayout card1 = (CardLayout) inner_cardlayout.getLayout();
											card1.show(inner_cardlayout, "panel_selectfile");
											model = (DefaultTreeModel) tree.getModel();
											root = new DefaultMutableTreeNode(hostName);
											model.setRoot(root);
											DefaultMutableTreeNode node = new DefaultMutableTreeNode(
													new File(str).getAbsolutePath());
											root.add(node);
											readThunderbird(new File(str), node);
											Icon open = new ImageIcon(
													Main_Frame.class.getResource("/Open-folder-accept-icon.png"));
											Icon close = new ImageIcon(
													Main_Frame.class.getResource("/closed-folder-add-icon.png"));
											Icon Ram = new ImageIcon(Main_Frame.class.getResource("/leaf-icon.png"));
											DefaultCheckboxTreeCellRenderer render = (DefaultCheckboxTreeCellRenderer) tree
													.getCellRenderer();
											render.setClosedIcon(close);
											render.setOpenIcon(open);
											render.setLeafIcon(Ram);

											tree.expandRow(0);
											tree.expandAll();
											CardLayout card = (CardLayout) Cardlayout.getLayout();
											card.show(Cardlayout, "panel_2");
										} catch (Exception e) {

											e.printStackTrace();
										}

										break;
									} else {
										JOptionPane.showMessageDialog(frame, "path not found", messageboxtitle,
												JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
										SwingUtilities.invokeLater(new Runnable() {

											public void run() {
												comboBox_FiletypeChooser.setSelectedItem("MBOX");
											}
										});

									}
								}
							}
						}
					} else {
						JOptionPane.showMessageDialog(frame,
								"Thunderbird not installed Please install it and try again", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));
						if (!projectTitle.contains("Mail Migration")) {
							System.exit(0);
						} else {
							SwingUtilities.invokeLater(new Runnable() {

								public void run() {
									comboBox_FiletypeChooser.setSelectedItem("MBOX");
								}
							});
						}
					}

				} else if (fileoption.equalsIgnoreCase("Apple Mail")) {

					if (OS.contains("mac")) {

						panel_mailfilter.setVisible(true);
						input = false;
						chckbx_Mail_Filter.setVisible(true);
						logger = logFile();
						cal = Calendar.getInstance();
						calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());
						logger = logFile();
						logger.info("Start Time : " + calendertime + System.lineSeparator() + "File Type : "
								+ fileoption + "                         " + System.lineSeparator()
								+ "======================================================================");

						try {

							InetAddress addr = InetAddress.getLocalHost();
							hostName = addr.getHostName();
						} catch (UnknownHostException e1) {

						}
						CardLayout card = (CardLayout) Cardlayout.getLayout();
						card.show(Cardlayout, "panel_2");
						model = (DefaultTreeModel) tree.getModel();
						root = new DefaultMutableTreeNode(hostName);
						model.setRoot(root);

						filepath = System.getProperty("user.home") + File.separator + "Library" + File.separator
								+ "Mail";

						DefaultMutableTreeNode node = new DefaultMutableTreeNode(filepath);

						root.add(node);
						kl = 0;

						try {

							readapple_mail(new File(filepath), node);
						} catch (Exception e) {
							e.printStackTrace();
						}

						if (kl == 0) {

							card.show(Cardlayout, "panel_1");
							JOptionPane.showMessageDialog(frame,
									"Apple Mail is not installed properly please try again after re-installing",
									messageboxtitle, JOptionPane.ERROR_MESSAGE,
									new ImageIcon(Main_Frame.class.getResource("/information.png")));
							if (!projectTitle.contains("Mail Migration")) {
								System.exit(0);
							} else {
								SwingUtilities.invokeLater(new Runnable() {

									public void run() {
										comboBox_FiletypeChooser.setSelectedItem("MBOX");
									}
								});
							}
						}
						kl = 0;
						Icon open = new ImageIcon(backup.class.getResource("/Open-folder-accept-icon.png"));
						Icon close = new ImageIcon(backup.class.getResource("/closed-folder-add-icon.png"));
						Icon Ram = new ImageIcon(backup.class.getResource("/leaf-icon.png"));
						DefaultCheckboxTreeCellRenderer render = (DefaultCheckboxTreeCellRenderer) tree
								.getCellRenderer();
						render.setClosedIcon(close);
						render.setOpenIcon(open);
						render.setLeafIcon(Ram);
						tree.expandRow(0);
						tree.expandAll();

					} else {
						JOptionPane.showMessageDialog(frame,
								"Apple Mail will not work in " + OS + " it reqiure mac OS ", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));
						SwingUtilities.invokeLater(new Runnable() {

							public void run() {
								comboBox_FiletypeChooser.setSelectedItem("MBOX");
							}
						});
					}

				} else if (fileoption.equalsIgnoreCase("Opera Mail")) {

					logger = logFile();
					logger.info("Start Time : " + calendertime + System.lineSeparator() + "File Type : " + fileoption
							+ "                         " + System.lineSeparator()
							+ "======================================================================");
					String str = null;
					str = System.getenv("APPDATA");

					if (OS.contains("windows")) {
						str = System.getenv("APPDATA").replace("Roaming", "Local") + File.separator + "Opera Mail"
								+ File.separator + "Opera Mail" + File.separator + "Mail" + File.separator + "store";
					} else {
						str = System.getProperty("user.home") + File.separator + "Library" + File.separator
								+ "Application Support" + File.separator + "Opera Mail" + File.separator + "mail";
					}

					if (new File(str).exists()) {
						thunderbirdpath = str;
						panel_mailfilter.setVisible(true);
						try {
							InetAddress addr = InetAddress.getLocalHost();
							hostName = addr.getHostName();
							CardLayout card1 = (CardLayout) inner_cardlayout.getLayout();
							card1.show(inner_cardlayout, "panel_selectfile");
							model = (DefaultTreeModel) tree.getModel();
							root = new DefaultMutableTreeNode(hostName);
							model.setRoot(root);
							DefaultMutableTreeNode node = new DefaultMutableTreeNode(
									"<html><b>" + new File(str).getAbsolutePath());
							root.add(node);
							readoperamail(new File(str), node);

							Icon open = new ImageIcon(Main_Frame.class.getResource("/Open-folder-accept-icon.png"));
							Icon close = new ImageIcon(Main_Frame.class.getResource("/closed-folder-add-icon.png"));
							Icon Ram = new ImageIcon(Main_Frame.class.getResource("/leaf-icon.png"));
							DefaultCheckboxTreeCellRenderer render = (DefaultCheckboxTreeCellRenderer) tree
									.getCellRenderer();
							render.setClosedIcon(close);
							render.setOpenIcon(open);
							render.setLeafIcon(Ram);

							tree.expandRow(0);
							tree.expandAll();
							CardLayout card = (CardLayout) Cardlayout.getLayout();
							card.show(Cardlayout, "panel_2");
							btnStopMigration.setVisible(false);
							btnStartTheMigration.setVisible(false);
						} catch (Exception e) {
						}

					} else {
						JOptionPane.showMessageDialog(frame, "Opera Mail not installed Please install it and try again",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));
						if (!projectTitle.contains("Mail Migration")) {
							System.exit(0);
						} else {
							SwingUtilities.invokeLater(new Runnable() {

								public void run() {
									comboBox_FiletypeChooser.setSelectedItem("MBOX");
								}
							});
						}
					}

				} else if (fileoption.equalsIgnoreCase("EML File (.eml)")
						|| fileoption.equalsIgnoreCase("EMLX File (.emlx)")
						|| fileoption.equalsIgnoreCase("OFT File (.oft)")
						|| fileoption.equalsIgnoreCase("Message File (.msg)")
						|| fileoption.equalsIgnoreCase("Maildir")) {
					comboBox.addItem("Original File Name");
					panel_mailfilter.setVisible(true);
					chckbx_Mail_Filter.setVisible(true);
					input = false;

					CardLayout card = (CardLayout) inner_cardlayout.getLayout();
					card.show(inner_cardlayout, "panel_selectfile");
					if (fileoption.equalsIgnoreCase("Message File (.msg)")) {
						label_7.setIcon(new ImageIcon(Main_Frame.class.getResource("/msg.png")));

					} else if (fileoption.equalsIgnoreCase("EMLX File (.emlx)")) {
						label_7.setIcon(new ImageIcon(Main_Frame.class.getResource("/emlx.png")));

					} else if (fileoption.equalsIgnoreCase("EML File (.eml)")) {
						label_7.setIcon(new ImageIcon(Main_Frame.class.getResource("/eml.png")));

					} else if (fileoption.equalsIgnoreCase("Maildir")) {
						label_7.setIcon(new ImageIcon(Main_Frame.class.getResource("/maildir.png")));

					} else if (fileoption.equalsIgnoreCase("OFT File (.oft)")) {
						label_7.setIcon(new ImageIcon(Main_Frame.class.getResource("/oft.png")));

					}

				}

				else if (fileoption.equalsIgnoreCase("OFFICE 365") || fileoption.equalsIgnoreCase("Yahoo Mail")
						|| fileoption.equalsIgnoreCase("Gmail") || fileoption.equalsIgnoreCase("Hotmail")
						|| fileoption.equalsIgnoreCase("Aol") || fileoption.equalsIgnoreCase("Zoho Mail")
						|| fileoption.equalsIgnoreCase("Yandex Mail") || fileoption.equalsIgnoreCase("Yandex Mail")

						|| fileoption.equalsIgnoreCase("GoDaddy email") || fileoption.equalsIgnoreCase("Icloud")) {
					lblthirdPartyPassword.setVisible(true);
					btn_SignIn.setToolTipText("Click here to Sign in to " + fileoption + ".");
					CardLayout card = (CardLayout) inner_cardlayout.getLayout();
					card.show(inner_cardlayout, "panel_loginpanel");
					panel_mailfilter.setVisible(true);
					lblTurnOffTwo.setText("<HTML><U>To access your " + fileoption
							+ " account , you'll need to generate and use an app password.</U></HTML>");
					chckbx_Mail_Filter.setVisible(true);
					lblEnabledImap.setVisible(false);
					panel_computername.setVisible(false);
					btn_next_pane1.setEnabled(false);
					input = true;
					lblthirdPartyPassword.setVisible(true);
					chckbxSetBackupSchedule.setVisible(true);
					label_backupschedul.setVisible(true);
					// chckbxSetBackupSchedule.setSelected(true);
					task_box.setVisible(false);

					panel_taskfilter.setVisible(false);

					if (fileoption.equalsIgnoreCase("Yahoo Mail")) {
						chckbxDeleteEmailFrom.setVisible(true);

						label_Free_up_Server_Space.setVisible(true);

						chckbxMaintainFolderHeirachy.setVisible(true);
						label_Maintain_Folder_Hierarchy.setVisible(true);
						chckbxAutoIncrementBackup.setVisible(true);
						lblSkip_Previously_Migrated_Items.setVisible(true);
						label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/yahoo-topbar.png")));

					} else if (fileoption.equalsIgnoreCase("OFFICE 365")) {
						backup backup = new backup(Main_Frame.this, demo);

						backup.setVisible(true);
						backup.setLocationRelativeTo(null);
						backup.label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/office.png")));
						backup.setResizable(false);
						dispose();
						task_box.setVisible(true);

						panel_taskfilter.setVisible(true);

					} else if (fileoption.equalsIgnoreCase("GoDaddy email")) {
						chckbxDeleteEmailFrom.setVisible(true);

						label_Free_up_Server_Space.setVisible(true);
						lblSkip_Previously_Migrated_Items.setVisible(true);
						chckbxAutoIncrementBackup.setVisible(true);
						chckbxMaintainFolderHeirachy.setVisible(true);
						lblPleaseMakeSure.setVisible(false);
						lblthirdPartyPassword.setVisible(false);
						lblEnabledImap.setVisible(false);
						lblTurnOffTwo.setVisible(false);
						label_Maintain_Folder_Hierarchy.setVisible(true);
						label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/godaddy.png")));

					} else if (fileoption.equalsIgnoreCase("Icloud")) {
						chckbxDeleteEmailFrom.setVisible(true);

						label_Free_up_Server_Space.setVisible(true);

						chckbxAutoIncrementBackup.setVisible(true);

						lblEnabledImap.setVisible(false);

						lblSkip_Previously_Migrated_Items.setVisible(true);
						chckbxMaintainFolderHeirachy.setVisible(true);
						label_Maintain_Folder_Hierarchy.setVisible(true);
						label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/icloud.png")));

					} else if (fileoption.equalsIgnoreCase("Yandex Mail")) {
						chckbxDeleteEmailFrom.setVisible(true);

						label_Free_up_Server_Space.setVisible(true);

						chckbxAutoIncrementBackup.setVisible(true);
						lblSkip_Previously_Migrated_Items.setVisible(true);
						chckbxMaintainFolderHeirachy.setVisible(true);
						label_Maintain_Folder_Hierarchy.setVisible(true);
						label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/yendex.png")));

					} else if (fileoption.equalsIgnoreCase("Aol")) {
						chckbxDeleteEmailFrom.setVisible(true);

						label_Free_up_Server_Space.setVisible(true);

						chckbxAutoIncrementBackup.setVisible(true);
						lblSkip_Previously_Migrated_Items.setVisible(true);
						chckbxMaintainFolderHeirachy.setVisible(true);
						label_Maintain_Folder_Hierarchy.setVisible(true);
						label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/aol-topbar.png")));

					} else if (fileoption.equalsIgnoreCase("Zoho Mail")) {
						chckbxDeleteEmailFrom.setVisible(true);

						label_Free_up_Server_Space.setVisible(true);

						chckbxAutoIncrementBackup.setVisible(true);
						lblSkip_Previously_Migrated_Items.setVisible(true);
						chckbxMaintainFolderHeirachy.setVisible(true);
						label_Maintain_Folder_Hierarchy.setVisible(true);
						label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/zoho.png")));

						lblEnabledImap.setVisible(true);
					} else if (fileoption.equalsIgnoreCase("Gmail")) {
						chckbxDeleteEmailFrom.setVisible(true);

						label_Free_up_Server_Space.setVisible(true);
						chckbxAutoIncrementBackup.setVisible(true);
						lblSkip_Previously_Migrated_Items.setVisible(true);
						btn_signout_p3.setVisible(true);
						label_Maintain_Folder_Hierarchy.setVisible(true);
						chckbxMaintainFolderHeirachy.setVisible(true);
						label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/gmail-topbar.png")));
						lblTurnOffTwo.setText("<HTML><U>* To access your " + fileoption
								+ " account , you'll need to generate and use an app password * Turn on less secure app</U></HTML>");

					} else if (fileoption.equalsIgnoreCase("Hotmail")) {

						label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/hotmail-topbar.png")));
						task_box.setVisible(true);
						chckbxSetBackupSchedule.setVisible(false);
						label_backupschedul.setVisible(false);
						chckbxSetBackupSchedule.setSelected(false);
						panel_taskfilter.setVisible(true);
					}

				} else if (fileoption.equalsIgnoreCase("Live Exchange")) {
					btn_SignIn.setToolTipText("Click here to Sign in to " + fileoption + ".");
					lbl_DomainName_computername.setText("Computer Name or IP Address");
					input = true;
					CardLayout card = (CardLayout) inner_cardlayout.getLayout();
					card.show(inner_cardlayout, "panel_loginpanel");
					panel_mailfilter.setVisible(true);
					lblthirdPartyPassword.setVisible(false);
					chckbx_Mail_Filter.setVisible(true);
					chckbxSetBackupSchedule.setVisible(false);
					label_backupschedul.setVisible(false);
					chckbxSetBackupSchedule.setSelected(false);
					lblthirdPartyPassword.setVisible(false);
					task_box.setVisible(true);

					panel_taskfilter.setVisible(true);
					lbl_DomainName_computername.setVisible(true);
					textField_Domainname_p2.setVisible(true);
					panel_computername.setVisible(true);

					btn_next_pane1.setEnabled(false);

					label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/exchange-topbar.png")));

					btn_signout_p3.setVisible(true);
					task_box.setVisible(true);

					panel_taskfilter.setVisible(true);
					lblPleaseMakeSure.setVisible(false);
					lblthirdPartyPassword.setVisible(false);
					lblEnabledImap.setVisible(false);
					lblTurnOffTwo.setVisible(false);

				} else if (fileoption.equalsIgnoreCase("Amazon WorkMail")) {

					lbl_DomainName_computername.setText("Amazon Domain Name");
					input = true;
					CardLayout card = (CardLayout) inner_cardlayout.getLayout();
					card.show(inner_cardlayout, "panel_loginpanel");
					panel_mailfilter.setVisible(true);
					lblthirdPartyPassword.setVisible(false);
					chckbx_Mail_Filter.setVisible(true);

					lblthirdPartyPassword.setVisible(false);
					// task_box.setVisible(true);
					label_Free_up_Server_Space.setVisible(true);
					chckbxDeleteEmailFrom.setVisible(true);
					lblSkip_Previously_Migrated_Items.setVisible(true);
					chckbxAutoIncrementBackup.setVisible(true);
					chckbxMaintainFolderHeirachy.setVisible(true);
					label_Maintain_Folder_Hierarchy.setVisible(true);

					lbl_DomainName_computername.setVisible(true);
					textField_Domainname_p2.setVisible(true);
					panel_computername.setVisible(true);

					btn_next_pane1.setEnabled(false);

					label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/amazon.png")));

					btn_signout_p3.setVisible(true);

					// panel_taskfilter.setVisible(true);
					lblPleaseMakeSure.setVisible(false);
					chckbxSetBackupSchedule.setVisible(true);
					label_backupschedul.setVisible(true);
					lblEnabledImap.setVisible(false);
					lblTurnOffTwo.setVisible(false);

					btn_signout_p3.setVisible(true);
				} else if (fileoption.equalsIgnoreCase("Hostgator email")) {
					lbl_DomainName_computername.setText("   Hostgator Host");
					input = true;
					CardLayout card = (CardLayout) inner_cardlayout.getLayout();
					card.show(inner_cardlayout, "panel_loginpanel");
					panel_mailfilter.setVisible(true);
					lblthirdPartyPassword.setVisible(false);
					chckbx_Mail_Filter.setVisible(true);
					task_box.setVisible(false);

					panel_taskfilter.setVisible(false);
					lbl_DomainName_computername.setVisible(true);
					textField_Domainname_p2.setVisible(true);
					panel_computername.setVisible(true);
					lblPleaseMakeSure.setVisible(false);
					label_Maintain_Folder_Hierarchy.setVisible(true);
					chckbxAutoIncrementBackup.setVisible(true);
					lblSkip_Previously_Migrated_Items.setVisible(true);
					lblEnabledImap.setVisible(false);
					lblTurnOffTwo.setVisible(false);
					btn_next_pane1.setEnabled(false);
					lblNewLabel_7.setVisible(true);
					tf_portNo_p2.setVisible(true);
					chckbxSetBackupSchedule.setVisible(true);
					label_backupschedul.setVisible(true);
					label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/hostgator.png")));

					btn_signout_p3.setVisible(true);

				} else if (fileoption.equalsIgnoreCase("IMAP")) {
					lbl_DomainName_computername.setText("      IMAP Host");
					input = true;
					CardLayout card = (CardLayout) inner_cardlayout.getLayout();
					card.show(inner_cardlayout, "panel_loginpanel");
					panel_mailfilter.setVisible(true);
					lblthirdPartyPassword.setVisible(false);
					chckbx_Mail_Filter.setVisible(true);
					task_box.setVisible(false);
					label_Maintain_Folder_Hierarchy.setVisible(true);
					chckbxAutoIncrementBackup.setVisible(true);
					lblSkip_Previously_Migrated_Items.setVisible(true);
					panel_taskfilter.setVisible(false);
					lbl_DomainName_computername.setVisible(true);
					textField_Domainname_p2.setVisible(true);
					panel_computername.setVisible(true);
					lblPleaseMakeSure.setVisible(false);
					chckbxSetBackupSchedule.setVisible(true);
					label_backupschedul.setVisible(true);
					lblEnabledImap.setVisible(false);
					lblTurnOffTwo.setVisible(false);
					btn_next_pane1.setEnabled(false);
					lblNewLabel_7.setVisible(true);
					tf_portNo_p2.setVisible(true);

					label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/imap-topbar.png")));

					btn_signout_p3.setVisible(true);

				}

			}
		});

		inner_cardlayout = new JPanel();
		inner_cardlayout.setBorder(new LineBorder(new Color(0, 0, 0)));
		inner_cardlayout.setLayout(new CardLayout(0, 0));

		panel_office365BR = new JPanel();
		panel_office365BR.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_office365BR.setBackground(Color.WHITE);
		inner_cardlayout.add(panel_office365BR, "panel_office365BR");

		btnI = new JButton("");
		btnI.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				btnI.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-oth-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				btnI.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-oth-btn.png")));
			}
		});
		btnI.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-oth-btn.png")));

		btnI.setRequestFocusEnabled(false);
		btnI.setRolloverEnabled(false);
		btnI.setOpaque(false);
		btnI.setFocusable(false);
		btnI.setFocusTraversalKeysEnabled(false);
		btnI.setFocusPainted(false);
		btnI.setDefaultCapable(false);
		btnI.setContentAreaFilled(false);
		btnI.setBorderPainted(false);
		btnI.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
			}
		});
		btnI.setToolTipText("Office 365 Backup Tool to Take Backup Office 365 Mailbox.");

		JButton btnI_1 = new JButton("");
		btnI_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
			}
		});
		btnI_1.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				btnI_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-oth-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				btnI_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-oth-btn.png")));
			}
		});
		btnI_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/info-oth-btn.png")));

		btnI_1.setRolloverEnabled(false);
		btnI_1.setRequestFocusEnabled(false);
		btnI_1.setOpaque(false);
		btnI_1.setFocusTraversalKeysEnabled(false);
		btnI_1.setFocusable(false);
		btnI_1.setFocusPainted(false);
		btnI_1.setDefaultCapable(false);
		btnI_1.setContentAreaFilled(false);
		btnI_1.setBorderPainted(false);
		btnI_1.setToolTipText("Restore Office 365 Mailbox from OST/PST Files.");

		btnbackup = new JButton("");
		btnbackup.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				btnbackup.setIcon(new ImageIcon(Main_Frame.class.getResource("/backup-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				btnbackup.setIcon(new ImageIcon(Main_Frame.class.getResource("/backup-btn.png")));
			}
		});
		btnbackup.setIcon(new ImageIcon(Main_Frame.class.getResource("/backup-btn.png")));
		btnbackup.setRolloverEnabled(false);
		btnbackup.setRequestFocusEnabled(false);
		btnbackup.setOpaque(false);
		btnbackup.setFocusTraversalKeysEnabled(false);
		btnbackup.setFocusable(false);
		btnbackup.setFocusPainted(false);
		btnbackup.setDefaultCapable(false);
		btnbackup.setContentAreaFilled(false);
		btnbackup.setBorderPainted(false);
		btnbackup.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				tools.setVisible(false);
				backup backup = new backup(Main_Frame.this, demo);

				backup.setVisible(true);
				backup.setLocationRelativeTo(null);
				backup.label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/office365-mian-topabr.png")));
				backup.lblTurnOffTwo_p3.setText(
						"<HTML><U>To access your 365 account , you'll need to generate and use an app password.</U></HTML>");
				if (fileoption.equalsIgnoreCase("Exchange Backup & Restore")) {
					backup.panel_computername.setVisible(true);
					backup.label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/exchange-topbar.png")));

				}
				backup.setResizable(false);
				dispose();
			}
		});

		JButton btnrestore = new JButton("");
		btnrestore.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				btnrestore.setIcon(new ImageIcon(Main_Frame.class.getResource("/restore-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				btnrestore.setIcon(new ImageIcon(Main_Frame.class.getResource("/restore-btn.png")));
			}
		});
		btnrestore.setIcon(new ImageIcon(Main_Frame.class.getResource("/restore-btn.png")));
		btnrestore.setRolloverEnabled(false);
		btnrestore.setRequestFocusEnabled(false);
		btnrestore.setOpaque(false);
		btnrestore.setFocusable(false);
		btnrestore.setFocusTraversalKeysEnabled(false);
		btnrestore.setFocusPainted(false);
		btnrestore.setDefaultCapable(false);
		btnrestore.setContentAreaFilled(false);
		btnrestore.setBorderPainted(false);
		btnrestore.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				restore restore = new restore(Main_Frame.this, demo, messageboxtitle);
				tools.setVisible(false);
				restore.setVisible(true);
				restore.setLocationRelativeTo(null);

				restore.setResizable(false);
				dispose();

			}
		});

		label_BR = new JLabel("");
		label_BR.setFont(new Font("Tahoma", Font.BOLD, 15));
		label_BR.setVisible(true);

		label_11 = new JLabel("");
		label_11.setOpaque(true);
		label_11.setBackground(new Color(0,0,0));

		lblThisOptionAllows = new JLabel(
				"This option allows us to take backup from office 365. It can take backup in 25+ format.\r\n");
		lblThisOptionAllows.setFont(new Font("Tahoma", Font.BOLD, 11));

		lblThisOptionAllows_1 = new JLabel(
				"This option allows the data to be read from PST/OST File and store it in office 365.");
		lblThisOptionAllows_1.setFont(new Font("Tahoma", Font.BOLD, 11));
		GroupLayout gl_panel_office365BR = new GroupLayout(panel_office365BR);
		gl_panel_office365BR.setHorizontalGroup(
			gl_panel_office365BR.createParallelGroup(Alignment.LEADING)
				.addComponent(label_11, GroupLayout.DEFAULT_SIZE, 657, Short.MAX_VALUE)
				.addGroup(gl_panel_office365BR.createSequentialGroup()
					.addGap(10)
					.addComponent(label_BR, GroupLayout.PREFERRED_SIZE, 301, GroupLayout.PREFERRED_SIZE))
				.addGroup(gl_panel_office365BR.createSequentialGroup()
					.addGap(62)
					.addComponent(btnbackup, GroupLayout.PREFERRED_SIZE, 210, GroupLayout.PREFERRED_SIZE)
					.addGap(311)
					.addComponent(btnI, GroupLayout.PREFERRED_SIZE, 62, GroupLayout.PREFERRED_SIZE))
				.addGroup(gl_panel_office365BR.createSequentialGroup()
					.addGap(62)
					.addComponent(lblThisOptionAllows, GroupLayout.PREFERRED_SIZE, 499, GroupLayout.PREFERRED_SIZE))
				.addGroup(gl_panel_office365BR.createSequentialGroup()
					.addGap(62)
					.addComponent(btnrestore, GroupLayout.PREFERRED_SIZE, 210, GroupLayout.PREFERRED_SIZE)
					.addGap(311)
					.addComponent(btnI_1, GroupLayout.PREFERRED_SIZE, 62, GroupLayout.PREFERRED_SIZE))
				.addGroup(gl_panel_office365BR.createSequentialGroup()
					.addGap(63)
					.addComponent(lblThisOptionAllows_1, GroupLayout.PREFERRED_SIZE, 477, GroupLayout.PREFERRED_SIZE))
		);
		gl_panel_office365BR.setVerticalGroup(
			gl_panel_office365BR.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_panel_office365BR.createSequentialGroup()
					.addGap(9)
					.addComponent(label_11, GroupLayout.PREFERRED_SIZE, 60, GroupLayout.PREFERRED_SIZE)
					.addGap(5)
					.addComponent(label_BR, GroupLayout.PREFERRED_SIZE, 32, GroupLayout.PREFERRED_SIZE)
					.addGap(10)
					.addGroup(gl_panel_office365BR.createParallelGroup(Alignment.LEADING)
						.addGroup(gl_panel_office365BR.createSequentialGroup()
							.addGap(1)
							.addComponent(btnbackup, GroupLayout.PREFERRED_SIZE, 37, GroupLayout.PREFERRED_SIZE))
						.addComponent(btnI, GroupLayout.PREFERRED_SIZE, 38, GroupLayout.PREFERRED_SIZE))
					.addGap(15)
					.addComponent(lblThisOptionAllows, GroupLayout.PREFERRED_SIZE, 38, GroupLayout.PREFERRED_SIZE)
					.addGap(28)
					.addGroup(gl_panel_office365BR.createParallelGroup(Alignment.LEADING)
						.addGroup(gl_panel_office365BR.createSequentialGroup()
							.addGap(1)
							.addComponent(btnrestore, GroupLayout.PREFERRED_SIZE, 37, GroupLayout.PREFERRED_SIZE))
						.addComponent(btnI_1, GroupLayout.PREFERRED_SIZE, 38, GroupLayout.PREFERRED_SIZE))
					.addGap(11)
					.addComponent(lblThisOptionAllows_1, GroupLayout.PREFERRED_SIZE, 16, GroupLayout.PREFERRED_SIZE))
		);
		panel_office365BR.setLayout(gl_panel_office365BR);

		panel_selectfile = new JPanel();
		panel_selectfile.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_selectfile.setBackground(Color.WHITE);
		inner_cardlayout.add(panel_selectfile, "panel_selectfile");

		rdbtn_MultipleFile = new JRadioButton("Select Folder");
		rdbtn_MultipleFile.setRolloverEnabled(false);
		rdbtn_MultipleFile.setRequestFocusEnabled(false);
		rdbtn_MultipleFile.setOpaque(false);
		rdbtn_MultipleFile.setFocusable(false);
		rdbtn_MultipleFile.setFocusPainted(false);
		rdbtn_MultipleFile.setContentAreaFilled(false);
		rdbtn_MultipleFile.setBackground(Color.WHITE);
		buttonGroup_file.add(rdbtn_MultipleFile);
		rdbtn_MultipleFile.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					btn_next_pane1.setEnabled(true);
				}
			}
		});

		rdbtn_MultipleFile.setFont(new Font("Bookman Old Style", Font.BOLD, 14));

		
		rdbtnSingleFile.setRolloverEnabled(false);
		rdbtnSingleFile.setRequestFocusEnabled(false);
		rdbtnSingleFile.setOpaque(false);
		rdbtnSingleFile.setFocusable(false);
		rdbtnSingleFile.setFocusPainted(false);
		rdbtnSingleFile.setContentAreaFilled(false);
		rdbtnSingleFile.setBackground(Color.WHITE);
		buttonGroup_file.add(rdbtnSingleFile);
		rdbtnSingleFile.setSelected(true);

		rdbtnSingleFile.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					btn_next_pane1.setEnabled(true);
				}
			}

		});
		rdbtnSingleFile.setFont(new Font("Bookman Old Style", Font.BOLD, 14));

		label_7 = new JLabel("");
		label_7.setBackground(new Color(0,0,0));
		label_7.setOpaque(true);

		lblthisOptionIs_1 = new JLabel("(Select This Option for Folders and then click on Next Button Below).");
		lblthisOptionIs_1.setFont(new Font("Bookman Old Style", Font.PLAIN, 12));
		lblthisOptionIs_1.setForeground(Color.RED);

		lblthisOptionIs = new JLabel(
				"(Select This option for single file or multiple file and then click on Next Button Below).");
		lblthisOptionIs.setFont(new Font("Bookman Old Style", Font.PLAIN, 12));
		lblthisOptionIs.setForeground(Color.RED);
		GroupLayout gl_panel_selectfile = new GroupLayout(panel_selectfile);
		gl_panel_selectfile.setHorizontalGroup(
			gl_panel_selectfile.createParallelGroup(Alignment.LEADING)
				.addComponent(label_7, GroupLayout.DEFAULT_SIZE, 657, Short.MAX_VALUE)
				.addGroup(gl_panel_selectfile.createSequentialGroup()
					.addGap(49)
					.addComponent(rdbtnSingleFile, GroupLayout.DEFAULT_SIZE, 304, Short.MAX_VALUE)
					.addGap(304))
				.addGroup(gl_panel_selectfile.createSequentialGroup()
					.addGap(69)
					.addComponent(lblthisOptionIs, GroupLayout.DEFAULT_SIZE, 537, Short.MAX_VALUE)
					.addGap(51))
				.addGroup(gl_panel_selectfile.createSequentialGroup()
					.addGap(49)
					.addGroup(gl_panel_selectfile.createParallelGroup(Alignment.LEADING)
						.addGroup(gl_panel_selectfile.createSequentialGroup()
							.addGap(20)
							.addComponent(lblthisOptionIs_1, GroupLayout.DEFAULT_SIZE, 557, Short.MAX_VALUE))
						.addGroup(gl_panel_selectfile.createSequentialGroup()
							.addComponent(rdbtn_MultipleFile, GroupLayout.DEFAULT_SIZE, 186, Short.MAX_VALUE)
							.addGap(391)))
					.addGap(31))
		);
		gl_panel_selectfile.setVerticalGroup(
			gl_panel_selectfile.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_panel_selectfile.createSequentialGroup()
					.addComponent(label_7, GroupLayout.PREFERRED_SIZE, 64, GroupLayout.PREFERRED_SIZE)
					.addGap(31)
					.addComponent(rdbtnSingleFile, GroupLayout.PREFERRED_SIZE, 25, GroupLayout.PREFERRED_SIZE)
					.addGap(4)
					.addComponent(lblthisOptionIs, GroupLayout.PREFERRED_SIZE, 23, GroupLayout.PREFERRED_SIZE)
					.addGap(53)
					.addGroup(gl_panel_selectfile.createParallelGroup(Alignment.LEADING)
						.addGroup(gl_panel_selectfile.createSequentialGroup()
							.addGap(21)
							.addComponent(lblthisOptionIs_1, GroupLayout.PREFERRED_SIZE, 23, GroupLayout.PREFERRED_SIZE))
						.addComponent(rdbtn_MultipleFile, GroupLayout.PREFERRED_SIZE, 23, GroupLayout.PREFERRED_SIZE)))
		);
		panel_selectfile.setLayout(gl_panel_selectfile);

		panel_loginpanel = new JPanel();

		panel_loginpanel.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_loginpanel.setBackground(Color.WHITE);
		inner_cardlayout.add(panel_loginpanel, "panel_loginpanel");
		panel_loginpanel.setLayout(null);

		JLabel lbl_username = new JLabel("USER NAME");
		lbl_username.setBounds(34, 80, 93, 27);
		panel_loginpanel.add(lbl_username);
		lbl_username.setFont(new Font("Tahoma", Font.BOLD, 12));

		JLabel lbl_password = new JLabel("PASSWORD");
		lbl_password.setBounds(34, 118, 93, 22);
		panel_loginpanel.add(lbl_password);
		lbl_password.setFont(new Font("Tahoma", Font.BOLD, 12));

		textField_username_p2 = new JTextField();
		textField_username_p2.setHorizontalAlignment(JTextField.CENTER);
		textField_username_p2.setComponentPopupMenu(menu);
		textField_username_p2.setFont(new Font("Tahoma", Font.BOLD, 15));
		textField_username_p2.setBounds(220, 79, 257, 27);
		panel_loginpanel.add(textField_username_p2);
		textField_username_p2.setColumns(10);

		passwordField_p2 = new JPasswordField();
		passwordField_p2.setHorizontalAlignment(JTextField.CENTER);
		passwordField_p2.setComponentPopupMenu(menu);
		passwordField_p2.setFont(new Font("Tahoma", Font.BOLD, 15));
		passwordField_p2.setBounds(220, 117, 257, 27);
		panel_loginpanel.add(passwordField_p2);

		lbl_connecting_p2 = new JLabel("");
		lbl_connecting_p2.setBounds(345, 290, 58, 30);
		panel_loginpanel.add(lbl_connecting_p2);
		lbl_connecting_p2.setIcon(new ImageIcon(Main_Frame.class.getResource("/loading.gif")));

		lbl_connecting_p2.setVisible(false);

		chckbxShowPassword_p2 = new JCheckBox("Show Password");
		chckbxShowPassword_p2.setRolloverEnabled(false);
		chckbxShowPassword_p2.setRequestFocusEnabled(false);
		chckbxShowPassword_p2.setOpaque(false);
		chckbxShowPassword_p2.setFocusable(false);
		chckbxShowPassword_p2.setFocusPainted(false);
		chckbxShowPassword_p2.setContentAreaFilled(false);
		chckbxShowPassword_p2.setBackground(Color.WHITE);
		chckbxShowPassword_p2.setBounds(483, 117, 148, 25);
		panel_loginpanel.add(chckbxShowPassword_p2);
		chckbxShowPassword_p2.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					passwordField_p2.setEchoChar((char) 0);
				}

				else {
					passwordField_p2.setEchoChar('●');
				}
			}
		});
		chckbxShowPassword_p2.setFont(new Font("Tahoma", Font.BOLD, 13));

		btn_SignIn = new JButton("");
		btn_SignIn.setRolloverEnabled(false);
		btn_SignIn.setRequestFocusEnabled(false);
		btn_SignIn.setOpaque(false);
		btn_SignIn.setFocusable(false);
		btn_SignIn.setFocusTraversalKeysEnabled(false);
		btn_SignIn.setFocusPainted(false);
		btn_SignIn.setDefaultCapable(false);
		btn_SignIn.setContentAreaFilled(false);
		btn_SignIn.setBorderPainted(false);
		btn_SignIn.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				btn_SignIn.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				btn_SignIn.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-btn.png")));
			}
		});

		btn_SignIn.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-btn.png")));
		btn_SignIn.setBounds(493, 215, 118, 68);

		panel_loginpanel.add(btn_SignIn);
		btn_SignIn.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {

				try {
					domain_p2 = textField_Domainname_p2.getText();
					domain_p2 = domain_p2.trim();
				} catch (Exception a) {
					domain_p2 = "";
				}
				try {
					username_p2 = textField_username_p2.getText();
					username_p2 = username_p2.trim();
				} catch (Exception a) {
					username_p2 = "";
				}
				try {
					password_p2 = new String(passwordField_p2.getPassword());
					password_p2 = password_p2.trim();
				} catch (Exception a) {
					password_p2 = "";
				}
				try {
					portmo = Integer.parseInt(tf_portNo_p2.getText());

				} catch (Exception a) {

				}
				comboBox_FiletypeChooser.setEnabled(false);
				if (username_p2.equalsIgnoreCase("") || password_p2.equalsIgnoreCase("")) {

					if (username_p2.equalsIgnoreCase("") && password_p2.equalsIgnoreCase("")) {
						JOptionPane.showMessageDialog(frame, "User name and Password fields cann't be empty.",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					} else if (username_p2.equalsIgnoreCase("")) {

						JOptionPane.showMessageDialog(frame, "User name field cann't be empty.", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					} else if (password_p2.equalsIgnoreCase("")) {

						JOptionPane.showMessageDialog(frame, "Password field cann't be empty.", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					}

					comboBox_FiletypeChooser.setEnabled(true);
					btn_Previous_pane2.setEnabled(true);
					btn_next_pane2.setEnabled(true);

					btn_signout_p3.setVisible(true);
					chckbxShowPassword_p2.setEnabled(true);
					btn_SignIn.setEnabled(true);
				} else if (fileoption.equalsIgnoreCase("Live Exchange") && domain_p2.equalsIgnoreCase("")) {

					JOptionPane.showMessageDialog(frame, "Computer Name or IP Address field can not be empty.",
							messageboxtitle, JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));
					comboBox_FiletypeChooser.setEnabled(true);
					btn_Previous_pane2.setEnabled(true);
					btn_next_pane2.setEnabled(true);

					chckbxShowPassword_p2.setEnabled(true);
					btn_SignIn.setEnabled(true);

				} else if (fileoption.equalsIgnoreCase("IMAP") && domain_p2.equalsIgnoreCase("")) {

					JOptionPane.showMessageDialog(frame, "IMAP Host field can not be empty.", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					comboBox_FiletypeChooser.setEnabled(true);
					btn_Previous_pane2.setEnabled(true);
					btn_next_pane2.setEnabled(true);

					chckbxShowPassword_p2.setEnabled(true);
					btn_SignIn.setEnabled(true);

				} else if (fileoption.equalsIgnoreCase("IMAP") && tf_portNo_p2.getText().isEmpty()) {

					JOptionPane.showMessageDialog(frame, "Port No field cann't be empty.", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					comboBox_FiletypeChooser.setEnabled(true);
					btn_Previous_pane2.setEnabled(true);
					btn_next_pane2.setEnabled(true);

					chckbxShowPassword_p2.setEnabled(true);
					btn_SignIn.setEnabled(true);

				} else if (!isValid(username_p2)) {

					JOptionPane.showMessageDialog(frame, "Please enter a valid username.", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					comboBox_FiletypeChooser.setEnabled(true);
					btn_Previous_pane2.setEnabled(true);
					btn_next_pane2.setEnabled(true);

					chckbxShowPassword_p2.setEnabled(true);
					btn_SignIn.setEnabled(true);

				} else {
					th = new Thread(new Runnable() {

						public void run() {

							List<String> ll = new ArrayList<String>();

							lbl_connecting_p2.setVisible(true);
							textField_Domainname_p2.setEnabled(false);
							passwordField_p2.setEnabled(false);
							tf_portNo_p2.setEnabled(false);
							textField_username_p2.setEnabled(false);
							btnSavingLog_1.setEnabled(false);
							btnTempPath.setEnabled(false);
							chckbxShowPassword_p2.setEnabled(false);
							btn_SignIn.setEnabled(false);

							try {

								if (clientforimap_input != null) {
									try {
										clientforimap_input.dispose();
									} catch (Exception e) {

									}

								}
								button_stop.setVisible(true);
								btn_Previous_pane2.setEnabled(false);
								btn_next_pane2.setEnabled(false);
								comboBox_FiletypeChooser.setEnabled(false);
								logger = logFile();
								cal = Calendar.getInstance();
								calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());
								logger.info("Start Time : " + calendertime + System.lineSeparator() + "File Type : "
										+ fileoption + "                         " + "Mailbox" + "    "
										+ textField_username_p2.getText() + System.lineSeparator()
										+ "======================================================================");

								String path = "";

								if (System.getProperty("os.name").toLowerCase().contains("windows")) {

									path = System.getenv("APPDATA") + File.separator + projectTitle + File.separator
											+ fileoption + ".ki";

								} else {

									path = System.getProperty("user.home") + File.separator + "Library" + File.separator
											+ "Application Support" + File.separator + projectTitle + File.separator
											+ fileoption + ".ki";
								}

								File file = new File(path);

								if (file.isFile()) {
									FileReader filereader = new FileReader(file);
									au.com.bytecode.opencsv.CSVReader csvReader = new au.com.bytecode.opencsv.CSVReader(
											filereader);
									String[] nextRecord;

									// we are going to read data line by line
									while ((nextRecord = csvReader.readNext()) != null) {
										for (String cell : nextRecord) {
											ll.add(cell);
										}
										System.out.println(ll);
									}

									csvReader.close();

								}

								System.out.println(ll.size());

								if (!ll.contains(username_p2)) {
									if (ll.size() >= accoutcount - 1) {
										JOptionPane.showMessageDialog(frame,
												"Maximum account allowed limit reached you can not add new account "
														+ System.lineSeparator()
														+ "please contact support for more information.",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));

										throw new NullPointerException();
									}
								}

								if (fileoption.equalsIgnoreCase("Yahoo Mail")) {

									connectiontoyahoo_input();

								} else if (fileoption.equalsIgnoreCase("Gmail")) {
									connectiontogmail_input();

								} else if (fileoption.equalsIgnoreCase("Aol")) {
									connectiontoaol_input();

								} else if (fileoption.equalsIgnoreCase("Icloud")) {
									connectiontoicloud_input();

								} else if (fileoption.equalsIgnoreCase("GoDaddy email")) {
									connectiontoGoDaddy_input();

								} else if (fileoption.equalsIgnoreCase("Hostgator email")) {
									connectiontoHostgator_input();

								} else if (fileoption.equalsIgnoreCase("Amazon WorkMail")) {
									connectiontoinaws_input();

								} else if (fileoption.equalsIgnoreCase("IMAP")) {
									connectiontoimap_input();

								} else if (fileoption.equalsIgnoreCase("Hotmail")) {
									conntiontohotmail_input();

								} else if (fileoption.equalsIgnoreCase("Yandex Mail")) {
									connectiontonYandex_input();

								} else if (fileoption.equalsIgnoreCase("Zoho Mail")) {
									connectiontozoho_input();

								}

								else if (fileoption.equalsIgnoreCase("OFFICE 365")) {
									conntiontooffice365_input();

								} else if (fileoption.equalsIgnoreCase("Live Exchange")) {
									connectionwithexchangeserver_input();

								}

								if (!ll.contains(username_p2)) {

									if (file.isFile())
										file.delete();
									CSVWriter writer = new CSVWriter(new FileWriter(file.getAbsolutePath()));
									String line1[] = { username_p2 };

									writer.writeNext(line1);

									for (int i = 0; i < ll.size(); i++) {
										String line[] = { ll.get(i) };
										writer.writeNext(line);

									}
									writer.close();

								}

								System.out.println(ll.size());
								chckbxSaveInSame.setVisible(false);
								lblSaveintheSameFolder.setVisible(false);
								logpath = textField_hi.getText();
								temppath = textField_1.getText();
								lblNewLabel_5.setIcon(new ImageIcon(Main_Frame.class.getResource("/topbar.png")));
								if (nextTime != null) {
									nextsign();

								} else {
									sign();

								}
								tools.setVisible(false);
								button_stop.setVisible(false);
								comboBox_FiletypeChooser.setEnabled(true);
								scrollPane_fortable_p2.setVisible(false);
								innerCardlayout.setVisible(false);
								btnAttachment.setVisible(false);
								btnViewer.setVisible(false);
								lbl_Date.setVisible(false);
								lbl_Email.setVisible(false);
								lbl_subject.setVisible(false);
								cal = Calendar.getInstance();

								if (tree.getHeight() < 3) {
									CardLayout card = (CardLayout) Cardlayout.getLayout();
									card.show(Cardlayout, "panel_1");
								}

								calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());

								chckbxShowPassword_p2.setEnabled(true);

							} catch (Error e) {
								logger.warning("Error : " + e.getMessage() + System.lineSeparator());
							} catch (Exception e) {

								logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
								if (fileoption.equalsIgnoreCase("Gmail")) {
									if (e.getMessage().equalsIgnoreCase(
											"AE_1_2_0002 NO [AUTHENTICATIONFAILED] Invalid credentials (Failure)")) {
										JOptionPane.showMessageDialog(frame,
												"Connection Not Estalished with Gmail please check your Credantial Otherwise allow 3rd party app to acess your account.",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else if (e.getMessage().contains("Application-specific password required:")) {
										JOptionPane.showMessageDialog(frame, "Application specific password required.",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else {
										JOptionPane.showMessageDialog(frame, "Connection not established.",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									}
								} else if (fileoption.equalsIgnoreCase("Yahoo Mail")) {
									if (e.getMessage().equalsIgnoreCase(
											"AE_1_2_0002 NO [AUTHORIZATIONFAILED] LOGIN Invalid credentials")) {
										JOptionPane.showMessageDialog(frame,
												"Connection Not Estalished with Yahoo Mail please check your Credantial Otherwise allow 3rd party app to acess your accoun.t",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else if (e.getMessage().contains(" Application-specific password required: ")) {
										JOptionPane.showMessageDialog(frame, "Application specific password required.",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else {
										JOptionPane.showMessageDialog(frame, "Connection not established.",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									}
								} else if (e.getMessage().contains("Application-specific password required: ")) {
									JOptionPane.showMessageDialog(frame, "Application specific password required",
											messageboxtitle, JOptionPane.ERROR_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/information.png")));
								} else {
									JOptionPane.showMessageDialog(frame, "Connection not established.", messageboxtitle,
											JOptionPane.ERROR_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/information.png")));
								}

								textField_Domainname_p2.setEnabled(true);
								passwordField_p2.setEnabled(true);
								tf_portNo_p2.setEnabled(true);
								btnSavingLog.setEnabled(true);
								btnTempPath.setEnabled(true);
								lblLiveChat.setVisible(true);
								textField_username_p2.setEnabled(true);
								return;

							} finally {
								lbl_connecting_p2.setVisible(false);
								btn_Previous_pane2.setEnabled(true);
								btn_next_pane2.setEnabled(true);
								btnSavingLog_1.setEnabled(true);
								btnTempPath.setEnabled(true);
								passwordField_p2.setEditable(true);
								chckbxShowPassword_p2.setEnabled(true);
								passwordField_p2.setEnabled(true);
								btn_SignIn.setEnabled(true);
								stop_tree = false;
								comboBox_FiletypeChooser.setEnabled(true);
								if (nextTime != null) {
									btn_next_pane2.doClick();
								}
							}

						}
					});
					th.start();

				}
			}
		});
		btn_SignIn.setFont(new Font("Tahoma", Font.BOLD, 14));

		panel_computername = new JPanel();
		panel_computername.setBackground(Color.WHITE);
		panel_computername.setBounds(10, 165, 417, 39);
		panel_loginpanel.add(panel_computername);
		panel_computername.setLayout(null);

		lbl_DomainName_computername = new JLabel("");
		lbl_DomainName_computername.setBounds(0, 0, 193, 27);
		panel_computername.add(lbl_DomainName_computername);
		lbl_DomainName_computername.setFont(new Font("Tahoma", Font.BOLD, 12));

		textField_Domainname_p2 = new JTextField();
		textField_Domainname_p2.setHorizontalAlignment(JTextField.CENTER);
		textField_Domainname_p2.setComponentPopupMenu(menu);
		textField_Domainname_p2.setFont(new Font("Tahoma", Font.BOLD, 15));
		textField_Domainname_p2.setBounds(210, 0, 257, 27);
		panel_computername.add(textField_Domainname_p2);
		textField_Domainname_p2.setVisible(false);
		textField_Domainname_p2.setColumns(10);

		btn_ques = new JButton("");
		btn_ques.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				btn_ques.setIcon(new ImageIcon(Main_Frame.class.getResource("/about-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				btn_ques.setIcon(new ImageIcon(Main_Frame.class.getResource("/about-btn.png")));
			}
		});

		btn_ques.setIcon(new ImageIcon(Main_Frame.class.getResource("/about-btn.png")));

		btn_ques.setCursor(cursor);
		btn_ques.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				openBrowser(All_Data.serverlogin);
			}
		});
		btn_ques.setRolloverEnabled(false);
		btn_ques.setRequestFocusEnabled(false);
		btn_ques.setOpaque(false);

		btn_ques.setFocusable(false);
		btn_ques.setFocusTraversalKeysEnabled(false);
		btn_ques.setFocusPainted(false);
		btn_ques.setDefaultCapable(false);
		btn_ques.setContentAreaFilled(false);
		btn_ques.setBorderPainted(false);
		btn_ques.setBounds(468, 0, 55, 33);
		panel_computername.add(btn_ques);

		label_8 = new JLabel("");
		label_8.setForeground(Color.RED);
		label_8.setBounds(0, 0, 655, 68);
		panel_loginpanel.add(label_8);

		tf_portNo_p2 = new JTextField();
		tf_portNo_p2.setHorizontalAlignment(JTextField.CENTER);
		tf_portNo_p2.setComponentPopupMenu(menu);
		tf_portNo_p2.setBounds(220, 215, 257, 27);
		panel_loginpanel.add(tf_portNo_p2);
		tf_portNo_p2.setVisible(false);
		tf_portNo_p2.setText(Integer.toString(993));
		tf_portNo_p2.setColumns(10);

		lblNewLabel_7 = new JLabel("PORT No\r\n");
		lblNewLabel_7.setFont(new Font("Tahoma", Font.BOLD, 14));
		lblNewLabel_7.setBounds(34, 215, 176, 22);
		lblNewLabel_7.setVisible(false);
		panel_loginpanel.add(lblNewLabel_7);

		lblLiveChat = new JLabel("More Help");
		lblLiveChat.setForeground(Color.RED);

		lblLiveChat.setCursor(cursor);
		lblLiveChat.setFont(new Font("Tahoma", Font.PLAIN, 14));

		lblLiveChat.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {

				openBrowser("http://messenger.providesupport.com/messenger/0pi295uz3ga080c7lxqxxuaoxr.html");

			}
		});

		lblLiveChat.setBounds(483, 70, 76, 27);
		panel_loginpanel.add(lblLiveChat);

		lblPleaseMakeSure = new JLabel("Please  Click on The Link");
		lblPleaseMakeSure.setForeground(Color.BLUE);
		lblPleaseMakeSure.setBounds(10, 259, 140, 14);
		panel_loginpanel.add(lblPleaseMakeSure);

		lblEnabledImap = new JLabel("<HTML><U>To Enabled Imap</U></HTML>");
		lblEnabledImap.setForeground(Color.RED);

		lblEnabledImap.setCursor(cursor);
		lblEnabledImap.setFont(new Font("Tahoma", Font.PLAIN, 11));
		lblEnabledImap.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				openbrowserenableimap(fileoption);
			}
		});
		lblEnabledImap.setBounds(173, 296, 125, 14);
		panel_loginpanel.add(lblEnabledImap);

		lblTurnOffTwo = new JLabel("");
		lblTurnOffTwo.setForeground(Color.RED);

		lblTurnOffTwo.setCursor(cursor);
		lblTurnOffTwo.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				openbrowserturntwostepoff(fileoption);
			}
		});
		lblTurnOffTwo.setFont(new Font("Tahoma", Font.PLAIN, 11));

		lblTurnOffTwo.setBounds(173, 259, 407, 29);
		panel_loginpanel.add(lblTurnOffTwo);

		lblthirdPartyPassword = new JLabel("(Use third party App Password)");
		lblthirdPartyPassword.setForeground(Color.RED);
		lblthirdPartyPassword.setBounds(10, 139, 356, 27);
		panel_loginpanel.add(lblthirdPartyPassword);

		JLabel lblemailAddress = new JLabel("(Email Address)");
		lblemailAddress.setForeground(Color.RED);
		lblemailAddress.setBounds(34, 98, 176, 22);
		panel_loginpanel.add(lblemailAddress);

		panel_Loading = new JPanel();
		panel_Loading.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_Loading.setBackground(Color.WHITE);
		inner_cardlayout.add(panel_Loading, "panel_Loading");
		panel_Loading.setLayout(null);

		lblNewLabel_2 = new JLabel("Login Successful");
		lblNewLabel_2.setForeground(SystemColor.textHighlight);
		lblNewLabel_2.setFont(new Font("Tahoma", Font.BOLD, 18));
		lblNewLabel_2.setBounds(218, 13, 318, 44);
		panel_Loading.add(lblNewLabel_2);

		lblReadingFoldersPlease = new JLabel("");
		lblReadingFoldersPlease.setText("Reading Folders Please wait......");
		lblReadingFoldersPlease.setFont(new Font("Tahoma", Font.BOLD, 17));
		lblReadingFoldersPlease.setBounds(25, 83, 279, 44);
		panel_Loading.add(lblReadingFoldersPlease);

		lblNewLabel_3 = new JLabel("");
		lblNewLabel_3.setFont(new Font("Tahoma", Font.BOLD, 10));
		lblNewLabel_3.setBounds(35, 124, 604, 41);
		panel_Loading.add(lblNewLabel_3);

		lblNewLabel_4 = new JLabel("");
		lblNewLabel_4.setIcon(new ImageIcon(Main_Frame.class.getResource("/loading.gif")));
		lblNewLabel_4.setBounds(279, 226, 73, 44);
		panel_Loading.add(lblNewLabel_4);

		button_stop = new JButton("");

		button_stop.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				button_stop.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				button_stop.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-btn.png")));
			}
		});

		button_stop.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-btn.png")));
		button_stop.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				String warn = "Do you want to stop the process?";
				int ans = JOptionPane.showConfirmDialog(Main_Frame.this, warn, messageboxtitle,
						JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {
					stop_tree = true;
				}

			}
		});
		button_stop.setRolloverEnabled(false);
		button_stop.setRequestFocusEnabled(false);
		button_stop.setOpaque(false);

		button_stop.setFocusable(false);
		button_stop.setFocusTraversalKeysEnabled(false);
		button_stop.setFocusPainted(false);
		button_stop.setDefaultCapable(false);
		button_stop.setContentAreaFilled(false);
		button_stop.setBorderPainted(false);
		button_stop.setBounds(246, 281, 142, 38);
		panel_Loading.add(button_stop);

		btnSavingLog = new JButton("");

		JPanel panel_8 = new JPanel();
		panel_8.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"),
				"Select Path For Log and Temporary File(s)", TitledBorder.LEADING, TitledBorder.TOP, null,
				new Color(0, 0, 0)));
		panel_8.setBackground(Color.WHITE);

		textField_1 = new JTextField();
		textField_1.setBackground(Color.WHITE);
		textField_1.setEditable(false);
		textField_1.setText(System.getProperty("java.io.tmpdir"));
		textField_1.setColumns(10);

		textField_hi = new JTextField();
		textField_hi.setBackground(Color.WHITE);
		textField_hi.setText(System.getProperty("user.home") + File.separator + "Documents");
		textField_hi.setEditable(false);
		textField_hi.setColumns(10);

		btnSavingLog_1 = new JButton();
		btnSavingLog_1.setRolloverEnabled(false);
		btnSavingLog_1.setRequestFocusEnabled(false);
		btnSavingLog_1.setOpaque(false);
		btnSavingLog_1.setFocusable(false);
		btnSavingLog_1.setFocusTraversalKeysEnabled(false);
		btnSavingLog_1.setFocusPainted(false);
		btnSavingLog_1.setDefaultCapable(false);
		btnSavingLog_1.setContentAreaFilled(false);
		btnSavingLog_1.setBorderPainted(false);
		btnSavingLog_1.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				btnSavingLog_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/log-path-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				btnSavingLog_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/log-path-btn.png")));
			}
		});

		btnSavingLog_1.setIcon(new ImageIcon(Main_Frame.class.getResource("/log-path-btn.png")));
		btnSavingLog_1.setToolTipText("Click here to Set the Log Path. ");

		btnTempPath = new JButton("");
		btnTempPath.setToolTipText("Click here to Set the Temp Path. ");
		btnTempPath.addMouseListener(new MouseAdapter() {

			public void mouseEntered(MouseEvent arg0) {
				btnTempPath.setIcon(new ImageIcon(Main_Frame.class.getResource("/temp-path-hvr-btn.png")));
			}

			public void mouseExited(MouseEvent e) {
				btnTempPath.setIcon(new ImageIcon(Main_Frame.class.getResource("/temp-path-btn.png")));
			}
		});

		btnTempPath.setIcon(new ImageIcon(Main_Frame.class.getResource("/temp-path-btn.png")));
		btnTempPath.setRolloverEnabled(false);
		btnTempPath.setRequestFocusEnabled(false);
		btnTempPath.setFocusable(false);
		btnTempPath.setFocusTraversalKeysEnabled(false);
		btnTempPath.setFocusPainted(false);
		btnTempPath.setDefaultCapable(false);
		btnTempPath.setContentAreaFilled(false);
		btnTempPath.setBorderPainted(false);
		btnTempPath.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser jFileChooser = new JFileChooser(System.getProperty("java.io.tmpdir"));

				jFileChooser.setBackground(Color.WHITE);

				jFileChooser.setAcceptAllFileFilterUsed(false);

				jFileChooser.setMultiSelectionEnabled(true);

				jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

				jFileChooser.showOpenDialog(Main_Frame.this);

				File file = jFileChooser.getSelectedFile();
				if (!(file == null)) {

					String destination = file.getAbsolutePath();

					textField_1.setText(destination);
				}

			}
		});

		JLabel lblSelectEmailClientsfile = new JLabel("");
		lblSelectEmailClientsfile.setForeground(Color.BLUE);
		lblSelectEmailClientsfile.setFont(new Font("Sitka Small", Font.BOLD, 18));
		projectTitle = All_Data.messageboxtitle;
		if (projectTitle.contains("Aryson Email Migration Tool") || projectTitle.contains("Cigati Email Migrator")
				|| projectTitle.contains("DRS Email Migration Tool")) {

			lblSelectEmailClientsfile.setIcon(new ImageIcon(Main_Frame.class.getResource("/strip.png")));
		}
		
		JPanel panel_12 = new JPanel();
		panel_12.setBackground(new Color( 156,	161, 172));
		
				label_6 = new JLabel("");
				label_6.setIcon(new ImageIcon(Main_Frame.class.getResource("/sidebar.png")));
				GroupLayout gl_panel_12 = new GroupLayout(panel_12);
				gl_panel_12.setHorizontalGroup(
					gl_panel_12.createParallelGroup(Alignment.LEADING)
						.addComponent(label_6)
				);
				gl_panel_12.setVerticalGroup(
					gl_panel_12.createParallelGroup(Alignment.LEADING)
						.addComponent(label_6, GroupLayout.PREFERRED_SIZE, 601, Short.MAX_VALUE)
				);
				panel_12.setLayout(gl_panel_12);
				
				JPanel panel_13 = new JPanel();
				panel_13.setBackground(new Color(0,	0 , 0));
						
								lblNewLabel_12 = new JLabel("<html><h2><i><u>>Explore more products.</u></i></h2></html>");
								lblNewLabel_12.setForeground(new Color(255, 255, 255));
								lblNewLabel_12.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
								
										lblNewLabel_12.addMouseListener(new MouseAdapter() {
								
											public void mouseClicked(MouseEvent e) {
								
												try {
													Desktop.getDesktop().browse(new URI(All_Data.exploremoreproducts));
												} catch (URISyntaxException | IOException ex) {
													// It looks like there's a problem
												}
								
											}
										});
										lblNewLabel_12.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
										lblNewLabel_12.setFont(new Font("Segoe UI Semibold", Font.BOLD, 16));
				
						btn_next_pane1 = new JButton("");
						btn_next_pane1.setRolloverEnabled(false);
						btn_next_pane1.addMouseListener(new MouseAdapter() {

							public void mouseEntered(MouseEvent arg0) {
								btn_next_pane1.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-hvr-btn.png")));
							}

							public void mouseExited(MouseEvent e) {
								btn_next_pane1.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-btn.png")));
							}
						});
						
								btn_next_pane1.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-btn.png")));
								
										btn_next_pane1.setRequestFocusEnabled(false);
										btn_next_pane1.setOpaque(false);
										btn_next_pane1.setFocusable(false);
										btn_next_pane1.setFocusPainted(false);
										btn_next_pane1.setToolTipText("Click here to Go to next panel. ");
										btn_next_pane1.setFocusTraversalKeysEnabled(false);
										btn_next_pane1.setDefaultCapable(false);
										btn_next_pane1.setContentAreaFilled(false);
										btn_next_pane1.setBorderPainted(false);
										btn_next_pane1.setFont(new Font("Tahoma", Font.BOLD, 12));
										GroupLayout gl_panel_1 = new GroupLayout(panel_1);
										gl_panel_1.setHorizontalGroup(
											gl_panel_1.createParallelGroup(Alignment.LEADING)
												.addGroup(gl_panel_1.createSequentialGroup()
													.addGap(396)
													.addComponent(panel_13, GroupLayout.DEFAULT_SIZE, 679, Short.MAX_VALUE))
												.addGroup(gl_panel_1.createSequentialGroup()
													.addGap(408)
													.addComponent(comboBox_FiletypeChooser, 0, 650, Short.MAX_VALUE)
													.addGap(17))
												.addGroup(gl_panel_1.createSequentialGroup()
													.addGap(408)
													.addComponent(lblSelectEmailClientsfile, GroupLayout.PREFERRED_SIZE, 657, GroupLayout.PREFERRED_SIZE))
												.addComponent(panel_12, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
												.addGroup(Alignment.TRAILING, gl_panel_1.createSequentialGroup()
													.addGap(408)
													.addGroup(gl_panel_1.createParallelGroup(Alignment.TRAILING)
														.addComponent(panel_8, Alignment.LEADING, GroupLayout.DEFAULT_SIZE, 650, Short.MAX_VALUE)
														.addComponent(inner_cardlayout, GroupLayout.PREFERRED_SIZE, 650, Short.MAX_VALUE))
													.addGap(17))
										);
										gl_panel_1.setVerticalGroup(
											gl_panel_1.createParallelGroup(Alignment.LEADING)
												.addComponent(panel_12, GroupLayout.PREFERRED_SIZE, 601, Short.MAX_VALUE)
												.addGroup(gl_panel_1.createSequentialGroup()
													.addGroup(gl_panel_1.createParallelGroup(Alignment.LEADING)
														.addGroup(Alignment.TRAILING, gl_panel_1.createSequentialGroup()
															.addGap(77)
															.addComponent(inner_cardlayout, GroupLayout.PREFERRED_SIZE, 320, GroupLayout.PREFERRED_SIZE)
															.addGap(11))
														.addGroup(gl_panel_1.createSequentialGroup()
															.addGap(39)
															.addComponent(comboBox_FiletypeChooser, GroupLayout.PREFERRED_SIZE, 39, GroupLayout.PREFERRED_SIZE)
															.addGap(330))
														.addGroup(gl_panel_1.createSequentialGroup()
															.addGap(11)
															.addComponent(lblSelectEmailClientsfile, GroupLayout.PREFERRED_SIZE, 25, GroupLayout.PREFERRED_SIZE)
															.addGap(372)))
													.addPreferredGap(ComponentPlacement.RELATED)
													.addComponent(panel_8, GroupLayout.PREFERRED_SIZE, 127, GroupLayout.PREFERRED_SIZE)
													.addPreferredGap(ComponentPlacement.RELATED, 7, Short.MAX_VALUE)
													.addComponent(panel_13, GroupLayout.PREFERRED_SIZE, 59, GroupLayout.PREFERRED_SIZE))
										);
										GroupLayout gl_panel_13 = new GroupLayout(panel_13);
										gl_panel_13.setHorizontalGroup(
											gl_panel_13.createParallelGroup(Alignment.LEADING)
												.addGroup(gl_panel_13.createSequentialGroup()
													.addGap(40)
													.addComponent(lblNewLabel_12, GroupLayout.DEFAULT_SIZE, 323, Short.MAX_VALUE)
													.addGap(153)
													.addComponent(btn_next_pane1)
													.addGap(21))
										);
										gl_panel_13.setVerticalGroup(
											gl_panel_13.createParallelGroup(Alignment.LEADING)
												.addGroup(gl_panel_13.createSequentialGroup()
													.addGap(11)
													.addGroup(gl_panel_13.createParallelGroup(Alignment.LEADING)
														.addComponent(lblNewLabel_12, GroupLayout.PREFERRED_SIZE, 24, GroupLayout.PREFERRED_SIZE)
														.addComponent(btn_next_pane1)))
										);
										panel_13.setLayout(gl_panel_13);
										GroupLayout gl_panel_8 = new GroupLayout(panel_8);
										gl_panel_8.setHorizontalGroup(
											gl_panel_8.createParallelGroup(Alignment.LEADING)
												.addGroup(gl_panel_8.createSequentialGroup()
													.addGap(4)
													.addGroup(gl_panel_8.createParallelGroup(Alignment.LEADING)
														.addComponent(textField_hi, Alignment.TRAILING, GroupLayout.DEFAULT_SIZE, 503, Short.MAX_VALUE)
														.addComponent(textField_1, Alignment.TRAILING, GroupLayout.DEFAULT_SIZE, 493, Short.MAX_VALUE))
													.addGap(18)
													.addGroup(gl_panel_8.createParallelGroup(Alignment.TRAILING, false)
														.addGroup(gl_panel_8.createSequentialGroup()
															.addComponent(btnTempPath, GroupLayout.PREFERRED_SIZE, 110, GroupLayout.PREFERRED_SIZE)
															.addGap(11))
														.addGroup(gl_panel_8.createSequentialGroup()
															.addComponent(btnSavingLog_1, 0, 0, Short.MAX_VALUE)
															.addContainerGap())))
										);
										gl_panel_8.setVerticalGroup(
											gl_panel_8.createParallelGroup(Alignment.LEADING)
												.addGroup(gl_panel_8.createSequentialGroup()
													.addGap(11)
													.addGroup(gl_panel_8.createParallelGroup(Alignment.LEADING)
														.addComponent(textField_hi, GroupLayout.PREFERRED_SIZE, 30, GroupLayout.PREFERRED_SIZE)
														.addComponent(btnSavingLog_1, GroupLayout.PREFERRED_SIZE, 30, GroupLayout.PREFERRED_SIZE))
													.addGap(16)
													.addGroup(gl_panel_8.createParallelGroup(Alignment.LEADING)
														.addComponent(textField_1, GroupLayout.PREFERRED_SIZE, 30, GroupLayout.PREFERRED_SIZE)
														.addComponent(btnTempPath, GroupLayout.PREFERRED_SIZE, 30, GroupLayout.PREFERRED_SIZE)))
										);
										panel_8.setLayout(gl_panel_8);
										panel_1.setLayout(gl_panel_1);
										btn_next_pane1.addActionListener(new ActionListener() {
											public void actionPerformed(ActionEvent e) {

												try {

													tools.setVisible(false);
													logpath = textField_hi.getText();
													temppath = textField_1.getText();
													scrollPane_fortable_p2.setVisible(true);
													innerCardlayout.setVisible(true);
													btnAttachment.setVisible(true);
													btnViewer.setVisible(true);
													lbl_Date.setVisible(true);
													lbl_Email.setVisible(true);
													lbl_subject.setVisible(true);
													stop_tree = false;
													cal = Calendar.getInstance();

													calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());

													if (rdbtn_MultipleFile.isSelected() || rdbtnSingleFile.isSelected()) {

														main_multiplefile main_m = new main_multiplefile(Main_Frame.this, demo, messageboxtitle);

														main_m.setVisible(true);
														main_m.setLocationRelativeTo(null);

														main_m.setResizable(true);
														dispose();

													}

												} catch (Exception e1) {
													e1.printStackTrace();
//					logger.warning(e1.getMessage() + System.lineSeparator());

												}

											}

										});
		btnSavingLog_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				JFileChooser jFileChooser = new JFileChooser();

				jFileChooser.setBackground(Color.WHITE);

				jFileChooser.setAcceptAllFileFilterUsed(false);

				jFileChooser.setMultiSelectionEnabled(true);

				jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

				jFileChooser.showOpenDialog(Main_Frame.this);

				File file = jFileChooser.getSelectedFile();
				if (!(file == null)) {
					String destination = file.getAbsolutePath();

					textField_hi.setText(destination);
				}

			}
		});

		panel_2 = new JPanel();
		panel_2.setBackground(Color.WHITE);
		Cardlayout.add(panel_2, "panel_2");
		panel_2.setLayout(null);

		JPanel panel_2_2 = new JPanel();
		panel_2_2.setBounds(0, 0, 1075, 620);
		panel_2.add(panel_2_2);
		panel_2_2.setLayout(new CardLayout(0, 0));

		panel_2_2_1 = new JPanel();
		panel_2_2_1.setBackground(Color.WHITE);
		panel_2_2.add(panel_2_2_1, "panel_2_2_1");
		panel_2_2_1.setLayout(null);

		label_13 = new JLabel("");
		label_13.setVisible(false);
		label_13.setIcon(new ImageIcon(Main_Frame.class.getResource("/progress-bar.gif")));
		label_13.setBounds(9, 562, 671, 30);
		panel_2_2_1.add(label_13);

		JScrollPane scrollPane_fortree_p2 = new JScrollPane();
		scrollPane_fortree_p2.setBounds(0, 0, 287, 550);
		panel_2_2_1.add(scrollPane_fortree_p2);

		tree = new CheckboxTree();
		tree.setShowsRootHandles(true);

		scrollPane_fortree_p2.setViewportView(tree);

		tree.addMouseListener(new MouseAdapter() {

			public void mouseClicked(MouseEvent arg0) {
				doubleclickcount++;
				th = new Thread(new Runnable() {

					public void run() {

						try {

							if (input == false) {

								btn_Cancel_pane2.setVisible(true);
								btn_Cancel_pane2.setEnabled(true);
								lblPleaseWatTable.setVisible(true);
								label_13.setVisible(true);
								lblNew_setemail.setText("");
								btn_next_pane2.setEnabled(false);
								btn_next_pane2.setVisible(false);

								lblNew_setsubject.setText("");

								label_date.setText("");
								lblTotalMessageCount.setVisible(true);
								lblTotalMessageCount.setText("Total Message Count : ");
								Stoppreview = false;
								btnAttachment.setEnabled(false);

								btnViewer.setEnabled(false);

								btn_Previous_pane2.setEnabled(false);

								listmail.clear();
								listmapi.clear();
								listPSTOSTgemesingo.clear();
								listExchangemesingo.clear();

							}
							CardLayout card = (CardLayout) innerCardlayout.getLayout();
							card.show(innerCardlayout, "viewer");
							editorPane.setText("");
							if (arg0.getClickCount() == 2) {

								TreePath tp = tree.getSelectionPath();

								DefaultMutableTreeNode node = (DefaultMutableTreeNode) tp.getLastPathComponent();

								foldername = node.getUserObject().toString();

								DefaultTableModel model = (DefaultTableModel) table_fileinformation.getModel();

								while (model.getRowCount() > 0) {

									for (int i = 0; i < model.getRowCount(); ++i) {

										model.removeRow(i);
									}
								}
								editorPane.setText("");

								DefaultTableModel model1 = (DefaultTableModel) table_1.getModel();

								while (model1.getRowCount() > 0) {

									for (int i = 0; i < model1.getRowCount(); ++i) {

										model1.removeRow(i);
									}
								}
								editorPane.setText("");

								if (fileoption.equalsIgnoreCase("MBOX")) {

									fileInformation_on_mbox();

								} else if (fileoption.equalsIgnoreCase("OLM File (.olm)")) {
									foldername = "";
									TreeNode[] folder = node.getPath();

									for (int i = 0; i < folder.length; i++) {
										String s = folder[i].toString().trim();

										s = s.replace("<html><b>", "");

										if (i == 2) {
											foldername = s.trim();
										} else if (i > 2) {
											foldername = foldername + File.separator + s.trim();
										}

									}

									if (!foldername.equalsIgnoreCase(root.toString())) {

										try {
											fileinformation_olm();
										} catch (Exception e) {

										}
									}

								} else if (fileoption.equalsIgnoreCase("Zimbra files (.tgz)")) {

									foldername = "";
									TreeNode[] folder = node.getPath();
									for (int i = 0; i < folder.length; i++) {
										String s = folder[i].toString().trim();

										s = s.replace("<html><b>", "");

										if (i == 2) {
											foldername = s.trim();
										} else if (i > 2) {
											foldername = foldername + "/" + s.trim();
										}
									}
									if (!foldername.equalsIgnoreCase(root.toString())) {

										try {
											fileinformation_Zimbra();
										} catch (Exception e) {

										}
									}

								} else if (fileoption.equalsIgnoreCase("Exchange Offline Storage (.ost)")
										|| fileoption.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
									foldername = "";

									TreeNode[] folder = node.getPath();

									for (int i = 1; i < folder.length; i++) {
										String s = folder[i].toString().trim();

										s = s.replace("<html><b>", "");
										if (i == 1) {
											foldername = s;
										} else if (i > 1) {
											foldername = foldername + File.separator + s;
										}

									}

									try {
										fileInhformation_Ost_Pst();
									} catch (Exception e) {

									}

								} else if (fileoption.equalsIgnoreCase("Opera Mail")
										|| fileoption.equalsIgnoreCase("Thunderbird")
										|| fileoption.equalsIgnoreCase("Apple Mail")) {

									foldername = ((CustomTreeNode) tp.getLastPathComponent()).filepath;

									if (!foldername.equalsIgnoreCase(root.toString())) {
										filepath = foldername;

										try {
											fileInformation_on_Thunderbird();
										} catch (Exception e) {

										}

									}

								} else if (fileoption.equalsIgnoreCase("Live Exchange")
										|| fileoption.equalsIgnoreCase("OFFICE 365")) {

								} else if (fileoption.equalsIgnoreCase("Gmail")
										|| fileoption.equalsIgnoreCase("Yahoo Mail")
										|| fileoption.equalsIgnoreCase("Aol")) {

								}

								else {

									try {
										readmailFile();
									} catch (Exception e) {

									}

								}

							}

							lblPleaseWatTable.setVisible(false);
							label_13.setVisible(false);
							table_fileinformation.setEnabled(true);
							btn_Cancel_pane2.setVisible(false);
							btn_Cancel_pane2.setEnabled(false);
							btn_next_pane2.setEnabled(true);
							btn_next_pane2.setVisible(true);
							btnAttachment.setEnabled(true);
							table_fileinformation.setEnabled(true);
							btnViewer.setEnabled(true);
							btn_next_pane2.setEnabled(true);
							btn_Previous_pane2.setEnabled(true);
						} catch (Error e) {
							logger.warning("Error : " + e.getMessage() + System.lineSeparator());
						} catch (Exception e) {
							logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
							e.printStackTrace();
							if (e.getMessage().contains("null")) {

							}

						} finally {
						}
					}
				});
				String strname = String.valueOf(doubleclickcount);
				th.setName(strname);
				th.start();

				try {
					Thread.sleep(400);
				} catch (InterruptedException e1) {

					e1.printStackTrace();
				}

			}
		});

		tree.setModel(new DefaultTreeModel(new DefaultMutableTreeNode("root folder") {

			private static final long serialVersionUID = 1L;

			{
			}
		}));

		scrollPane_fortable_p2 = new JScrollPane();
		scrollPane_fortable_p2.setBounds(288, 0, 322, 527);
		panel_2_2_1.add(scrollPane_fortable_p2);

		table_fileinformation = new JTable() {
			/**
			 *
			 */
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {

				return false;
			}
		};

		table_fileinformation.getTableHeader().setReorderingAllowed(false);
		table_fileinformation.addMouseListener(new MouseAdapter() {

			public void mouseClicked(MouseEvent arg0) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {
						label_13.setVisible(true);
						editorPane.setText("");
						DefaultTableModel model = (DefaultTableModel) table_1.getModel();
						while (model.getRowCount() > 0) {

							for (int i = 0; i < model.getRowCount(); ++i) {

								model.removeRow(i);
							}
						}
						contatcheck = false;
						calendarcheck = false;
						lblNew_setemail.setText("");

						lblNew_setsubject.setText("");

						label_date.setText("");
						CardLayout card = (CardLayout) innerCardlayout.getLayout();
						card.show(innerCardlayout, "viewer");

						btn_next_pane2.setEnabled(false);
						btn_next_pane2.setVisible(false);

						if (fileoption.equalsIgnoreCase("MBOX")) {
							MailMessage message = listmail.get(table_fileinformation.getSelectedRow());

							try {
								try {
									lblNew_setemail.setText(message.getFrom().toString());

								} catch (Exception a) {
									lblNew_setemail.setText("");
								}
								try {
									lblNew_setsubject.setText(message.getSubject());

								} catch (Exception a) {
									lblNew_setsubject.setText("");
								}
								try {

									label_date.setText(message.getDate().toString());

								} catch (Exception a) {
									label_date.setText("");
								}

								HTMLEditorKit kit = new HTMLEditorKit();
								editorPane.setEditorKit(kit);
								FileOutputStream os = new FileOutputStream(
										textField_1.getText() + File.separator + "previewHtml.html");
								message.save(os, EmlSaveOptions.getDefaultHtml());
								os.close();
								URL url = new URL(
										"file:///" + textField_1.getText() + File.separator + "previewHtml.html");
								editorPane.setPage(url);

								for (int j = 0; j < message.getAttachments().size(); j++) {
									Attachment att = message.getAttachments().get_Item(j);

									String attFileName = att.getName();
									ImageIcon icon = null;

									if (attFileName.endsWith(".pdf")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
									} else if (attFileName.endsWith(".txt")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
									} else if (attFileName.endsWith(".docx")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
									} else if (attFileName.endsWith(".zip")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
									} else {
										icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
									}
									JLabel imagelabl = new JLabel();
									imagelabl.setIcon(icon);
									DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
									modeli.addRow(new Object[] { "<html><b>" + (j + 1), "<html><b>" + attFileName,
											imagelabl });

								}

							} catch (Error e) {
								logger.warning("Error : " + e.getMessage() + System.lineSeparator());
							} catch (Exception e) {
								logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
								e.printStackTrace();
							}

						} else if (fileoption.equalsIgnoreCase("Zimbra files (.tgz)")) {

							MailMessage message = listmail.get(table_fileinformation.getSelectedRow());
							MapiMessage msg = MapiMessage.fromMailMessage(message);

							try {
								lblNew_setemail.setText(message.getFrom().toString());
							} catch (Exception a) {
								lblNew_setemail.setText("");
							}
							try {
								lblNew_setsubject.setText(message.getSubject());
							} catch (Exception a) {
								lblNew_setsubject.setText("");
							}
							try {
								label_date.setText(message.getDate().toString());
							} catch (Exception a) {
								label_date.setText("");
							}

							if (msg.getMessageClass().equals("IPM.Contact")) {
								CardLayout card1 = (CardLayout) innerCardlayout.getLayout();
								card1.show(innerCardlayout, "panel_Contact");

								MapiContact con = (MapiContact) msg.toMapiMessageItem();
								try {
									String[] compa = con.getCompanies();

									label_contactcompany.setText(compa[0]);
								} catch (Exception e) {
									label_contactcompany.setText("");
								}
								try {
									String gn = "";
									String mn = "";
									String sn = "";
									try {
										gn = con.getNameInfo().getGivenName();
									} catch (Exception e) {

									}
									try {
										mn = con.getNameInfo().getMiddleName();
									} catch (Exception e) {

									}
									try {
										sn = con.getNameInfo().getSurname();
									} catch (Exception e) {

									}

									String fn = gn + " " + mn + " " + sn;
									label_contactfullname.setText(fn);
								} catch (Exception e) {
									label_contactfullname.setText("");
								}
								try {
									label_contactemail
											.setText(con.getElectronicAddresses().getEmail1().getEmailAddress());
								} catch (Exception e) {
									label_contactemail.setText("");
								}
								try {
									label_contactphonenumber.setText(con.getTelephones().getMobileTelephoneNumber());
								} catch (Exception e) {
									label_contactphonenumber.setText("");
								}
								try {
									textArea_contact.setText(con.getPersonalInfo().getNotes());
								} catch (Exception e) {
									textArea_contact.setText("");
								}

								contatcheck = true;

							} else if (msg.getMessageClass().equals("IPM.Appointment")

									|| msg.getMessageClass().contains("IPM.Schedule.Meeting")) {
								CardLayout card1 = (CardLayout) innerCardlayout.getLayout();
								card1.show(innerCardlayout, "panel_Callendar");

								MapiCalendar cal = (MapiCalendar) msg.toMapiMessageItem();

								calendarcheck = true;
								try {
									label_Calendarsubject.setText(cal.getSubject());
								} catch (Exception e) {
									label_Calendarsubject.setText("");
								}
								try {
									label_calendarstartdate.setText(cal.getStartDate().toString());
								} catch (Exception e) {
									label_calendarstartdate.setText("");
								}
								try {
									label_Calendarenddate.setText(cal.getEndDate().toString());
								} catch (Exception e) {
									label_Calendarenddate.setText("");
								}

							} else {

								try {

									HTMLEditorKit kit = new HTMLEditorKit();
									editorPane.setEditorKit(kit);
									FileOutputStream os = new FileOutputStream(
											textField_1.getText() + File.separator + "previewHtml.html");
									message.save(os, EmlSaveOptions.getDefaultHtml());
									os.close();
									URL url = new URL(
											"file:///" + textField_1.getText() + File.separator + "previewHtml.html");
									editorPane.setPage(url);
								} catch (Error e) {
									logger.warning("Error : " + e.getMessage() + System.lineSeparator());
								} catch (Exception e) {
									logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
									editorPane.setContentType("text/html");
									editorPane.setText("<html>Page not found.</html>");
								}

							}

							int k = 1;
							for (int j = 0; j < message.getAttachments().size(); j++) {
								MapiAttachment att = msg.getAttachments().get_Item(j);

								String attFileName = att.getLongFileName();
								ImageIcon icon = null;

								if (attFileName.endsWith(".pdf")) {
									icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
								} else if (attFileName.endsWith(".txt")) {
									icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
								} else if (attFileName.endsWith(".docx")) {
									icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
								} else if (attFileName.endsWith(".zip")) {
									icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
								} else {
									icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
								}
								JLabel imagelabl = new JLabel();
								imagelabl.setIcon(icon);

								DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
								modeli.addRow(new Object[] { "<html><b>" + k, "<html><b>" + attFileName, imagelabl });
								k++;

							}

						} else if (fileoption.equalsIgnoreCase("DBX")) {
							MailMessage message = listmail.get(table_fileinformation.getSelectedRow());

							try {
								try {
									lblNew_setemail.setText(message.getFrom().toString());

								} catch (Exception a) {
									lblNew_setemail.setText("");
								}
								try {
									lblNew_setsubject.setText(message.getSubject());

								} catch (Exception a) {
									lblNew_setsubject.setText("");
								}
								try {

									label_date.setText(message.getDate().toString());

								} catch (Exception a) {
									label_date.setText("");
								}

								HTMLEditorKit kit = new HTMLEditorKit();
								editorPane.setEditorKit(kit);
								FileOutputStream os = new FileOutputStream(
										textField_1.getText() + File.separator + "previewHtml.html");
								message.save(os, EmlSaveOptions.getDefaultHtml());
								os.close();
								URL url = new URL(
										"file:///" + textField_1.getText() + File.separator + "previewHtml.html");
								editorPane.setPage(url);

								for (int j = 0; j < message.getAttachments().size(); j++) {
									Attachment att = message.getAttachments().get_Item(j);

									String attFileName = att.getName();
									ImageIcon icon = null;

									if (attFileName.endsWith(".pdf")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
									} else if (attFileName.endsWith(".txt")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
									} else if (attFileName.endsWith(".docx")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
									} else if (attFileName.endsWith(".zip")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
									} else {
										icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
									}
									JLabel imagelabl = new JLabel();
									imagelabl.setIcon(icon);
									DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
									modeli.addRow(new Object[] { "<html><b>" + (j + 1), "<html><b>" + attFileName,
											imagelabl });

								}

							} catch (Error e) {
								logger.warning("Error : " + e.getMessage() + System.lineSeparator());
							} catch (Exception e) {
								logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
								e.printStackTrace();
							}

						} else if (fileoption.equalsIgnoreCase("Opera Mail")
								|| fileoption.equalsIgnoreCase("Thunderbird")
								|| fileoption.equalsIgnoreCase("Apple Mail")) {
							MailMessage message = listmail.get(table_fileinformation.getSelectedRow());

							try {
								try {
									lblNew_setemail.setText(message.getFrom().toString());

								} catch (Exception a) {
									lblNew_setemail.setText("");
								}
								try {
									lblNew_setsubject.setText(message.getSubject());

								} catch (Exception a) {
									lblNew_setsubject.setText("");
								}
								try {

									label_date.setText(message.getDate().toString());

								} catch (Exception a) {
									label_date.setText("");
								}

								HTMLEditorKit kit = new HTMLEditorKit();
								editorPane.setEditorKit(kit);
								FileOutputStream os = new FileOutputStream(
										textField_1.getText() + File.separator + "previewHtml.html");
								message.save(os, EmlSaveOptions.getDefaultHtml());
								os.close();
								URL url = new URL(
										"file:///" + textField_1.getText() + File.separator + "previewHtml.html");
								editorPane.setPage(url);

								for (int j = 0; j < message.getAttachments().size(); j++) {
									Attachment att = message.getAttachments().get_Item(j);

									String attFileName = att.getName();
									ImageIcon icon = null;

									if (attFileName.endsWith(".pdf")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
									} else if (attFileName.endsWith(".txt")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
									} else if (attFileName.endsWith(".docx")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
									} else if (attFileName.endsWith(".zip")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
									} else {
										icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
									}
									JLabel imagelabl = new JLabel();
									imagelabl.setIcon(icon);
									DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
									modeli.addRow(new Object[] { "<html><b>" + (j + 1), "<html><b>" + attFileName,
											imagelabl });

								}

							} catch (Error e) {
								logger.warning("Error : " + e.getMessage() + System.lineSeparator());
							} catch (Exception e) {
								logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
								e.printStackTrace();
							}

						} else if (fileoption.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
								|| fileoption.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

							MapiMessage message = listmapi.get(table_fileinformation.getSelectedRow());

							try {
								lblNew_setemail.setText(message.getSenderEmailAddress());
							} catch (Exception a) {
								lblNew_setemail.setText("");
							}
							try {
								lblNew_setsubject.setText(message.getSubject());
							} catch (Exception a) {
								lblNew_setsubject.setText("");
							}
							try {
								label_date.setText(message.getDeliveryTime().toString());
							} catch (Exception a) {
								label_date.setText("");
							}

							if (message.getMessageClass().equals("IPM.Contact")) {
								CardLayout card1 = (CardLayout) innerCardlayout.getLayout();
								card1.show(innerCardlayout, "panel_Contact");

								MapiContact con = (MapiContact) message.toMapiMessageItem();
								try {
									String[] compa = con.getCompanies();

									label_contactcompany.setText(compa[0]);
								} catch (Exception e) {
									label_contactcompany.setText("");
								}
								try {
									String gn = "";
									String mn = "";
									String sn = "";
									try {
										gn = con.getNameInfo().getGivenName();
									} catch (Exception e) {

									}
									if (gn == null) {
										gn = "";
									}
									try {
										mn = con.getNameInfo().getMiddleName();
									} catch (Exception e) {

									}
									if (mn == null) {
										mn = "";
									}
									try {
										sn = con.getNameInfo().getSurname();
									} catch (Exception e) {

									}
									if (sn == null) {
										sn = "";
									}

									String fn = gn + " " + mn + " " + sn;
									label_contactfullname.setText(fn);
								} catch (Exception e) {
									label_contactfullname.setText("");
								}
								try {
									label_contactemail
											.setText(con.getElectronicAddresses().getEmail1().getEmailAddress());
								} catch (Exception e) {
									label_contactemail.setText("");
								}
								try {
									label_contactphonenumber.setText(con.getTelephones().getMobileTelephoneNumber());
								} catch (Exception e) {
									label_contactphonenumber.setText("");
								}
								try {
									textArea_contact.setText(con.getPersonalInfo().getNotes());
								} catch (Exception e) {
									textArea_contact.setText("");
								}

								contatcheck = true;

							} else if (message.getMessageClass().equals("IPM.Appointment")
									|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {

								CardLayout card1 = (CardLayout) innerCardlayout.getLayout();
								card1.show(innerCardlayout, "panel_Callendar");

								MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();

								calendarcheck = true;
								try {
									label_Calendarsubject.setText(cal.getSubject());
								} catch (Exception e) {
									label_Calendarsubject.setText("");
								}
								try {
									label_calendarstartdate.setText(cal.getStartDate().toString());
								} catch (Exception e) {
									label_calendarstartdate.setText("");
								}
								try {
									label_Calendarenddate.setText(cal.getEndDate().toString());
								} catch (Exception e) {
									label_Calendarenddate.setText("");
								}

							} else {

								try {

									HTMLEditorKit kit = new HTMLEditorKit();
									editorPane.setEditorKit(kit);
									FileOutputStream os = new FileOutputStream(
											textField_1.getText() + File.separator + "previewHtml.html");
									message.save(os, EmlSaveOptions.getDefaultHtml());
									os.close();
									URL url = new URL(
											"file:///" + textField_1.getText() + File.separator + "previewHtml.html");
									editorPane.setPage(url);
								} catch (Error e) {
									logger.warning("Error : " + e.getMessage() + System.lineSeparator());
								} catch (Exception e) {
									logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
									editorPane.setContentType("text/html");
									editorPane.setText("<html>Page not found.</html>");
								}

							}

							int k = 1;
							for (int j = 0; j < message.getAttachments().size(); j++) {
								MapiAttachment att = message.getAttachments().get_Item(j);

								String attFileName = att.getLongFileName();
								ImageIcon icon = null;

								if (attFileName.endsWith(".pdf")) {
									icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
								} else if (attFileName.endsWith(".txt")) {
									icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
								} else if (attFileName.endsWith(".docx")) {
									icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
								} else if (attFileName.endsWith(".zip")) {
									icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
								} else {
									icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
								}
								JLabel imagelabl = new JLabel();
								imagelabl.setIcon(icon);

								DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
								modeli.addRow(new Object[] { "<html><b>" + k, "<html><b>" + attFileName, imagelabl });
								k++;

							}

						} else if (fileoption.equalsIgnoreCase("OLM File (.olm)")) {
							MapiMessage message = null;

							try {
								message = listmapi.get(table_fileinformation.getSelectedRow());
							} catch (Exception e) {

							}

							try {
								lblNew_setemail.setText(message.getSenderEmailAddress());
							} catch (Exception a) {
								lblNew_setemail.setText("");
							}
							try {
								lblNew_setsubject.setText(message.getSubject());
							} catch (Exception a) {
								lblNew_setsubject.setText("");
							}
							try {
								label_date.setText(message.getDeliveryTime().toString());
							} catch (Exception a) {
								label_date.setText("");
							}
							if (message.getMessageClass().equals("IPM.Contact")) {
								CardLayout card1 = (CardLayout) innerCardlayout.getLayout();
								card1.show(innerCardlayout, "panel_Contact");

								MapiContact con = (MapiContact) message.toMapiMessageItem();
								try {
									String[] compa = con.getCompanies();

									label_contactcompany.setText(compa[0]);
								} catch (Exception e) {
									label_contactcompany.setText("");
								}
								try {
									label_contactfullname.setText(con.getNameInfo().getDisplayName());
								} catch (Exception e) {
									label_contactfullname.setText("");
								}
								try {
									label_contactemail.setText(
											con.getElectronicAddresses().getDefaultEmailAddress().getEmailAddress());
								} catch (Exception e) {
									label_contactemail.setText("");
								}
								try {
									label_contactphonenumber.setText(con.getTelephones().getDefaultTelephoneNumber());
								} catch (Exception e) {
									label_contactphonenumber.setText("");
								}
								try {
									textArea_contact.setText(con.getPersonalInfo().getNotes());
								} catch (Exception e) {
									textArea_contact.setText("");
								}

								contatcheck = true;

							} else if (message.getMessageClass().equals("IPM.Appointment")
									|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {
								CardLayout card1 = (CardLayout) innerCardlayout.getLayout();
								card1.show(innerCardlayout, "panel_Callendar");

								MapiCalendar cal = (MapiCalendar) message.toMapiMessageItem();
								calendarcheck = true;
								try {
									label_Calendarsubject.setText(cal.getSubject());
								} catch (Exception e) {
									label_Calendarsubject.setText("");
								}
								try {
									label_calendarstartdate.setText(cal.getStartDate().toString());
								} catch (Exception e) {
									label_calendarstartdate.setText("");
								}
								try {
									label_Calendarenddate.setText(cal.getEndDate().toString());
								} catch (Exception e) {
									label_Calendarenddate.setText("");
								}

							} else {

								try {

									HTMLEditorKit kit = new HTMLEditorKit();
									editorPane.setEditorKit(kit);
									FileOutputStream os = new FileOutputStream(
											textField_1.getText() + File.separator + "previewHtml.html");
									message.save(os, EmlSaveOptions.getDefaultHtml());
									os.close();
									URL url = new URL(
											"file:///" + textField_1.getText() + File.separator + "previewHtml.html");
									editorPane.setPage(url);
								} catch (Error e) {
									logger.warning("Error : " + e.getMessage() + System.lineSeparator());
								} catch (Exception e) {
									logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
									editorPane.setContentType("text/html");
									editorPane.setText("<html>Page not found.</html>");
								}

							}
							int k = 1;
							for (int j = 0; j < message.getAttachments().size(); j++) {
								MapiAttachment att = message.getAttachments().get_Item(j);

								String attFileName = att.getLongFileName();
								ImageIcon icon = null;

								if (attFileName.endsWith(".pdf")) {
									icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
								} else if (attFileName.endsWith(".txt")) {
									icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
								} else if (attFileName.endsWith(".docx")) {
									icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
								} else if (attFileName.endsWith(".zip")) {
									icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
								} else {
									icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
								}
								JLabel imagelabl = new JLabel();
								imagelabl.setIcon(icon);

								DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
								modeli.addRow(new Object[] { "<html><b>" + k, "<html><b>" + attFileName, imagelabl });
								k++;

							}
						} else if (fileoption.equalsIgnoreCase("Nodes Storage (.nsf)")) {

						} /*
							 * else if (fileoption.equalsIgnoreCase("Yahoo Mail") ||
							 * fileoption.equalsIgnoreCase("Gmail") || fileoption.equalsIgnoreCase("Aol")) {
							 *
							 * try { MailMessage msg1 =
							 * listmail.get(table_fileinformation.getSelectedRow());
							 *
							 * ImapMessageInfo msgInfo =
							 * listImapmesinfo.get(table_fileinformation.getSelectedRow());
							 *
							 * //////System.out.println("message found");
							 *
							 * lblNew_setemail.setText(msgInfo.getSender().toString());
							 *
							 * lblNew_setsubject.setText(msgInfo.getSubject());
							 *
							 * label_date.setText(msg1.getDate().toString());
							 *
							 * //////System.out.println("found");
							 *
							 * HTMLEditorKit kit = new HTMLEditorKit(); editorPane.setEditorKit(kit);
							 * FileOutputStream os = new FileOutputStream( textField_1.getText() +
							 * File.separator + "previewHtml.html"); msg1.save(os,
							 * EmlSaveOptions.getDefaultHtml()); os.close(); URL url = new URL( "file:///" +
							 * textField_1.getText() + File.separator + "previewHtml.html");
							 * editorPane.setPage(url);
							 *
							 * for (int k = 0; k < msg1.getAttachments().size(); k++) { Attachment att =
							 * msg1.getAttachments().get_Item(k);
							 *
							 * String attFileName = att.getName(); ImageIcon icon = null;
							 *
							 * if (attFileName.endsWith(".pdf")) { icon = new
							 * ImageIcon(Main_Frame.class.getResource("/pdf-icon.png")); } else if
							 * (attFileName.endsWith(".txt")) { icon = new
							 * ImageIcon(Main_Frame.class.getResource("/txt-icon.png")); } else if
							 * (attFileName.endsWith(".docx")) { icon = new
							 * ImageIcon(Main_Frame.class.getResource("/Doc-icon.png")); } else if
							 * (attFileName.endsWith(".zip")) { icon = new
							 * ImageIcon(Main_Frame.class.getResource("/zip-icon.png")); } else { icon = new
							 * ImageIcon(Main_Frame.class.getResource("/attachment-icon.png")); } JLabel
							 * imagelabl = new JLabel(); imagelabl.setIcon(icon);
							 *
							 * DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
							 * modeli.addRow(new Object[] { "<html><b>" + (j + 1), "<html><b>" +
							 * attFileName, imagelabl }); //////System.out.println(attFileName);
							 *
							 * }
							 *
							 * } catch (Error e) { logger.warning("Error : " + e.getMessage() +
							 * System.lineSeparator()); } catch (Exception e) {
							 * logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
							 * //e.printStackTrace();
							 *
							 * }
							 *
							 * }
							 *
							 * else if (fileoption.equals("OFFICE 365") ||
							 * fileoption.equals("Live Exchange")) {
							 *
							 * MailMessage message = listmail.get(table_fileinformation.getSelectedRow());
							 *
							 * lblNew_setemail.setText(message.getFrom().toString());
							 *
							 * lblNew_setsubject.setText(message.getSubject());
							 *
							 * label_date.setText(message.getDate().toString());
							 *
							 * //////System.out.println("found");
							 *
							 * try { HTMLEditorKit kit = new HTMLEditorKit(); editorPane.setEditorKit(kit);
							 * FileOutputStream os = new FileOutputStream( textField_1.getText() +
							 * File.separator + "previewHtml.html"); message.save(os,
							 * EmlSaveOptions.getDefaultHtml()); os.close(); URL url = new URL( "file:///" +
							 * textField_1.getText() + File.separator + "previewHtml.html");
							 * editorPane.setPage(url);
							 *
							 * for (int j = 0; j < message.getAttachments().size(); j++) { Attachment att =
							 * message.getAttachments().get_Item(j);
							 *
							 * String attFileName = att.getName(); ImageIcon icon = null;
							 *
							 * if (attFileName.endsWith(".pdf")) { icon = new
							 * ImageIcon(Main_Frame.class.getResource("/pdf-icon.png")); } else if
							 * (attFileName.endsWith(".txt")) { icon = new
							 * ImageIcon(Main_Frame.class.getResource("/txt-icon.png")); } else if
							 * (attFileName.endsWith(".docx")) { icon = new
							 * ImageIcon(Main_Frame.class.getResource("/Doc-icon.png")); } else if
							 * (attFileName.endsWith(".zip")) { icon = new
							 * ImageIcon(Main_Frame.class.getResource("/zip-icon.png")); } else { icon = new
							 * ImageIcon(Main_Frame.class.getResource("/attachment-icon.png")); } JLabel
							 * imagelabl = new JLabel(); imagelabl.setIcon(icon);
							 *
							 * DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
							 * modeli.addRow(new Object[] { "<html><b>" + (j + 1), "<html><b>" +
							 * attFileName, imagelabl }); //////System.out.println(attFileName);
							 *
							 * }
							 *
							 * } catch (Error e) { logger.warning("Error : " + e.getMessage() +
							 * System.lineSeparator()); } catch (Exception e) {
							 * logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
							 * //e.printStackTrace(); }
							 *
							 * }
							 */ else {
							MailMessage message = MailMessage.load(filepath);

							// //System.out.println("found");

							try {
								try {
									lblNew_setemail.setText(message.getFrom().toString());
								} catch (Exception e) {

								}
								try {
									lblNew_setsubject.setText(message.getSubject());
								} catch (Exception e) {

								}
								try {
									label_date.setText(message.getDate().toString());
								} catch (Exception e) {

								}
								HTMLEditorKit kit = new HTMLEditorKit();
								editorPane.setEditorKit(kit);
								FileOutputStream os = new FileOutputStream(
										textField_1.getText() + File.separator + "previewHtml.html");
								message.save(os, EmlSaveOptions.getDefaultHtml());
								os.close();
								URL url = new URL(
										"file:///" + textField_1.getText() + File.separator + "previewHtml.html");
								editorPane.setPage(url);

								for (int j = 0; j < message.getAttachments().size(); j++) {
									Attachment att = message.getAttachments().get_Item(j);

									String attFileName = att.getName();
									ImageIcon icon = null;

									if (attFileName.endsWith(".pdf")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/pdf-icon.png"));
									} else if (attFileName.endsWith(".txt")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/txt-icon.png"));
									} else if (attFileName.endsWith(".docx")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/Doc-icon.png"));
									} else if (attFileName.endsWith(".zip")) {
										icon = new ImageIcon(Main_Frame.class.getResource("/zip-icon.png"));
									} else {
										icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
									}
									JLabel imagelabl = new JLabel();
									imagelabl.setIcon(icon);

									DefaultTableModel modeli = (DefaultTableModel) table_1.getModel();
									modeli.addRow(new Object[] { "<html><b>" + (j + 1), "<html><b>" + attFileName,
											imagelabl });
									// //System.out.println(attFileName);
								}
							} catch (Error e) {
								logger.warning("Error : " + e.getMessage() + System.lineSeparator());
							} catch (Exception e) {
								logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
								e.printStackTrace();
							}
						}
						btn_next_pane2.setEnabled(true);
						btn_next_pane2.setVisible(true);

						label_13.setVisible(false);
					}

				});
				// th1.start();
			}
		});

		scrollPane_fortable_p2.setViewportView(table_fileinformation);

		table_fileinformation.setBackground(Color.WHITE);

		table_fileinformation.setModel(new DefaultTableModel(new Object[][] {},
				new String[] { "From", "Subject", "Received Date", "Attachment" }));
		table_fileinformation.getColumn("Attachment").setCellRenderer(new Renderer());
		table_fileinformation.setFocusable(false);
		table_fileinformation.setRowSelectionAllowed(true);

		lbl_Email = new JLabel("From");
		lbl_Email.setForeground(SystemColor.textHighlight);
		lbl_Email.setFont(new Font("Tahoma", Font.BOLD, 11));
		lbl_Email.setBounds(611, 2, 53, 22);
		panel_2_2_1.add(lbl_Email);

		lbl_subject = new JLabel("Subject");
		lbl_subject.setForeground(SystemColor.textHighlight);
		lbl_subject.setFont(new Font("Tahoma", Font.BOLD, 11));
		lbl_subject.setBounds(838, 32, 56, 27);
		panel_2_2_1.add(lbl_subject);

		lblNew_setemail = new JLabel("");
		lblNew_setemail.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblNew_setemail.setBounds(650, 2, 178, 22);
		panel_2_2_1.add(lblNew_setemail);

		lblNew_setsubject = new JLabel("");
		lblNew_setsubject.setFont(new Font("Tahoma", Font.PLAIN, 10));
		lblNew_setsubject.setBounds(891, 38, 177, 16);
		panel_2_2_1.add(lblNew_setsubject);

		lbl_Date = new JLabel("Date");
		lbl_Date.setForeground(SystemColor.textHighlight);
		lbl_Date.setFont(new Font("Tahoma", Font.BOLD, 11));
		lbl_Date.setBounds(838, 4, 36, 16);
		panel_2_2_1.add(lbl_Date);

		label_date = new JLabel("");
		label_date.setFont(new Font("Tahoma", Font.PLAIN, 10));
		label_date.setBounds(891, 4, 177, 16);
		panel_2_2_1.add(label_date);

		innerCardlayout = new JPanel();
		innerCardlayout.setBounds(611, 58, 464, 492);
		panel_2_2_1.add(innerCardlayout);
		innerCardlayout.setLayout(new CardLayout(0, 0));

		viewer = new JPanel();
		innerCardlayout.add(viewer, "viewer");

		btnViewer = new JButton("");
		btnViewer.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnViewer.setIcon(new ImageIcon(Main_Frame.class.getResource("/preview-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent arg0) {
				btnViewer.setIcon(new ImageIcon(Main_Frame.class.getResource("/preview-btn.png")));
			}
		});

		btnViewer.setIcon(new ImageIcon(Main_Frame.class.getResource("/preview-btn.png")));
		btnViewer.setRolloverEnabled(false);
		btnViewer.setRequestFocusEnabled(false);
		btnViewer.setOpaque(false);
		btnViewer.setToolTipText("Click here to Preview of the mail. ");
		btnViewer.setFocusable(false);
		btnViewer.setFocusTraversalKeysEnabled(false);
		btnViewer.setFocusPainted(false);
		btnViewer.setDefaultCapable(false);
		btnViewer.setContentAreaFilled(false);
		btnViewer.setBorderPainted(false);
		btnViewer.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent e) {

				if (contatcheck) {

					CardLayout card1 = (CardLayout) innerCardlayout.getLayout();
					card1.show(innerCardlayout, "panel_Contact");

				} else if (calendarcheck) {
					CardLayout card1 = (CardLayout) innerCardlayout.getLayout();
					card1.show(innerCardlayout, "panel_Callendar");
				} else {
					CardLayout card = (CardLayout) innerCardlayout.getLayout();

					card.show(innerCardlayout, "viewer");
				}

			}
		});
		btnViewer.setBounds(610, 27, 112, 33);
		panel_2_2_1.add(btnViewer);

		btnAttachment = new JButton("");
		btnAttachment.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnAttachment.setIcon(new ImageIcon(Main_Frame.class.getResource("/attachment-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent arg0) {
				btnAttachment.setIcon(new ImageIcon(Main_Frame.class.getResource("/attachment-btn.png")));
			}
		});

		btnAttachment.setIcon(new ImageIcon(Main_Frame.class.getResource("/attachment-btn.png")));
		btnAttachment.setDefaultCapable(false);
		btnAttachment.setFocusTraversalKeysEnabled(false);
		btnAttachment.setRolloverEnabled(false);
		btnAttachment.setToolTipText("Click here to See the attachment  of the mail. ");
		btnAttachment.setRequestFocusEnabled(false);
		btnAttachment.setOpaque(false);
		btnAttachment.setFocusable(false);
		btnAttachment.setFocusPainted(false);
		btnAttachment.setContentAreaFilled(false);
		btnAttachment.setBorderPainted(false);
		btnAttachment.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				CardLayout card = (CardLayout) innerCardlayout.getLayout();

				card.show(innerCardlayout, "attachment");
			}
		});
		btnAttachment.setBounds(723, 27, 112, 32);
		panel_2_2_1.add(btnAttachment);
		viewer.setLayout(null);

		scrollPane = new JScrollPane();
		scrollPane.setBounds(0, 0, 464, 492);
		viewer.add(scrollPane);

		editorPane = new JEditorPane();
		scrollPane.setViewportView(editorPane);
		editorPane.setEditable(false);

		panel_Contact = new JPanel();
		panel_Contact.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_Contact.setBackground(Color.WHITE);
		innerCardlayout.add(panel_Contact, "panel_Contact");
		panel_Contact.setLayout(null);

		lblFullName = new JLabel("Full Name ");
		lblFullName.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblFullName.setBounds(10, 67, 64, 14);
		panel_Contact.add(lblFullName);

		label_contactfullname = new JLabel("");
		label_contactfullname.setBounds(83, 67, 215, 14);
		panel_Contact.add(label_contactfullname);

		lblEMailAd = new JLabel("E Mail :");
		lblEMailAd.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblEMailAd.setBounds(10, 92, 64, 21);
		panel_Contact.add(lblEMailAd);

		lblCompany = new JLabel("Company");
		lblCompany.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblCompany.setBounds(10, 124, 64, 14);
		panel_Contact.add(lblCompany);

		label_contactemail = new JLabel("");
		label_contactemail.setBounds(83, 95, 215, 18);
		panel_Contact.add(label_contactemail);

		label_contactcompany = new JLabel("");
		label_contactcompany.setBounds(83, 124, 215, 21);
		panel_Contact.add(label_contactcompany);

		lblPhoneNo = new JLabel("Phone No");
		lblPhoneNo.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblPhoneNo.setBounds(10, 149, 56, 21);
		panel_Contact.add(lblPhoneNo);

		label_contactphonenumber = new JLabel("");
		label_contactphonenumber.setBounds(80, 152, 218, 18);
		panel_Contact.add(label_contactphonenumber);

		JLabel lblNotes = new JLabel("Notes");
		lblNotes.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNotes.setBounds(554, 48, 115, 21);
		panel_Contact.add(lblNotes);

		textArea_contact = new JTextArea();
		textArea_contact.setEditable(false);
		textArea_contact.setBounds(10, 223, 427, 258);
		panel_Contact.add(textArea_contact);

		label_contacticon = new JLabel("");
		label_contacticon.setIcon(new ImageIcon(Main_Frame.class.getResource("/User-Chat-icon.png")));
		label_contacticon.setBounds(10, 11, 64, 51);
		panel_Contact.add(label_contacticon);

		panel_Callendar = new JPanel();
		panel_Callendar.setBorder(new LineBorder(new Color(0, 0, 0)));
		panel_Callendar.setBackground(Color.WHITE);
		innerCardlayout.add(panel_Callendar, "panel_Callendar");
		panel_Callendar.setLayout(null);

		lblSubject = new JLabel("Subject");
		lblSubject.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblSubject.setBounds(10, 92, 52, 28);
		panel_Callendar.add(lblSubject);

		label_Calendarsubject = new JLabel("");
		label_Calendarsubject.setBounds(85, 98, 323, 22);
		panel_Callendar.add(label_Calendarsubject);

		JLabel lblStartDate = new JLabel("Start Date");
		lblStartDate.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblStartDate.setBounds(10, 148, 65, 22);
		panel_Callendar.add(lblStartDate);

		label_calendarstartdate = new JLabel("");
		label_calendarstartdate.setBounds(85, 148, 323, 28);
		panel_Callendar.add(label_calendarstartdate);

		label_Calendaricon = new JLabel("");
		label_Calendaricon.setIcon(new ImageIcon(Main_Frame.class.getResource("/calender.png")));
		label_Calendaricon.setBounds(10, 11, 116, 77);
		panel_Callendar.add(label_Calendaricon);

		JLabel lblEndDate = new JLabel("End Date");
		lblEndDate.setFont(new Font("Tahoma", Font.BOLD, 12));
		lblEndDate.setBounds(10, 204, 71, 17);
		panel_Callendar.add(lblEndDate);

		label_Calendarenddate = new JLabel("");
		label_Calendarenddate.setBounds(85, 204, 323, 28);
		panel_Callendar.add(label_Calendarenddate);

		attachment = new JPanel();
		innerCardlayout.add(attachment, "attachment");
		attachment.setLayout(null);

		scrollPane_1 = new JScrollPane();
		scrollPane_1.setBounds(0, 0, 474, 492);
		attachment.add(scrollPane_1);

		table_1 = new JTable() {

			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {

				return false;
			}
		};

		table_1.getTableHeader().setReorderingAllowed(false);
		table_1.setModel(new DefaultTableModel(new Object[][] {}, new String[] { "S No", "File Name", "File Type" }));
		table_1.getColumnModel().getColumn(0).setPreferredWidth(35);
		table_1.getColumnModel().getColumn(1).setPreferredWidth(198);
		table_1.getColumn("File Type").setCellRenderer(new Renderer());
		scrollPane_1.setViewportView(table_1);

		btn_next_pane2 = new JButton("");
		btn_next_pane2.setBounds(950, 562, 110, 30);
		panel_2_2_1.add(btn_next_pane2);
		btn_next_pane2.setDefaultCapable(false);
		btn_next_pane2.setContentAreaFilled(false);
		btn_next_pane2.setBorderPainted(false);
		btn_next_pane2.setRequestFocusEnabled(false);
		btn_next_pane2.setOpaque(false);
		btn_next_pane2.setRolloverEnabled(false);
		btn_next_pane2.setFocusable(false);
		btn_next_pane2.setFocusTraversalKeysEnabled(false);
		btn_next_pane2.setFocusPainted(false);
		btn_next_pane2.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_next_pane2.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent arg0) {
				btn_next_pane2.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-btn.png")));
			}
		});

		btn_next_pane2.setIcon(new ImageIcon(Main_Frame.class.getResource("/next-btn.png")));

		btn_next_pane2.setFont(new Font("Tahoma", Font.BOLD, 12));

		btn_Previous_pane2 = new JButton("");
		btn_Previous_pane2.setToolTipText("Click here to Go Back.");
		btn_Previous_pane2.setBounds(822, 562, 110, 30);
		panel_2_2_1.add(btn_Previous_pane2);
		btn_Previous_pane2.setRolloverEnabled(false);
		btn_Previous_pane2.setRequestFocusEnabled(false);
		btn_Previous_pane2.setOpaque(false);
		btn_Previous_pane2.setFocusable(false);
		btn_Previous_pane2.setFocusTraversalKeysEnabled(false);
		btn_Previous_pane2.setFocusPainted(false);
		btn_Previous_pane2.setDefaultCapable(false);
		btn_Previous_pane2.setContentAreaFilled(false);
		if (messageboxtitle.contains("Thunderbird") || messageboxtitle.contains("Opera Mail")) {
			btn_Previous_pane2.setVisible(false);
		}
		btn_Previous_pane2.setBorderPainted(false);
		btn_Previous_pane2.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_Previous_pane2.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_Previous_pane2.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
			}
		});

		btn_Previous_pane2.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
		btn_Previous_pane2.setFont(new Font("Tahoma", Font.BOLD, 12));

		btn_Cancel_pane2 = new JButton("");
		btn_Cancel_pane2.setToolTipText("Click here to Stop the Process.");
		btn_Cancel_pane2.setBounds(690, 562, 110, 30);
		btn_Cancel_pane2.setVisible(false);
		btn_Cancel_pane2.setEnabled(false);
		panel_2_2_1.add(btn_Cancel_pane2);
		btn_Cancel_pane2.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_Cancel_pane2.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_Cancel_pane2.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-btn.png")));
			}
		});

		btn_Cancel_pane2.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-btn.png")));
		btn_Cancel_pane2.setFocusable(false);
		btn_Cancel_pane2.setFocusTraversalKeysEnabled(false);
		btn_Cancel_pane2.setFocusPainted(false);
		btn_Cancel_pane2.setRolloverEnabled(false);
		btn_Cancel_pane2.setRequestFocusEnabled(false);
		btn_Cancel_pane2.setOpaque(false);
		btn_Cancel_pane2.setContentAreaFilled(false);
		btn_Cancel_pane2.setBorderPainted(false);
		btn_Cancel_pane2.setDefaultCapable(false);
		btn_Cancel_pane2.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {

				String warn = "Do you want to stop the process?";
				int ans = JOptionPane.showConfirmDialog(Main_Frame.this, warn, messageboxtitle,
						JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {
					Stoppreview = true;
					// th.interrupted();
					lblPleaseWatTable.setVisible(false);
					label_13.setVisible(false);
					btnAttachment.setEnabled(true);
					table_fileinformation.setEnabled(true);
					btnViewer.setEnabled(true);
					btn_next_pane2.setEnabled(true);
					btn_Previous_pane2.setEnabled(true);
				}

			}
		});
		btn_Cancel_pane2.setFont(new Font("Tahoma", Font.BOLD, 12));

		lblPleaseWatTable = new JLabel("Please wait...");
		lblPleaseWatTable.setForeground(SystemColor.textHighlight);
		lblPleaseWatTable.setFont(new Font("Tahoma", Font.BOLD | Font.ITALIC, 9));
		lblPleaseWatTable.setVisible(false);
		lblPleaseWatTable.setBounds(9, 552, 82, 14);
		panel_2_2_1.add(lblPleaseWatTable);

		JLabel lblNewLabel_6 = new JLabel("");
		lblNewLabel_6.setBounds(0, 552, 1099, 68);
		panel_2_2_1.add(lblNewLabel_6);
		lblNewLabel_6.setIcon(new ImageIcon(Main_Frame.class.getResource("/bottom.png")));
		btn_next_pane2.setToolTipText("Click here to Go to the Previous panel. ");
		btn_Previous_pane2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				tools.setVisible(true);
				passwordField_p2.setEditable(true);
				buttonGroup_file.clearSelection();
				rdbtnSingleFile.setSelected(true);
				lblTotalMessageCount.setText("Total Message Count :");
				tf_portNo_p2.setEnabled(true);
				label_contactfullname.setText("");
				label_contactemail.setText("");
				label_contactcompany.setText("");
				label_contactphonenumber.setText("");
				textArea_contact.setText("");
				textField_Domainname_p2.setEnabled(true);
				textField_username_p2.setEnabled(true);
				passwordField_p2.setEnabled(true);
				textField_Domainname_p2.setText("");
				textField_username_p2.setText("");
				passwordField_p2.setText("");
				btnSavingLog_1.setEnabled(true);
				btnSavingLog_1.setVisible(true);
				btnTempPath.setEnabled(true);
				lblTotalMessageCount.setVisible(false);
				btnTempPath.setVisible(true);
				CardLayout card1 = (CardLayout) innerCardlayout.getLayout();
				card1.show(innerCardlayout, "panel_Contact");
				lblPleaseWatTable.setVisible(false);
				label_13.setVisible(false);
				btnAttachment.setEnabled(true);
				table_fileinformation.setEnabled(true);
				btnViewer.setEnabled(true);
				btn_next_pane2.setEnabled(true);
				btn_Previous_pane2.setEnabled(true);
				try {
					if (input) {
						if (fileoption.equalsIgnoreCase("OFFICE 365") || fileoption.equalsIgnoreCase("Live Exchange")
								|| fileoption.equalsIgnoreCase("Hotmail")) {
							clientforexchange_input.dispose();

						} else {
							iconnforimap_input.dispose();

						}
					}
				} catch (Exception e1) {
				}

				buttonGroup_file.clearSelection();
				rdbtnSingleFile.setSelected(true);
				lblNew_setemail.setText("");

				lblNew_setsubject.setText("");

				label_date.setText("");

				if (ad.sop.length == 1) {
					comboBox_FiletypeChooser.setEnabled(false);
				} else
					comboBox_FiletypeChooser.setEnabled(true);
				if (projectTitle.contains("Aryson Email Migration Tool")
						|| projectTitle.contains("Cigati Email Migrator")
						|| projectTitle.contains("DRS Email Migration Tool")) {

					comboBox_fileDestination_type.removeAll();
					comboBox_fileDestination_type.setModel(new DefaultComboBoxModel<String>(file_sfd));
				}

				file = null;

				editorPane.setText("");

				DefaultTableModel model = (DefaultTableModel) table_fileinformation.getModel();

				while (model.getRowCount() > 0) {

					for (int i = 0; i < model.getRowCount(); ++i) {

						model.removeRow(i);
					}
				}

				DefaultTableModel model1 = (DefaultTableModel) table_1.getModel();

				while (model1.getRowCount() > 0) {

					for (int i = 0; i < model1.getRowCount(); ++i) {

						model1.removeRow(i);
					}
				}

				tree.getCheckingModel();
				file = null;
				tree.clearSelection();
				DefaultTreeModel model1s = (DefaultTreeModel) tree.getModel();

				DefaultMutableTreeNode root = (DefaultMutableTreeNode) model1s.getRoot();
				root.removeAllChildren();
				model1s.reload();
				TreePath[] ac = new TreePath[0];
				tree.setCheckingPaths(ac);

				CardLayout card = (CardLayout) Cardlayout.getLayout();

				card.show(Cardlayout, "panel_1");

				if (fileoption.equalsIgnoreCase("Thunderbird") || fileoption.equalsIgnoreCase("Opera Mail")) {

					SwingUtilities.invokeLater(new Runnable() {

						public void run() {
							comboBox_FiletypeChooser.setSelectedItem("MBOX");
						}
					});
				}

			}
		});
		btn_next_pane2.setToolTipText("Click here to Go to the Conversion.");

		lblTotalMessageCount = new JLabel("Total Message Count :");
		lblTotalMessageCount.setForeground(new Color(0, 120, 215));
		lblTotalMessageCount.setBounds(288, 527, 322, 22);
		lblTotalMessageCount.setVisible(false);
		panel_2_2_1.add(lblTotalMessageCount);
		btn_next_pane2.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {

				if (nextTime != null) {
					DefaultMutableTreeNode firstLeaf = ((DefaultMutableTreeNode) tree.getModel().getRoot());
					tree.setCheckingPath(new TreePath(firstLeaf.getPath()));
				}
				TreePath[] tp = tree.getCheckingPaths();
				System.out.println(tree.isSelectionEmpty());
				String selectedpath = "";
				if (tp.length == 0) {
					JOptionPane.showMessageDialog(frame, "Select File From the Tree", messageboxtitle,
							JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				} else {
					stop = false;
					chckbxRemoveDuplicacy.setVisible(true);
					label_Remove_Duplicate.setVisible(true);
					listFolderinfofinal.clear();
					listExchangdinal.clear();
					pstfolderlist = new ArrayList<String>();
					for (int i = 0; i < tp.length; i++) {

						String pathoffile = "";
						try {
							DefaultMutableTreeNode d1 = (DefaultMutableTreeNode) tp[i].getLastPathComponent();

							String[] str = (tp[i].toString().replace("<html><b>", "").replaceAll("[\\[\\]]", ""))
									.split(",");

							if (fileoption.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
									|| fileoption.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {
								for (int j = 1; j < str.length; j++) {
									if (j == 1) {
										pathoffile = str[j].trim();

									} else if (j > 1) {
										pathoffile = pathoffile + File.separator + str[j].trim();
									}

								}

								selectedpath = pathoffile;
								System.out.println(selectedpath);
							} else if (fileoption.equalsIgnoreCase("OLM File (.olm)")) {
								for (int j = 2; j < str.length; j++) {
									if (j == 2) {
										pathoffile = str[j].trim();
										if (!pstfolderlist.contains(pathoffile)) {
											pstfolderlist.add(pathoffile);
										}
									} else if (j > 2) {
										pathoffile = pathoffile + File.separator + str[j].trim();
										if (!pstfolderlist.contains(pathoffile)) {
											pstfolderlist.add(pathoffile);
										}

									}
								}

							} else if (fileoption.equalsIgnoreCase("Zimbra files (.tgz)")) {
								for (int j = 2; j < str.length; j++) {
									if (j == 2) {
										pathoffile = str[j].trim();
										if (!pstfolderlist.contains(pathoffile)) {
											pstfolderlist.add(pathoffile);
										}
									} else if (j > 2) {
										pathoffile = pathoffile + File.separator + str[j].trim();
										if (!pstfolderlist.contains(pathoffile)) {
											pstfolderlist.add(pathoffile);
										}

									}
								}

							} else if (fileoption.equalsIgnoreCase("Opera Mail")
									|| fileoption.equalsIgnoreCase("Thunderbird")
									|| fileoption.equalsIgnoreCase("Apple Mail")) {

								try {
									pathoffile = ((CustomTreeNode) tp[i].getLastPathComponent()).filepath;
								} catch (Exception e) {
									pathoffile = "";
								}

								if (new File(pathoffile).isFile()) {
									if (!pstfolderlist.contains(pathoffile)) {
										pstfolderlist.add(pathoffile);

									}
								}

							} else if (fileoption.equalsIgnoreCase("Yahoo Mail") || fileoption.equalsIgnoreCase("Gmail")
									|| fileoption.equalsIgnoreCase("AOL") || fileoption.equalsIgnoreCase("IMAP")
									|| fileoption.equalsIgnoreCase("Zoho Mail")
									|| fileoption.equalsIgnoreCase("Yandex Mail")
									|| fileoption.equalsIgnoreCase("Amazon WorkMail")
									|| fileoption.equalsIgnoreCase("Hostgator email")
									|| fileoption.equalsIgnoreCase("Icloud")
									|| fileoption.equalsIgnoreCase("GoDaddy email")) {

								for (int j = 2; j < str.length; j++) {
									if (j == 2) {

										pathoffile = str[j].trim();
										if (pathoffile.equals("Gmail")) {

											continue;

										}
										try {
											int kki = listFolderinfostring.indexOf(pathoffile);

											if (!listFolderinfofinal.contains(listFolderinfo.get(kki))) {
												listFolderinfofinal.add(listFolderinfo.get(kki));

											}

										} catch (Exception e) {
											e.printStackTrace();
											continue;

										}
										if (!pstfolderlist.contains(pathoffile)) {
											pstfolderlist.add(pathoffile);
										}
									} else if (j > 2) {
										pathoffile = pathoffile + File.separator + str[j].trim();
										try {
											int kki = listFolderinfostring.indexOf(pathoffile);

											if (!listFolderinfofinal.contains(listFolderinfo.get(kki))) {
												listFolderinfofinal.add(listFolderinfo.get(kki));

											}
										} catch (Exception e) {

											continue;

										}
										if (!pstfolderlist.contains(pathoffile)) {
											pstfolderlist.add(pathoffile);

										}

									}
								}

							} else if (fileoption.equalsIgnoreCase("Live Exchange")
									|| fileoption.equalsIgnoreCase("OFFICE 365")
									|| fileoption.equalsIgnoreCase("Hotmail")) {
								String pathews = "";
								for (int j = 2; j < str.length; j++) {
									if (j == 2) {
										pathoffile = str[j].trim();
										pathews = getRidOfIllegalFileNameCharacters(str[j].trim());
										try {
											int kki = listFolderinfostring.indexOf(pathoffile);

											if (!listExchangdinal.contains(listExchangemesingos.get(kki))) {
												listExchangdinal.add(listExchangemesingos.get(kki));
											}

										} catch (Exception e) {
											continue;

										}
										if (!pstfolderlist.contains(pathoffile)) {

											pstfolderlist.add(pathews);
										}
									} else if (j > 2) {
										pathoffile = pathoffile + File.separator + str[j].trim();
										pathews = pathews + File.separator
												+ getRidOfIllegalFileNameCharacters(str[j].trim());
										try {
											int kki = listFolderinfostring.indexOf(pathoffile);

											if (!listExchangdinal.contains(listExchangemesingos.get(kki))) {
												listExchangdinal.add(listExchangemesingos.get(kki));
											}
										} catch (Exception e) {
											continue;

										}
										if (!pstfolderlist.contains(pathoffile)) {
											pstfolderlist.add(pathews);
										}

									}
								}

							}
							if (!pstfolderlist.contains(selectedpath)) {
								pstfolderlist.add(selectedpath);
							}

						} catch (Exception e) {
							throw new RuntimeException();
						}
					}
					System.out.println(pstfolderlist);
					listmail.clear();
					listmapi.clear();

					listPSTOSTgemesingo.clear();
					listExchangemesingo.clear();

					btn_signout_p3.setVisible(false);
					DefaultTableModel model = (DefaultTableModel) table_fileinformation.getModel();

					while (model.getRowCount() > 0) {

						for (int i = 0; i < model.getRowCount(); ++i) {

							model.removeRow(i);
						}
					}
					editorPane.setText("");

					DefaultTableModel model1 = (DefaultTableModel) table_1.getModel();

					while (model1.getRowCount() > 0) {

						for (int i = 0; i < model1.getRowCount(); ++i) {

							model1.removeRow(i);
						}
					}
					editorPane.setText("");
					System.gc();
					panel_progress.setVisible(false);
					if (fileoption.equalsIgnoreCase("Exchange Offline Storage (.ost)")
							|| fileoption.equalsIgnoreCase("OLM File (.olm)")
							|| fileoption.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {

						try {
							comboBox_fileDestination_type.removeItem("VCF");
							imageMap_output.remove("VCF");
						} catch (Exception e) {

						}
						try {
							comboBox_fileDestination_type.removeItem("ICS");
							imageMap_output.remove("ICS");
						} catch (Exception e) {

						}

						try {

							try {

								for (int i = 0; i < file_sfd.length; i++) {

									imageMap_output.put(file_sfd[i],
											new ImageIcon(Main_Frame.class.getResource(filesfd_img[i])));

								}
							} catch (Exception ex) {
								ex.printStackTrace();
							}

//							comboBox_fileDestination_type.addItem("VCF");
//							comboBox_fileDestination_type.addItem("ICS");
//
//							imageMap_output.put("VCF", new ImageIcon(Main_Frame.class.getResource("/vcf.png")));
//							imageMap_output.put("ICS", new ImageIcon(Main_Frame.class.getResource("/ics.png")));

						} catch (Exception e1) {

						}

					} else {
						try {

							try {
								comboBox_fileDestination_type.removeItem("VCF");

							} catch (Exception e) {
								e.printStackTrace();
							}
							try {
								comboBox_fileDestination_type.removeItem("ICS");

							} catch (Exception e) {
								e.printStackTrace();
							}

						} catch (Exception e) {
							e.printStackTrace();
						}

					}

					if (fileoption.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

						comboBox_fileDestination_type.removeItem("OST");
					} else if (fileoption.equalsIgnoreCase("EML File (.eml)")) {
						chckbxRemoveDuplicacy.setVisible(false);
						label_Remove_Duplicate.setVisible(false);
						comboBox_fileDestination_type.removeItem("EML");
					} else if (fileoption.equalsIgnoreCase("EMLX File (.emlx)")) {
						chckbxRemoveDuplicacy.setVisible(false);
						label_Remove_Duplicate.setVisible(false);
						comboBox_fileDestination_type.removeItem("EMLX");

					} else if (fileoption.equalsIgnoreCase("OFT File (.oft)")) {
						chckbxRemoveDuplicacy.setVisible(false);
						label_Remove_Duplicate.setVisible(false);
						comboBox_fileDestination_type.removeItem("EMLX");

					} else if (fileoption.equalsIgnoreCase("Message File (.msg)")) {
						chckbxRemoveDuplicacy.setVisible(false);
						label_Remove_Duplicate.setVisible(false);
						comboBox_fileDestination_type.removeItem("MSG");

					} else if (fileoption.equalsIgnoreCase("Maildir")) {
						chckbxRemoveDuplicacy.setVisible(false);
						label_Remove_Duplicate.setVisible(false);

					} else if (fileoption.equalsIgnoreCase("OFFICE 365")) {
						comboBox_fileDestination_type.removeItem("OFFICE 365");
					} else if (fileoption.equalsIgnoreCase("MBOX")) {
						comboBox_fileDestination_type.removeItem("MBOX");
					} else if (fileoption.equalsIgnoreCase("Opera Mail")) {
						comboBox_fileDestination_type.removeItem("Opera Mail");
					} else if (fileoption.equalsIgnoreCase("Thunderbird")) {
						comboBox_fileDestination_type.removeItem("THUNDERBIRD");
					}
					panel_3_2.setVisible(false);

					panel_3_.setVisible(false);
					panel_3_1_1.setVisible(false);

					panel_3_.setVisible(true);

					CardLayout card1 = (CardLayout) panel_3_.getLayout();
					card1.show(panel_3_, "panel_3_1_2");
					panel_3_2.setVisible(true);

					panel_progress.setVisible(true);
					filetype = "PST";
					cal = Calendar.getInstance();
					calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());
					btn_converter.setEnabled(true);
					comboBox_fileDestination_type.setSelectedItem("PST");

					tf_Destination_Location.setText(System.getProperty("user.home") + File.separator + "Desktop");
					if (nextTime != null) {
						filetype = nextfiletype;
						SwingUtilities.invokeLater(new Runnable() {

							public void run() {

								comboBox_fileDestination_type.setSelectedItem(filetype);
							}
						});

						table_fileConvertionreport_panel4.setModel(
								new DefaultTableModel(new Object[][] {}, new String[] { "From", "To", "Status",
										"Duration", "Message Count", "Path", "Last Runtime", "Next RunTime" }));
						for (int i = 0; i < listnextendtime.size(); i++) {
							mode = (DefaultTableModel) table_fileConvertionreport_panel4.getModel();
							Calendar cs = Calendar.getInstance();
							cs.setTimeInMillis(listnextstarttime.get(i));
							Calendar cs1 = Calendar.getInstance();
							cs1.setTimeInMillis(listnextendtime.get(i));

							mode.addRow(new Object[] { fileoption, filetype, Status, listnextduration.get(i),
									listnextcount.get(i), nextPAth, cs.getTime(), cs1.getTime() });

						}

						Starting_Frame.mf.setVisible(true);

						btnConvertAgain.setVisible(false);
						btnDowloadReport.setVisible(false);

						Date date = java.util.Calendar.getInstance().getTime();
						System.out.println(date);
						long start_time = System.currentTimeMillis();
						long difference = nextTime - start_time;

						Calendar cal1 = Calendar.getInstance();
						cal1.setTimeInMillis(nextTime);

						spinner.setValue(cal1.getTime());

						System.out.println("waiting");
						Boolean fonce = Boolean.valueOf(once);
						Boolean feveryday = Boolean.valueOf(everyday);
						Boolean feveryweek = Boolean.valueOf(everyweek);
						Boolean fonweekday = Boolean.valueOf(OnWeekday);
						Boolean fonmonthday = Boolean.valueOf(OnMonthday);
						Boolean feverymonth = Boolean.valueOf(everymonth);
						Boolean removeduplic = Boolean.valueOf(removeduplica);
						Boolean maintainfolder = Boolean.valueOf(maintainfolderh);
						Boolean savepdfatta = Boolean.valueOf(savepdfattac);
						Boolean freeupserverspac = Boolean.valueOf(freeupserverspace);

						chckbxRemoveDuplicacy.setSelected(removeduplic);
						chckbxMaintainFolderHeirachy.setSelected(maintainfolder);
						chckbxDeleteEmailFrom.setSelected(savepdfatta);
						chckbxSavePdfAttachment.setSelected(freeupserverspac);
						System.out.println("once " + fonce);
						System.out.println("everyday " + feveryday);
						System.out.println("everyweek " + feveryweek);
						System.out.println("OnWeekday " + fonweekday);
						System.out.println("OnMonthday " + fonmonthday);
						System.out.println("everymonth " + feverymonth);
						if (fonce) {
							rdbtnOnce.setSelected(true);
						} else if (feveryday) {
							chckbxSetBackupSchedule.setSelected(true);
							rdbtnEveryday.setSelected(true);
						} else if (feveryweek) {
							buttonGroup_Schedulling.clearSelection();
							buttonGroup_Schedulling.setSelected(rdbtnEveryWeek.getModel(), true);
							chckbxSetBackupSchedule.setSelected(true);
						} else if (fonweekday) {
							buttonGroup_Schedulling.clearSelection();
							buttonGroup_Schedulling.setSelected(rdbtnOnWeekDay.getModel(), true);
							chckbxSetBackupSchedule.setSelected(true);
						} else if (fonmonthday) {
							buttonGroup_Schedulling.clearSelection();
							buttonGroup_Schedulling.setSelected(rdbtnOnmonthDay.getModel(), true);
							chckbxSetBackupSchedule.setSelected(true);
						} else if (feverymonth) {
							buttonGroup_Schedulling.clearSelection();
							buttonGroup_Schedulling.setSelected(rdbtnEveryMonth.getModel(), true);
							chckbxSetBackupSchedule.setSelected(true);
						}

						lblNextMigrationStart.setVisible(true);
						long i1 = TimeUnit.MILLISECONDS.toSeconds(difference);
						if (trayIcon != null) {
							trayIcon.displayMessage(
									"Next Migration Start On " + Long.valueOf(i1).toString() + " seconds ", " ",
									TrayIcon.MessageType.NONE);
						}
						if (nexttime) {
							CardLayout card = (CardLayout) Cardlayout.getLayout();
							card.show(Cardlayout, "panel_4");
							while (i1 > 0) {
								if (nextstart) {
									nextstart = false;
									break;
								}

								Long day = TimeUnit.SECONDS.toDays(i1);

								long p1 = i1 % 60;
								long p2 = i1 / 60;
								long p3 = p2 % 60;
								p2 = p2 / 60;

								String s = null;
								if (day > 0) {
									s = "Remaining Days for Next Migration: " + day;
									System.out.print(s);
									lblNextMigrationStart.setText(s);
								} else {
									s = "Remaining Time  for Next Migration " + p2 + " Hrs:" + p3 + " Min:" + p1
											+ " Sec";
									System.out.print(s);
									lblNextMigrationStart.setText(s);
								}
								try {

									Thread.sleep(1000L);
									start_time = System.currentTimeMillis();

									difference = nextTime - start_time;

									System.out.println(difference);

									lblNextMigrationStart.setVisible(true);
									i1 = TimeUnit.MILLISECONDS.toSeconds(difference);
									lblNextMigrationStart.setText(
											"Next Migration Start On : " + Long.valueOf(i1).toString() + " seconds ");
								} catch (InterruptedException e) {

								}
							}
							card = (CardLayout) panel_3_.getLayout();
							card.show(panel_3_, "panel_3_1_2");
						}
						lblNextMigrationStart.setText("Backup about to start ");

						if (trayIcon != null) {
							trayIcon.displayMessage(" ", messageboxtitle + " Backup about to start",
									TrayIcon.MessageType.NONE);
						}

						if (rdbtnOnce.isSelected()) {
							lblNextMigrationStart.setVisible(false);
							chckbxSetBackupSchedule.setSelected(false);
						}
						connectionHandle1();
						btn_converter.setEnabled(true);
						chckbxAutoIncrementBackup.setSelected(true);
						btn_converter.doClick();

						CardLayout card = (CardLayout) Cardlayout.getLayout();
						card.show(Cardlayout, "panel_3");
						lblNextMigrationStart.setText("");

					} else {
						Date startdate = java.util.Calendar.getInstance().getTime();
						dateChooserNextSchedular.setDate(startdate);
						CardLayout card = (CardLayout) Cardlayout.getLayout();
						card.show(Cardlayout, "panel_3");

					}

					chckbxMaintainFolderHeirachy.setSelected(true);
				}

			}
		});

		panel_3 = new JPanel();
		panel_3.setBackground(Color.WHITE);
		Cardlayout.add(panel_3, "panel_3");
		panel_3.setLayout(null);

		btn_signout_p3 = new JButton("");
		btn_signout_p3.setFocusable(false);
		btn_signout_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_signout_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sing-out-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_signout_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sing-out-btn.png")));
			}
		});

		btn_signout_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sing-out-btn.png")));
		btn_signout_p3.setFocusTraversalKeysEnabled(false);
		btn_signout_p3.setFocusPainted(false);
		btn_signout_p3.setDefaultCapable(false);
		btn_signout_p3.setContentAreaFilled(false);
		btn_signout_p3.setBorderPainted(false);
		btn_signout_p3.setToolTipText("Click here to Sign Out of the Server.");
		btn_signout_p3.setRolloverEnabled(false);
		btn_signout_p3.setRequestFocusEnabled(false);
		btn_signout_p3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					String warn = "Do you want to sign out?";
					int ans = JOptionPane.showConfirmDialog(Main_Frame.this, warn, messageboxtitle,
							JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
					if (ans == JOptionPane.YES_OPTION) {

						if (input) {
							if (fileoption.equalsIgnoreCase("OFFICE 365")
									|| fileoption.equalsIgnoreCase("Live Exchange")
									|| fileoption.equalsIgnoreCase("Hotmail")) {
								try {
									clientforexchange_input.dispose();
									clientforexchange_input.dispose();
								} catch (Exception e) {

								}
								textField_domain_name_p3.setText("");
								passwordField_p3.setText("");
								textField_username_p3.setText("");
								textField_domain_name_p3.setText("");
								comboBox_fileDestination_type.setEnabled(true);
								CardLayout card = (CardLayout) panel_3_.getLayout();
								card.show(panel_3_, "panel_3_1_1");

							} else {
								textField_domain_name_p3.setText("");
								passwordField_p3.setText("");
								textField_username_p3.setText("");
								textField_domain_name_p3.setText("");
								comboBox_fileDestination_type.setEnabled(true);
								CardLayout card = (CardLayout) panel_3_.getLayout();
								card.show(panel_3_, "panel_3_1_1");
								try {
									iconnforimap_input.dispose();
									iconnforimap_input.dispose();
								} catch (Exception e) {

								}
							}
						}
						if (output) {
							if (filetype.equalsIgnoreCase("OFFICE 365") || filetype.equalsIgnoreCase("Live Exchange")
									|| filetype.equalsIgnoreCase("Hotmail")) {
								try {
									textField_domain_name_p3.setText("");
									passwordField_p3.setText("");
									textField_username_p3.setText("");
									textField_domain_name_p3.setText("");
									comboBox_fileDestination_type.setEnabled(true);
									CardLayout card = (CardLayout) panel_3_.getLayout();
									card.show(panel_3_, "panel_3_1_1");
									clientforexchange_input.dispose();
									clientforexchange_input.dispose();
								} catch (Exception e) {

								}

							} else {

								try {
									textField_domain_name_p3.setText("");
									passwordField_p3.setText("");
									textField_username_p3.setText("");
									textField_domain_name_p3.setText("");
									comboBox_fileDestination_type.setEnabled(true);
									CardLayout card = (CardLayout) panel_3_.getLayout();
									card.show(panel_3_, "panel_3_1_1");
									iconnforimap_input.dispose();
									iconnforimap_input.dispose();
								} catch (Exception e) {

								}
							}
						}
						textField_domain_name_p3.setText("");
						passwordField_p3.setText("");
						textField_username_p3.setText("");
						textField_domain_name_p3.setText("");
						btn_converter.setEnabled(false);
						btn_converter.setEnabled(false);
						btn_converter.setEnabled(false);
						btn_converter.setEnabled(false);
						btn_signout_p3.setVisible(false);
						comboBox_fileDestination_type.setEnabled(true);

						CardLayout card = (CardLayout) panel_3_.getLayout();
						card.show(panel_3_, "panel_3_1_1");
					}

				} catch (Error e) {
					logger.warning("Error : " + e.getMessage() + System.lineSeparator());
				} catch (Exception e) {
					logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
					return;
				} finally {

					// stop = true;

				}
			}
		});
		btn_previous_p3 = new JButton("");
		btn_previous_p3.setBorderPainted(false);
		btn_previous_p3.setContentAreaFilled(false);
		btn_previous_p3.setDefaultCapable(false);
		btn_previous_p3.setFocusable(false);
		btn_previous_p3.setFocusTraversalKeysEnabled(false);
		btn_previous_p3.setFocusPainted(false);
		btn_previous_p3.setRolloverEnabled(false);
		btn_previous_p3.setRequestFocusEnabled(false);
		btn_previous_p3.setOpaque(false);
		btn_previous_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_previous_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_previous_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
			}
		});

		btn_previous_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/previous-btn.png")));
		btn_previous_p3.setBounds(812, 562, 118, 39);
		panel_3.add(btn_previous_p3);
		btn_previous_p3.setFont(new Font("Tahoma", Font.BOLD, 12));
		btn_previous_p3.setToolTipText("Click here to Go to the Previous Panel. ");
		btn_previous_p3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				lblNew_setemail.setText("");

				lblNew_setsubject.setText("");

				label_date.setText("");
				editorPane.setText("");

				DefaultTableModel model = (DefaultTableModel) table_fileinformation.getModel();

				while (model.getRowCount() > 0) {

					for (int i = 0; i < model.getRowCount(); ++i) {

						model.removeRow(i);
					}
				}

				DefaultTableModel model1 = (DefaultTableModel) table_1.getModel();

				while (model1.getRowCount() > 0) {

					for (int i = 0; i < model1.getRowCount(); ++i) {

						model1.removeRow(i);
					}
				}
				try {
					if (output) {
						if (filetype.equals("OFFICE 365") || filetype.equals("Live Exchange")
								|| filetype.equals("Hotmail")) {
							clientforexchange_output.dispose();

						} else {

							iconnforimap_output.dispose();
						}
					}
				} catch (Exception e1) {

				}
				comboBox_fileDestination_type.setEnabled(true);

				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_2");

			}
		});
		btn_signout_p3.setBounds(912, 13, 142, 38);
		btn_signout_p3.setVisible(false);
		panel_3.add(btn_signout_p3);

		panel_3_ = new JPanel();
		panel_3_.setBounds(12, 61, 1094, 321);
		panel_3.add(panel_3_);
		panel_3_.setLayout(new CardLayout(0, 0));

		panel_3_1_1 = new JPanel();
		panel_3_1_1.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_3_1_1.setBackground(Color.WHITE);
		panel_3_.add(panel_3_1_1, "panel_3_1_1");
		panel_3_1_1.setLayout(null);

		JLabel lblNewLabel = new JLabel("User Name");
		lblNewLabel.setFont(new Font("Tahoma", Font.BOLD, 14));
		lblNewLabel.setBounds(109, 68, 216, 32);
		panel_3_1_1.add(lblNewLabel);

		JLabel lblNewLabel_1 = new JLabel("Password ");
		lblNewLabel_1.setFont(new Font("Tahoma", Font.BOLD, 14));
		lblNewLabel_1.setBounds(109, 111, 192, 25);
		panel_3_1_1.add(lblNewLabel_1);

		panel_3_1_1_1 = new JPanel();
		panel_3_1_1_1.setBackground(Color.WHITE);
		panel_3_1_1_1.setBounds(81, 11, 678, 46);
		panel_3_1_1.add(panel_3_1_1_1);
		panel_3_1_1_1.setLayout(null);

		lbl_Domain = new JLabel("");
		lbl_Domain.setFont(new Font("Tahoma", Font.BOLD, 14));
		lbl_Domain.setBounds(27, 11, 205, 26);
		panel_3_1_1_1.add(lbl_Domain);

		textField_domain_name_p3 = new JTextField();
		textField_domain_name_p3.setHorizontalAlignment(JTextField.CENTER);
		textField_domain_name_p3.setComponentPopupMenu(menu);
		textField_domain_name_p3.setBounds(276, 11, 387, 26);
		panel_3_1_1_1.add(textField_domain_name_p3);
		textField_domain_name_p3.setColumns(10);

		textField_username_p3 = new JTextField();
		textField_username_p3.setHorizontalAlignment(JTextField.CENTER);
		textField_username_p3.setComponentPopupMenu(menu);
		textField_username_p3.setBounds(357, 68, 390, 25);
		panel_3_1_1.add(textField_username_p3);
		textField_username_p3.setColumns(10);

		passwordField_p3 = new JPasswordField();
		passwordField_p3.setHorizontalAlignment(JTextField.CENTER);
		passwordField_p3.setComponentPopupMenu(menu);
		passwordField_p3.setBounds(357, 104, 390, 25);
		panel_3_1_1.add(passwordField_p3);

		chckbxShowPassword_p3 = new JCheckBox("Show Password");
		chckbxShowPassword_p3.setRolloverEnabled(false);
		chckbxShowPassword_p3.setRequestFocusEnabled(false);
		chckbxShowPassword_p3.setOpaque(false);
		chckbxShowPassword_p3.setFocusable(false);
		chckbxShowPassword_p3.setFocusPainted(false);
		chckbxShowPassword_p3.setContentAreaFilled(false);
		chckbxShowPassword_p3.setBackground(Color.WHITE);
		chckbxShowPassword_p3.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					passwordField_p3.setEchoChar((char) 0);
				}

				else {
					passwordField_p3.setEchoChar('●');
				}
			}
		});

		chckbxShowPassword_p3.setFont(new Font("Tahoma", Font.BOLD, 13));
		chckbxShowPassword_p3.setBounds(756, 107, 263, 25);
		panel_3_1_1.add(chckbxShowPassword_p3);

		btn_Sign_p3 = new JButton("");
		btn_Sign_p3.setDefaultCapable(false);
		btn_Sign_p3.setBorderPainted(false);
		btn_Sign_p3.setRolloverEnabled(false);
		btn_Sign_p3.setRequestFocusEnabled(false);
		btn_Sign_p3.setOpaque(false);
		btn_Sign_p3.setFocusable(false);
		btn_Sign_p3.setFocusTraversalKeysEnabled(false);
		btn_Sign_p3.setFocusPainted(false);
		btn_Sign_p3.setContentAreaFilled(false);
		btn_Sign_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_Sign_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_Sign_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-btn.png")));
			}
		});

		btn_Sign_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/sign-in-btn.png")));

		btn_Sign_p3.setBounds(470, 194, 118, 39);
		panel_3_1_1.add(btn_Sign_p3);
		btn_Sign_p3.setToolTipText("Click here to Sign In the Server. ");
		btn_Sign_p3.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				username_p3 = "";
				password_p3 = "";
				domain_p3 = "";
				comboBox_fileDestination_type.setEnabled(false);
				textField_domain_name_p3.setEnabled(false);
				textField_username_p3.setEnabled(false);
				passwordField_p3.setEnabled(false);
				tf_portNo_p3.setEnabled(false);
				btn_converter.setEnabled(false);

				chckbxShowPassword_p3.setEnabled(false);
				btn_Sign_p3.setEnabled(false);
				try {
					domain_p3 = textField_domain_name_p3.getText().replaceAll("//s", "");
					domain_p3 = domain_p3.trim();
				} catch (Exception a) {
					domain_p3 = "";
				}
				try {
					username_p3 = textField_username_p3.getText().replaceAll("//s", "");
					username_p3 = username_p3.trim();
				} catch (Exception a) {
					username_p3 = "";
				}
				try {
					password_p3 = new String(passwordField_p3.getPassword());
					password_p3 = password_p3.trim();
				} catch (Exception a) {
					password_p3 = "";
				}
				try {
					portnofiletype = Integer.parseInt(tf_portNo_p3.getText().replaceAll("//s", ""));

				} catch (Exception a) {

				}
				if (username_p3.equalsIgnoreCase("") || password_p3.equalsIgnoreCase("")) {

					if (username_p3.equalsIgnoreCase("") && password_p3.equalsIgnoreCase("")) {
						JOptionPane.showMessageDialog(frame, "User name and Password fields can't be empty",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					} else if (username_p3.equalsIgnoreCase("")) {

						JOptionPane.showMessageDialog(frame, "User name field can't be empty", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					} else if (password_p3.equalsIgnoreCase("")) {

						JOptionPane.showMessageDialog(frame, "Password field can't be empty", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

					}

					comboBox_fileDestination_type.setEnabled(true);
					chckbxShowPassword_p3.setEnabled(true);
					btn_previous_p3.setEnabled(true);
					textField_domain_name_p3.setEnabled(true);
					textField_username_p3.setEnabled(true);
					passwordField_p3.setEnabled(true);
					// tf_portNo_p3.setVisible(false);
					btn_Sign_p3.setEnabled(true);
					btn_signout_p3.setVisible(false);
					chckbx_Mail_Filter.setEnabled(true);
					task_box.setEnabled(true);

				} else if (filetype.equalsIgnoreCase("Live Exchange") && domain_p3.equalsIgnoreCase("")) {

					JOptionPane.showMessageDialog(frame, "Computer Name or IP Address field can not be empty",
							messageboxtitle, JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));
					comboBox_FiletypeChooser.setEnabled(true);
					btn_Previous_pane2.setEnabled(true);
					btn_next_pane2.setEnabled(true);
					// tf_portNo_p3.setVisible(true);
					chckbxShowPassword_p2.setEnabled(true);
					btn_SignIn.setEnabled(true);

				} else if (filetype.equalsIgnoreCase("IMAP") && domain_p3.equalsIgnoreCase("")) {

					JOptionPane.showMessageDialog(frame, "IMAP Host field can't be empty", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					comboBox_FiletypeChooser.setEnabled(true);
					btn_Previous_pane2.setEnabled(true);
					btn_next_pane2.setEnabled(true);
					// tf_portNo_p3.setVisible(true);
					chckbxShowPassword_p2.setEnabled(true);
					btn_SignIn.setEnabled(true);

				} else if (filetype.equalsIgnoreCase("IMAP") && tf_portNo_p3.getText().isEmpty()) {

					JOptionPane.showMessageDialog(frame, "Port No field can't be empty", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					comboBox_FiletypeChooser.setEnabled(true);
					btn_Previous_pane2.setEnabled(true);
					btn_next_pane2.setEnabled(true);

					chckbxShowPassword_p2.setEnabled(true);
					btn_SignIn.setEnabled(true);

				} else if (!(isValid(username_p3))) {

					JOptionPane.showMessageDialog(frame, "Please enter a valid username", messageboxtitle,
							JOptionPane.ERROR_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));
					comboBox_fileDestination_type.setEnabled(true);
					chckbxShowPassword_p3.setEnabled(true);
					btn_previous_p3.setEnabled(true);
					textField_domain_name_p3.setEnabled(true);
					textField_username_p3.setEnabled(true);
					passwordField_p3.setEnabled(true);
					tf_portNo_p3.setEnabled(true);
					btn_Sign_p3.setEnabled(true);
					btn_signout_p3.setVisible(false);
					// tf_portNo_p3.setVisible(true);

				} else {

					th = new Thread(new Runnable() {

						@Override
						public void run() {
							lbl_connecting_p3.setVisible(true);

							try {
								btn_converter.setEnabled(false);
								btn_previous_p3.setEnabled(false);

								if (filetype.equalsIgnoreCase("OFFICE 365")) {
									conntiontooffice365_output();
								}
								if (filetype.equalsIgnoreCase("HOTMAIL")) {
									conntiontohotmail_output();

								} else if (filetype.equalsIgnoreCase("Yandex Mail")) {
									connectiontoYandex_output();

								} else if (filetype.equalsIgnoreCase("Zoho Mail")) {
									connectiontozoho_output();

								} else if (filetype.equalsIgnoreCase("Amazon WorkMail")) {
									connectiontoinaws_output();

								} else if (filetype.equalsIgnoreCase("Hostgator email")) {
									connectiontoHostgator_output();

								} else if (filetype.equalsIgnoreCase("Icloud")) {
									connectiontoicloud_output();

								} else if (filetype.equalsIgnoreCase("GoDaddy email")) {
									connectiontoGoDaddy_output();

								} else if (filetype.equalsIgnoreCase("GMAIL")) {
									connectiontogmail_output();

								} else if (filetype.equalsIgnoreCase("Live Exchange")) {
									connectionwithexchangeserver_output();

								} else if (filetype.equalsIgnoreCase("IMAP")) {
									connectiontoimap_output();
								} else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {
									connectiontoyahoo_output();

								} else if (filetype.equalsIgnoreCase("AOL")) {
									connectiontoaol_output();

								}

								panel_3_1_1.setVisible(false);
								panel_3_1_2.setVisible(true);
								btn_converter.setEnabled(true);
								lblEnableImap_p3.setEnabled(true);
								btn_signout_p3.setVisible(true);
								btn_signout_p3.setVisible(true);
								CardLayout card = (CardLayout) panel_3_.getLayout();
								card.show(panel_3_, "panel_3_1_2");
							} catch (Error e) {
								logger.warning("Error : " + e.getMessage() + System.lineSeparator());
							} catch (Exception e) {
								logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
								lblLive_Chat_p3.setVisible(true);
								btn_signout_p3.setVisible(false);
								btn_signout_p3.setEnabled(false);

								comboBox_fileDestination_type.setEnabled(true);
								System.out.println(e.getMessage());

								if (filetype.equalsIgnoreCase("Gmail")) {
									if (e.getMessage().equalsIgnoreCase(
											"AE_1_2_0002 NO [AUTHENTICATIONFAILED] Invalid credentials (Failure)")) {
										JOptionPane.showMessageDialog(frame,
												"Connection Not Estalished with Gmail please check your Credantial OR Otherwise allow 3rd party app to acess your account",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else if (e.getMessage().contains(" Application-specific password required: ")) {
										JOptionPane.showMessageDialog(frame, "Application specific password required",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else {
										JOptionPane.showMessageDialog(frame, "Connection not established.",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									}
								} else if (filetype.equalsIgnoreCase("Yahoo Mail")) {
									if (e.getMessage().equalsIgnoreCase(
											"AE_3_2_0002 NO [AUTHORIZATIONFAILED] LOGIN Invalid credentials")) {
										JOptionPane.showMessageDialog(frame,
												"Connection Not Estalished with Yahoo Mail please check your Credantial Otherwise allow 3rd party app to acess your account",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else if (e.getMessage().contains(" Application-specific password required: ")) {
										JOptionPane.showMessageDialog(frame, "Application specific password required",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									} else {
										JOptionPane.showMessageDialog(frame, "Connection not established.",
												messageboxtitle, JOptionPane.ERROR_MESSAGE,
												new ImageIcon(Main_Frame.class.getResource("/information.png")));
									}
								} else if (e.getMessage().contains(" Application-specific password required: ")) {
									JOptionPane.showMessageDialog(frame, "Application specific password required",
											messageboxtitle, JOptionPane.ERROR_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/information.png")));
								} else {
									JOptionPane.showMessageDialog(frame, "Connection not established.", messageboxtitle,
											JOptionPane.ERROR_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/information.png")));
								}

							} finally {
								lbl_connecting_p3.setVisible(false);
								// tf_portNo_p3.setEnabled(true);
								chckbxShowPassword_p3.setEnabled(true);
								btn_previous_p3.setEnabled(true);
								textField_domain_name_p3.setEnabled(true);
								textField_username_p3.setEnabled(true);
								passwordField_p3.setEnabled(true);
								tf_portNo_p3.setEnabled(true);

								btn_Sign_p3.setEnabled(true);

							}

						}
					});
					th.start();

				}
			}
		});
		btn_Sign_p3.setFont(new Font("Tahoma", Font.BOLD, 14));

		lbl_connecting_p3 = new JLabel("");
		lbl_connecting_p3.setBounds(370, 194, 85, 32);
		panel_3_1_1.add(lbl_connecting_p3);
		lbl_connecting_p3.setIcon(new ImageIcon(Main_Frame.class.getResource("/loading.gif")));

		tf_portNo_p3 = new JTextField();
		tf_portNo_p3.setHorizontalAlignment(JTextField.CENTER);
		tf_portNo_p3.setComponentPopupMenu(menu);
		tf_portNo_p3.setBounds(357, 147, 390, 25);
		tf_portNo_p3.setText(Integer.toString(993));
		panel_3_1_1.add(tf_portNo_p3);
		tf_portNo_p3.setColumns(10);

		lblPortNo = new JLabel("Port No.");
		lblPortNo.setFont(new Font("Tahoma", Font.BOLD, 14));
		lblPortNo.setBounds(109, 157, 192, 27);
		panel_3_1_1.add(lblPortNo);

		lblLive_Chat_p3 = new JLabel("More Help");
		lblLive_Chat_p3.setForeground(Color.RED);
		lblLive_Chat_p3.setCursor(cursor);
		lblLive_Chat_p3.setFont(new Font("Tahoma", Font.PLAIN, 14));
		lblLive_Chat_p3.setBounds(935, 11, 71, 25);
		lblLive_Chat_p3.addMouseListener(new MouseAdapter() {

			public void mouseClicked(MouseEvent e) {
				openBrowser("http://messenger.providesupport.com/messenger/0pi295uz3ga080c7lxqxxuaoxr.html");
			}
		});

		panel_3_1_1.add(lblLive_Chat_p3);

		panel = new JPanel();
		panel.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel.setBackground(Color.WHITE);
		panel.setBounds(31, 258, 937, 52);
		panel_3_1_1.add(panel);
		panel.setLayout(null);

		lblMakeSureYou = new JLabel("Please  Click on The Link");
		lblMakeSureYou.setForeground(Color.BLACK);
		lblMakeSureYou.setBounds(10, 11, 162, 32);
		panel.add(lblMakeSureYou);
		lblMakeSureYou.setFont(new Font("Tahoma", Font.PLAIN, 13));

		lblEnableImap_p3 = new JLabel("<HTML><U>To Enable IMAP</U><HTML>");
		lblEnableImap_p3.setBounds(820, 17, 84, 25);
		lblEnableImap_p3.setCursor(cursor);
		panel.add(lblEnableImap_p3);
		lblEnableImap_p3.setForeground(Color.BLUE);
		lblEnableImap_p3.setFont(new Font("Tahoma", Font.PLAIN, 11));

		lblTurnOffTwo_p3 = new JLabel("<HTML><U>Turn Off Two Step Verification</U></HTML>");
		lblTurnOffTwo_p3.setBounds(182, 11, 628, 32);
		lblTurnOffTwo_p3.setCursor(cursor);
		panel.add(lblTurnOffTwo_p3);
		lblTurnOffTwo_p3.setForeground(Color.BLUE);
		lblTurnOffTwo_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				openbrowserturntwostepoff(filetype);

			}
		});
		lblTurnOffTwo_p3.setFont(new Font("Tahoma", Font.PLAIN, 11));

		lblthirdPartyPassword_1 = new JLabel("(Use third party App Password)");
		lblthirdPartyPassword_1.setForeground(Color.RED);
		lblthirdPartyPassword_1.setBounds(69, 137, 278, 14);
		panel_3_1_1.add(lblthirdPartyPassword_1);

		lblemailAddress_1 = new JLabel("(Email Address)");
		lblemailAddress_1.setForeground(Color.RED);
		lblemailAddress_1.setBounds(109, 93, 216, 25);
		panel_3_1_1.add(lblemailAddress_1);
		lblEnableImap_p3.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				openbrowserenableimap(filetype);

			}
		});
		lbl_connecting_p3.setVisible(false);

		panel_3_1_2 = new JPanel();
		panel_3_1_2.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_3_1_2.setBackground(Color.WHITE);
		panel_3_.add(panel_3_1_2, "panel_3_1_2");

		panel_3_2 = new JPanel();
		panel_3_2.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_3_2.setBackground(Color.WHITE);
		panel_3_2.setBounds(12, 395, 1053, 48);
		panel_3.add(panel_3_2, "panel_3_2");
		panel_3_2.setLayout(null);
		panel_3_2.setVisible(false);

		panel_mailfilter = new JPanel();
		panel_mailfilter.setBounds(1076, 140, 18, 32);
		panel_mailfilter.setVisible(false);
		panel_3_1_2.setLayout(null);

		DateFilter = new JCheckBox("Date Filter");
		DateFilter.setBounds(511, 5, 85, 23);
		DateFilter.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (DateFilter.isSelected()) {
					dateChooser_newFrom.setEnabled(true);
					dateChooser_newTo.setEnabled(true);
					removeButton.setEnabled(true);
					addButton.setEnabled(true);
				} else {
					dateChooser_newFrom.setEnabled(false);
					dateChooser_newTo.setEnabled(false);
					removeButton.setEnabled(false);
					addButton.setEnabled(false);
				}
			}
		});
		DateFilter.setFont(new Font("Tahoma", Font.BOLD, 11));
		DateFilter.setBackground(Color.WHITE);
		panel_3_1_2.add(DateFilter);
		panel_mailfilter.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_mailfilter.setBackground(Color.WHITE);
		panel_3_1_2.add(panel_mailfilter);
		panel_mailfilter.setLayout(null);
		panel_mailfilter.setEnabled(false);
		panel_mailfilter.setVisible(false);

		JLabel label_1 = new JLabel("Start Date");
		label_1.setFont(new Font("Tahoma", Font.PLAIN, 11));
		label_1.setBounds(10, 11, 24, 20);
		panel_mailfilter.add(label_1);

		JLabel label_3 = new JLabel("End Date");
		label_3.setFont(new Font("Tahoma", Font.PLAIN, 11));
		label_3.setBounds(54, 12, 24, 19);
		panel_mailfilter.add(label_3);

		dateChooser_mail_fromdate = new JDateChooser();
		dateChooser_mail_fromdate.setVisible(false);
		dateChooser_mail_fromdate.addMouseListener(new MouseAdapter() {

			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_mail_fromdate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_mail_fromdate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_mail_fromdate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_mail_fromdate.setBackground(Color.WHITE);
		dateChooser_mail_fromdate.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				Calendar cal2 = Calendar.getInstance();

				cal2.set(Calendar.HOUR_OF_DAY, 00);
				cal2.set(Calendar.MINUTE, 00);
				cal2.set(Calendar.SECOND, 00);
				Date startdate = cal2.getTime();
				dateChooser_mail_fromdate.setMaxSelectableDate(startdate);
			}
		});
		dateChooser_mail_fromdate.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));
		dateChooser_mail_fromdate.setBounds(102, 11, 17, 23);
		dateChooser_mail_fromdate.setEnabled(false);
		panel_mailfilter.add(dateChooser_mail_fromdate);

		dateChooser_mail_tilldate = new JDateChooser();
		dateChooser_mail_fromdate.setVisible(false);
		dateChooser_mail_tilldate.getCalendarButton().setBackground(Color.WHITE);
		dateChooser_mail_tilldate.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_mail_tilldate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_mail_tilldate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_mail_tilldate.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_mail_tilldate.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal3 = Calendar.getInstance();
				cal3.set(Calendar.HOUR_OF_DAY, 23);
				cal3.set(Calendar.MINUTE, 59);
				cal3.set(Calendar.SECOND, 59);
				Date enddate = cal3.getTime();
				dateChooser_mail_tilldate.setMaxSelectableDate(enddate);
				try {
					Calendar calendarstartdate = dateChooser_mail_fromdate.getCalendar();
					calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
					calendarstartdate.set(Calendar.MINUTE, 00);
					calendarstartdate.set(Calendar.SECOND, 00);
					dateChooser_mail_tilldate.setMinSelectableDate(calendarstartdate.getTime());
				} catch (Error e1) {
					logger.warning("Error : " + e1.getMessage() + System.lineSeparator());
				} catch (Exception e1) {
					logger.warning("Exception : " + e1.getMessage() + System.lineSeparator());
					return;
				}
			}
		});
		dateChooser_mail_tilldate.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));
		dateChooser_mail_tilldate.setBounds(129, 11, 24, 19);
		dateChooser_mail_tilldate.setEnabled(false);
		panel_mailfilter.add(dateChooser_mail_tilldate);

		chckbx_Mail_Filter = new JCheckBox("Mail Filter");
		chckbx_Mail_Filter.setBounds(159, 4, 29, 32);
		panel_mailfilter.add(chckbx_Mail_Filter);
		chckbx_Mail_Filter.setRolloverEnabled(false);
		chckbx_Mail_Filter.setRequestFocusEnabled(false);
		chckbx_Mail_Filter.setOpaque(false);
		chckbx_Mail_Filter.setFocusable(false);
		chckbx_Mail_Filter.setFocusPainted(false);
		chckbx_Mail_Filter.setContentAreaFilled(false);
		chckbx_Mail_Filter.setVisible(false);
		chckbx_Mail_Filter.setFont(new Font("Tahoma", Font.BOLD, 12));
		chckbx_Mail_Filter.setBackground(Color.WHITE);
		chckbx_Mail_Filter.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					panel_mailfilter.setEnabled(true);
					dateChooser_mail_fromdate.setEnabled(true);
					dateChooser_mail_tilldate.setEnabled(true);
				}

				else {
					panel_mailfilter.setEnabled(false);
					dateChooser_mail_fromdate.setEnabled(false);
					dateChooser_mail_tilldate.setEnabled(false);
				}

			}
		});

		panel_Calender = new JPanel();
		panel_Calender.setBounds(34, 108, -12, 18);
		panel_Calender.setBackground(Color.WHITE);
		panel_Calender.setBorder(new TitledBorder(UIManager.getBorder("TitledBorder.border"), "Calender Filter",
				TitledBorder.CENTER, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel_Calender.setEnabled(false);
		panel_Calender.setVisible(false);
		panel_3_1_2.add(panel_Calender);
		panel_Calender.setLayout(null);

		JDateChooser dateChooser_calender_start = new JDateChooser();
		dateChooser_calender_start.setBackground(Color.WHITE);
		dateChooser_calender_start.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_calender_start.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_calender_start.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_calender_start.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_calender_start.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal2 = Calendar.getInstance();

				cal2.set(Calendar.HOUR_OF_DAY, 00);
				cal2.set(Calendar.MINUTE, 00);
				cal2.set(Calendar.SECOND, 00);
				Date startdate = cal2.getTime();
				dateChooser_calender_start.setMaxSelectableDate(startdate);
			}
		});
		dateChooser_calender_start.setBounds(169, 19, 168, 20);
		panel_Calender.add(dateChooser_calender_start);
		dateChooser_calender_start.setEnabled(false);
		dateChooser_calender_start.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));

		JDateChooser dateChooser_calendar_end = new JDateChooser();
		dateChooser_calendar_end.setBackground(Color.WHITE);
		dateChooser_calendar_end.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_calendar_end.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_calendar_end.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_calendar_end.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_calendar_end.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal3 = Calendar.getInstance();
				cal3.set(Calendar.HOUR_OF_DAY, 23);
				cal3.set(Calendar.MINUTE, 59);
				cal3.set(Calendar.SECOND, 59);
				Date enddate = cal3.getTime();
				dateChooser_calendar_end.setMaxSelectableDate(enddate);

				try {
					Calendar calendarstartdate = dateChooser_calender_start.getCalendar();
					calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
					calendarstartdate.set(Calendar.MINUTE, 00);
					calendarstartdate.set(Calendar.SECOND, 00);
					dateChooser_calendar_end.setMinSelectableDate(calendarstartdate.getTime());
				} catch (Error e1) {
					logger.warning("Error : " + e1.getMessage() + System.lineSeparator());
				}

				catch (Exception e1) {
					logger.warning("Exceptin :" + e1.getMessage() + System.lineSeparator());
					return;
				}
			}
		});
		dateChooser_calendar_end.setBounds(808, 28, 149, 19);
		panel_Calender.add(dateChooser_calendar_end);
		dateChooser_calendar_end.setEnabled(false);
		dateChooser_calendar_end.getCalendarButton().setFont(new Font("Tahoma", Font.BOLD, 15));

		JLabel label = new JLabel("End Date\r\n");
		label.setFont(new Font("Tahoma", Font.BOLD, 15));
		label.setBounds(695, 28, 102, 19);
		panel_Calender.add(label);

		JLabel label_2 = new JLabel("Start Date");
		label_2.setFont(new Font("Tahoma", Font.BOLD, 15));
		label_2.setBounds(76, 19, 114, 20);
		panel_Calender.add(label_2);

		JPanel panel_7 = new JPanel();
		panel_7.setBounds(39, 94, 1, 11);
		panel_7.setBackground(Color.WHITE);
		panel_7.setVisible(false);
		panel_3_1_2.add(panel_7);
		panel_7.setLayout(null);

		chckbx_calender_box = new JCheckBox("Calender Filter");
		chckbx_calender_box.setBackground(Color.WHITE);
		chckbx_calender_box.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					panel_Calender.setEnabled(true);
					dateChooser_calendar_end.setEnabled(true);
					dateChooser_calender_start.setEnabled(true);
				} else {
					panel_Calender.setEnabled(false);
					dateChooser_calendar_end.setEnabled(false);
					dateChooser_calender_start.setEnabled(false);
				}
			}
		});
		chckbx_calender_box.setBounds(8, 0, 0, 25);
		panel_7.add(chckbx_calender_box);

		panel_taskfilter = new JPanel();
		panel_taskfilter.setBounds(1081, 45, 13, 32);
		panel_taskfilter.setVisible(false);
		panel_taskfilter.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_taskfilter.setBackground(Color.WHITE);
		panel_3_1_2.add(panel_taskfilter);
		panel_taskfilter.setLayout(null);

		dateChooser_task_start_date = new JDateChooser();
		dateChooser_task_start_date.setBounds(104, 9, 23, 22);
		panel_taskfilter.add(dateChooser_task_start_date);
		dateChooser_task_start_date.setBackground(Color.WHITE);
		dateChooser_task_start_date.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_task_start_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_task_start_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_task_start_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_task_start_date.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal2 = Calendar.getInstance();
				cal2.set(Calendar.HOUR_OF_DAY, 00);
				cal2.set(Calendar.MINUTE, 00);
				cal2.set(Calendar.SECOND, 00);
				Date startdate = cal2.getTime();
				dateChooser_task_start_date.setMaxSelectableDate(startdate);

			}
		});
		dateChooser_task_start_date.setEnabled(false);

		JLabel label_4 = new JLabel("Start Date");
		label_4.setBounds(191, 8, 23, 20);
		panel_taskfilter.add(label_4);
		label_4.setFont(new Font("Tahoma", Font.PLAIN, 11));

		JLabel label_5 = new JLabel("End Date\r\n");
		label_5.setBounds(224, 9, 23, 19);
		panel_taskfilter.add(label_5);
		label_5.setFont(new Font("Tahoma", Font.PLAIN, 11));

		dateChooser_task_end_date = new JDateChooser();
		dateChooser_task_end_date.setBounds(137, 12, 23, 19);
		panel_taskfilter.add(dateChooser_task_end_date);
		dateChooser_task_end_date.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_task_end_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_task_end_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_task_end_date.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_task_end_date.setBackground(Color.WHITE);
		dateChooser_task_end_date.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				Calendar cal3 = Calendar.getInstance();
				cal3.set(Calendar.HOUR_OF_DAY, 23);
				cal3.set(Calendar.MINUTE, 59);
				cal3.set(Calendar.SECOND, 59);
				Date enddate = cal3.getTime();
				dateChooser_task_end_date.setMaxSelectableDate(enddate);

				try {
					Calendar calendarstartdate = dateChooser_task_start_date.getCalendar();
					calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
					calendarstartdate.set(Calendar.MINUTE, 00);
					calendarstartdate.set(Calendar.SECOND, 00);
					dateChooser_task_end_date.setMinSelectableDate(calendarstartdate.getTime());
				} catch (Error e1) {
					logger.warning("Error : " + e1.getMessage() + System.lineSeparator());
				} catch (Exception e1) {
					logger.warning("Exception : " + e1.getMessage() + System.lineSeparator());
					return;
				}

			}
		});
		dateChooser_task_end_date.setEnabled(false);

		task_box = new JCheckBox("Task Filter");
		task_box.setBounds(69, 9, 29, 18);
		panel_taskfilter.add(task_box);
		task_box.setRolloverEnabled(false);
		task_box.setRequestFocusEnabled(false);
		task_box.setOpaque(false);
		task_box.setFocusable(false);
		task_box.setFocusPainted(false);
		task_box.setContentAreaFilled(false);
		task_box.setFont(new Font("Tahoma", Font.BOLD, 12));
		task_box.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {

					dateChooser_task_end_date.setEnabled(true);
					dateChooser_task_start_date.setEnabled(true);
				} else {

					dateChooser_task_end_date.setEnabled(false);
					dateChooser_task_start_date.setEnabled(false);
				}
			}
		});
		task_box.setBackground(Color.WHITE);

		JPanel panel_5 = new JPanel();
		panel_5.setBounds(518, 258, 498, 48);
		panel_5.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_5.setBackground(Color.WHITE);
		panel_5.setVisible(false);
		panel_3_1_2.add(panel_5);
		panel_5.setLayout(null);

		comboBox = new JComboBox();
		comboBox.setBounds(170, 14, 242, 21);
		panel_5.add(comboBox);
		comboBox.addItem("Subject");
		comboBox.addItem("Subject_Date(DD-MM-YYYY)");
		comboBox.addItem("Subject_Date(MM-DD-YYYY)");
		comboBox.addItem("Subject_Date(YYYY-MM-DD)");
		comboBox.addItem("Subject_Date(YYYY-DD-MM)");
		comboBox.addItem("(DD-MM-YYYY)Date_Subject");
		comboBox.addItem("(MM-DD-YYYY)Date_Subject");
		comboBox.addItem("(YYYY-MM-DD)Date_Subject");
		comboBox.addItem("(YYYY-DD-MM)Date_Subject");
		comboBox.addItem("From_Subject_Date(DD-MM-YYYY)");
		comboBox.addItem("From_Subject_Date(MM-DD-YYYY)");
		comboBox.addItem("From_Subject_Date(YYYY-MM-DD)");
		comboBox.addItem("From_Subject_Date(YYYY-DD-MM)");
		comboBox.addItem("(DD-MM-YYYY)Date_From_Subject");
		comboBox.addItem("(MM-DD-YYYY)Date_From_Subject");
		comboBox.addItem("(YYYY-MM-DD)Date_From_Subject");
		comboBox.addItem("(YYYY-DD-MM)Date_From_Subject");
		comboBox.setVisible(false);

		// ----------------------------------------Enterprice----------------------------------------------------//


		if (versiontype == 4) {
//			PST, MBOX, EML, EMLX, MSG, O365, GMAIL, YAHOO, AOL, HOTMAIL, IMAP, ZOHO,
//			THUNDERBIRD, YANDEX, ICLOUD, PDF, CSV, GIF, JPG, TIFF, HTML, MHTML, PNG, DOC, DOCX, DOCM
			
			file_sfd = new String[] { "PST", "MBOX", "EML", "EMLX", "MSG", "PDF", "CSV", "GIF",
						"JPG", "TIFF", "HTML", "MHTML", "PNG", "DOC", "DOCX", "DOCM", "VCF", "ICS" };

			filesfd_img = new String[] { "/pst-cb.png", "/mbox-cb.png", "/eml-cb.png", "/emlx-cb.png", "/msg-cb.png",
						"/pdf-cb.png", "/csv-cb.png", "/gif-cb.png", "/jpg-cb.png", "/tiff-cb.png", "/html-cb.png",
						"/mhtml-cb.png", "/png-cb.png", "/doc-cb.png", "/docx-cb.png", "/docm-cb.png", "/vcf.png",
						"/ics.png" };	
			
			
			
				email_sfd = new String[] { "OFFICE 365", "GMAIL", "G-SUITE", "YAHOO MAIL",
						"THUNDERBIRD", "AOL", "HOTMAIL", "IMAP", "Zoho Mail", "Yandex Mail", "Icloud"};

				emailsfd_img = new String[] {"/office365-cb.png", "/gmail-cb.png","/gsuite-cb.png", "/yahoo-cb.png", "/thunderbird-cb.png", "/aol-cb.png",
						"/hotmail-cb.png", "/imap-cb.png", "/zoho-cb.png", "/yandex-cb.png", "/icloud-cb.png" };

			
			
//			sfd = new String[] { "PST", "MBOX", "EML", "EMLX", "MSG", "OFFICE 365", "GMAIL", "YAHOO MAIL",
//					"THUNDERBIRD", "AOL", "HOTMAIL", "IMAP", "Zoho Mail", "Yandex Mail", "Icloud", "PDF", "CSV", "GIF",
//					"JPG", "TIFF", "HTML", "MHTML", "PNG", "DOC", "DOCX", "DOCM", "VCF", "ICS" };
//
//			sfd_img = new String[] { "/pst-cb.png", "/mbox-cb.png", "/eml-cb.png", "/emlx-cb.png", "/msg-cb.png",
//					"/office365-cb.png", "/gmail-cb.png", "/yahoo-cb.png", "/thunderbird-cb.png", "/aol-cb.png",
//					"/hotmail-cb.png", "/imap-cb.png", "/zoho-cb.png", "/yandex-cb.png", "/icloud-cb.png",
//					"/pdf-cb.png", "/csv-cb.png", "/gif-cb.png", "/jpg-cb.png", "/tiff-cb.png", "/html-cb.png",
//					"/mhtml-cb.png", "/png-cb.png", "/doc-cb.png", "/docx-cb.png", "/docm-cb.png", "/vcf.png",
//					"/ics.png" };

		}

		// ----------------------------------------Technician----------------------------------------------------//

		if (Main_Frame.versiontype == 3) {
//			PST, MBOX, EML, EMLX, MSG, O365, GMAIL, YAHOO, AOL, HOTMAIL, IMAP, THUNDERBIRD, PDF, CSV, GIF,
//			JPG, TIFF, HTML, MHTML, PNG, DOC, DOCX, DOCM
			
			
			email_sfd = new String[] {"OFFICE 365", "GMAIL","G-SUITE", "YAHOO MAIL","THUNDERBIRD", "AOL", "HOTMAIL", "IMAP" };
			emailsfd_img = new String[] {"/office365-cb.png", "/gmail-cb.png","/gsuite-cb.png", "/yahoo-cb.png", "/thunderbird-cb.png", "/aol-cb.png",
						"/hotmail-cb.png", "/imap-cb.png"};
			
			
				file_sfd = new String[] { "PST", "MBOX", "EML", "EMLX", "MSG", 
						 "PDF", "CSV", "GIF", "JPG", "TIFF", "HTML", "MHTML", "PNG",
						"DOC", "DOCX", "DOCM", "VCF", "ICS" };
				filesfd_img = new String[] { "/pst-cb.png", "/mbox-cb.png", "/eml-cb.png", "/emlx-cb.png", "/msg-cb.png",
						 "/pdf-cb.png", "/csv-cb.png", "/gif-cb.png", "/jpg-cb.png",
						"/tiff-cb.png", "/html-cb.png", "/mhtml-cb.png", "/png-cb.png", "/doc-cb.png", "/docx-cb.png",
						"/docm-cb.png", "/vcf.png", "/ics.png" };
			
//			sfd = new String[] { "PST", "MBOX", "EML", "EMLX", "MSG", "OFFICE 365", "GMAIL", "YAHOO MAIL",
//					"THUNDERBIRD", "AOL", "HOTMAIL", "IMAP", "PDF", "CSV", "GIF", "JPG", "TIFF", "HTML", "MHTML", "PNG",
//					"DOC", "DOCX", "DOCM", "VCF", "ICS" };
//
//			sfd_img = new String[] { "/pst-cb.png", "/mbox-cb.png", "/eml-cb.png", "/emlx-cb.png", "/msg-cb.png",
//					"/office365-cb.png", "/gmail-cb.png", "/yahoo-cb.png", "/thunderbird-cb.png", "/aol-cb.png",
//					"/hotmail-cb.png", "/imap-cb.png", "/pdf-cb.png", "/csv-cb.png", "/gif-cb.png", "/jpg-cb.png",
//					"/tiff-cb.png", "/html-cb.png", "/mhtml-cb.png", "/png-cb.png", "/doc-cb.png", "/docx-cb.png",
//					"/docm-cb.png", "/vcf.png", "/ics.png" };

		}
		// ----------------------------------Admin----------------------------------------------------//

		else if (versiontype == 2) {
		
				file_sfd = new String[] { "PST", "MBOX", "EML", "EMLX", "MSG",  "PDF", "CSV",
						"GIF", "JPG", "TIFF", "HTML", "MHTML", "PNG", "DOC", "DOCX", "DOCM", "VCF", "ICS" };
				filesfd_img = new String[] { "/pst-cb.png", "/mbox-cb.png", "/eml-cb.png", "/emlx-cb.png", "/msg-cb.png",
						"/pdf-cb.png", "/csv-cb.png", "/gif-cb.png",
						"/jpg-cb.png", "/tiff-cb.png", "/html-cb.png", "/mhtml-cb.png", "/png-cb.png", "/doc-cb.png",
						"/docx-cb.png", "/docm-cb.png", "/vcf.png", "/ics.png" };
			
//				PST, MBOX, EML, EMLX, MSG, O365, GMAIL, YAHOO, PDF, CSV, GIF, JPG, TIFF, HTML, MHTML, PNG, DOC, DOCX, DOCM
			
			email_sfd = new String[] {  "OFFICE 365", "GMAIL","G-SUITE", "YAHOO MAIL" };

			emailsfd_img = new String[] {"/office365-cb.png", "/gmail-cb.png","/gsuite-cb.png", "/yahoo-cb.png",};
	
		}
		// ----------------------------------------SingleUser----------------------------------------------------//
		else if (versiontype == 1) {
//			PST, MBOX, EML, EMLX, PDF, CSV
			
				file_sfd = new String[] { "PST", "MBOX", "EML", "EMLX", "PDF", "CSV" };

				filesfd_img = new String[] { "/pst-cb.png", "/mbox-cb.png", "/eml-cb.png", "/emlx-cb.png", "/pdf-cb.png",
					"/csv-cb.png" };
			

		}
		

		l_output = new DefaultComboBoxModel<>();
		for (int i1 = 0; i1 < file_sfd.length; i1++) {

			l_output.addElement(file_sfd[i1]);

		}

		comboBox_fileDestination_type = new JComboBox();

		imageMap_output = createImageMap_output(l_output);
		comboBox_fileDestination_type.setRenderer(new ListRenderer_output());

		comboBox_fileDestination_type.setBackground(Color.WHITE);
		comboBox_fileDestination_type.setFont(new Font("Tahoma", Font.BOLD, 15));
		comboBox_fileDestination_type.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				panel_3_.setVisible(false);

				panel_3_2.setVisible(false);

				chckbxSaveInSame.setSelected(false);

				panel_3_1_2.setVisible(false);

				panel_3_1_1.setVisible(false);

				panel_3_1_1_1.setVisible(false);

				btnStop.setVisible(false);
				comboBox.setEnabled(true);
				lblPortNo.setVisible(false);

				tf_portNo_p3.setVisible(false);
				chckbx_splitpst.setVisible(false);
				lbl_splitpst.setVisible(false);
				chckbx_splitpst.setSelected(false);

				textField_domain_name_p3.setText("");
				panel_5.setVisible(false);
				textField_username_p3.setText("");

				panel_progress.setVisible(true);
				lblMakeSureYou.setVisible(true);
				lblEnableImap_p3.setVisible(true);
				lblTurnOffTwo_p3.setVisible(true);
				chckbxSavePdfAttachment.setVisible(false);
				chckbx_convert_pdf_to_pdf.setVisible(false);
				label_convert_pdf_attch_pdf.setVisible(false);
				label_Save_PDF_Attachments_Separately.setVisible(false);
				tf_Destination_Location.setText(System.getProperty("user.home") + File.separator + "Desktop");
				passwordField_p3.setText("");
				comboBox.setVisible(false);
				btn_signout_p3.setVisible(false);
				btn_signout_p3.setEnabled(true);
				chckbxAutoIncrementBackup.setSelected(false);
				if (!(fileoption.equalsIgnoreCase("GMAIL") || fileoption.equalsIgnoreCase("YAHOO MAIL")
						|| fileoption.equalsIgnoreCase("OFFICE 365") || fileoption.equalsIgnoreCase("AOL")
						|| fileoption.equalsIgnoreCase("Amazon WorkMail")
						|| fileoption.equalsIgnoreCase("Live Exchange") || fileoption.equalsIgnoreCase("HOTMAIL")
						|| fileoption.equalsIgnoreCase("IMAP") || fileoption.equalsIgnoreCase("Zoho Mail")
						|| fileoption.equalsIgnoreCase("Icloud") || fileoption.equalsIgnoreCase("Hostgator email")
						|| fileoption.equalsIgnoreCase("GoDaddy email")
						|| fileoption.equalsIgnoreCase("Yandex Mail"))) {

					chckbxSaveInSame.setVisible(true);
					lblSaveintheSameFolder.setVisible(true);
				}
				btn_converter.setEnabled(false);
				panel_9.setVisible(false);
				lbl_progressreport.setText("");
				panel_9.setVisible(false);

				if (e.getSource() == comboBox_fileDestination_type) {

					JComboBox cb = (JComboBox) e.getSource();

					filetype = (String) cb.getSelectedItem();

				}

				if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("YAHOO MAIL")
						|| filetype.equalsIgnoreCase("OFFICE 365") || filetype.equalsIgnoreCase("AOL")
						|| filetype.equalsIgnoreCase("Amazon WorkMail") || filetype.equalsIgnoreCase("Live Exchange")
						|| filetype.equalsIgnoreCase("HOTMAIL") || filetype.equalsIgnoreCase("IMAP")
						|| filetype.equalsIgnoreCase("Zoho Mail") || filetype.equalsIgnoreCase("Icloud")
						|| filetype.equalsIgnoreCase("Hostgator email") || filetype.equalsIgnoreCase("GoDaddy email")
						|| filetype.equalsIgnoreCase("Yandex Mail")) {
					output = true;
					btn_converter.setEnabled(false);
					lblthirdPartyPassword_1.setVisible(true);
					chckbxSaveInSame.setVisible(false);
					lblSaveintheSameFolder.setVisible(false);
					panel.setVisible(true);
					panel_9.setVisible(false);

					if (filetype.equalsIgnoreCase("GoDaddy email")) {
						panel.setVisible(false);
						lblthirdPartyPassword_1.setVisible(false);

					}

					if (filetype.equalsIgnoreCase("OFFICE 365") || filetype.equalsIgnoreCase("GoDaddy email")) {
						panel.setVisible(false);
						lblthirdPartyPassword_1.setVisible(false);

					}

					lblEnableImap_p3.setVisible(false);
					lblTurnOffTwo_p3.setText("<HTML><U>To access your " + filetype
							+ " account , you'll need to generate and use an app password.</U></HTML>");
					lblMakeSureYou.setText("Please  Click on The Link");

					if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("Zoho Mail")) {
						lblEnableImap_p3.setVisible(true);

						lblTurnOffTwo_p3.setText("<HTML><U>To access your " + filetype
								+ " account , you'll need to generate and use an app password or turn on less secure app</U></HTML>");

					}
					if (filetype.equalsIgnoreCase("Live Exchange")) {
						panel_3_.setVisible(true);

						panel_3_1_1.setVisible(true);

						panel_3_1_1_1.setVisible(true);
						chckbxSaveInSame.setVisible(false);
						lblSaveintheSameFolder.setVisible(false);

						lbl_Domain.setText("IP or Computer Name");
						lblMakeSureYou.setVisible(false);
						lblEnableImap_p3.setVisible(false);
						lblTurnOffTwo_p3.setVisible(false);
						lblthirdPartyPassword_1.setVisible(false);
						lblTurnOffTwo_p3.setText("");
						lblMakeSureYou.setText("");
						lblEnableImap_p3.setText("");

					}
					if (filetype.equalsIgnoreCase("Hostgator email") || filetype.equalsIgnoreCase("IMAP")
							|| filetype.equalsIgnoreCase("Amazon WorkMail")) {
						panel_3_.setVisible(true);

						panel_3_1_1.setVisible(true);

						panel_3_1_1_1.setVisible(true);
						chckbxSaveInSame.setVisible(false);

						lbl_Domain.setText(filetype + " HOST");
						lblthirdPartyPassword_1.setVisible(false);
						lblPortNo.setVisible(true);

						tf_portNo_p3.setVisible(true);
						lblTurnOffTwo_p3.setText("");
						lblMakeSureYou.setText("");
						lblEnableImap_p3.setText("");
						lblMakeSureYou.setVisible(false);
						lblEnableImap_p3.setVisible(false);
						lblTurnOffTwo_p3.setVisible(false);
						lblthirdPartyPassword_1.setVisible(false);

					}

					else {

						if (filetype.equalsIgnoreCase("OFFICE 365") || filetype.equalsIgnoreCase("Live Exchange")
								|| filetype.equalsIgnoreCase("HOTMAIL")) {
							// panel_9.setVisible(true);
							if (fileoption.equalsIgnoreCase("Exchange Offline Storage (.ost)")
									|| fileoption.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
									|| fileoption.equalsIgnoreCase("OLM File (.olm)")) {
								// panel_9.setVisible(true);
							}
							lblTurnOffTwo_p3.setText("<HTML><U>To access your " + filetype
									+ " account , you'll need to generate and use an app password.</U></HTML>");
						}

						panel_3_.setVisible(true);

						panel_3_1_1.setVisible(true);

					}
				} else {
					if (filetype.equalsIgnoreCase("pst")) {

						chckbx_splitpst.setVisible(true);
						lbl_splitpst.setVisible(true);
					}
					if (filetype.equalsIgnoreCase("pdf")) {

						chckbxSavePdfAttachment.setVisible(true);
						label_Save_PDF_Attachments_Separately.setVisible(true);
						chckbx_convert_pdf_to_pdf.setVisible(true);
						label_convert_pdf_attch_pdf.setVisible(true);
					}

					if (filetype.equalsIgnoreCase("Opera Mail")) {

						String str = null;

						if (OS.contains("windows")) {
							str = System.getenv("APPDATA").replace("Roaming", "Local") + File.separator + "Opera Mail"
									+ File.separator + "Opera Mail" + File.separator + "Mail" + File.separator
									+ "store";
						} else {
							str = System.getProperty("user.home") + File.separator + "Library" + File.separator
									+ "Application Support" + File.separator + "Opera Mail" + File.separator + "mail";
						}

						// System.out.println(str);

						if (new File(str).exists()) {

							thunderbirdpath = str;
							tf_Destination_Location.setText(str);

						} else {
							String warn = filetype + " Not Installed Do you want to proced ?";
							int ans = JOptionPane.showConfirmDialog(Main_Frame.this, warn, messageboxtitle,
									JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
									new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
							if (ans == JOptionPane.NO_OPTION) {

								SwingUtilities.invokeLater(new Runnable() {

									public void run() {
										comboBox_fileDestination_type.setSelectedItem("PST");
									}
								});
							}

						}

					} else if (filetype.equalsIgnoreCase("Thunderbird")) {

						String str = null;

						if (OS.contains("windows")) {
							str = System.getenv("APPDATA") + File.separator + "Thunderbird" + File.separator
									+ "Profiles";
						} else {
							str = System.getProperty("user.home") + File.separator + "Library" + File.separator
									+ "Thunderbird" + File.separator + "Profiles";
						}

						// System.out.println(str);

						if (new File(str).exists()) {

							File[] f = new File(str).listFiles();
							for (File fl : f) {
								if (fl != null) {
									if (fl.isDirectory()) {
										String filename = fl.getName();
										String extension = filename.substring(filename.lastIndexOf(".") + 1,
												filename.length());
										String ext = "default";
										if (ext.equals(extension)) {
											// //System.out.println(file);

											String defaultfolder = fl.getName();

											str = str + File.separator + defaultfolder + File.separator + "Mail"
													+ File.separator + "Local Folders";
											thunderbirdpath = str;
											tf_Destination_Location.setText(str);
											break;
										} else {

										}
									}
								}
							}
						} else {
							String warn = filetype + " Not Installed Do you want to proced ?";
							int ans = JOptionPane.showConfirmDialog(Main_Frame.this, warn, messageboxtitle,
									JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
									new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
							if (ans == JOptionPane.NO_OPTION) {

								SwingUtilities.invokeLater(new Runnable() {

									public void run() {
										comboBox_fileDestination_type.setSelectedItem("PST");
									}
								});
							}

						}

					}
					output = false;
					panel_3_.setVisible(true);

					CardLayout card = (CardLayout) panel_3_.getLayout();
					card.show(panel_3_, "panel_3_1_2");

					panel_3_2.setVisible(true);

					btn_converter.setEnabled(true);

					if (!(filetype.equalsIgnoreCase("PST") || filetype.equalsIgnoreCase("Thunderbird")
							|| filetype.equalsIgnoreCase("Opera Mail") || filetype.equalsIgnoreCase("OST")
							|| filetype.equalsIgnoreCase("MBOX") || filetype.equalsIgnoreCase("CSV"))) {
						panel_5.setVisible(true);
						comboBox.setVisible(true);
					}
				}

			}
		});
		comboBox_fileDestination_type.setBounds(236, 13, 638, 35);

		panel_3.add(comboBox_fileDestination_type);

		tf_Destination_Location = new JTextField();
		tf_Destination_Location.setBackground(Color.WHITE);
		tf_Destination_Location.setBounds(198, 8, 638, 30);
		panel_3_2.add(tf_Destination_Location);
		tf_Destination_Location.setEditable(false);
		tf_Destination_Location.setColumns(10);

		btn_Destination = new JButton("");
		btn_Destination.setRolloverEnabled(false);
		btn_Destination.setRequestFocusEnabled(false);
		btn_Destination.setOpaque(false);
		btn_Destination.setFocusable(false);
		btn_Destination.setFocusTraversalKeysEnabled(false);
		btn_Destination.setFocusPainted(false);
		btn_Destination.setDefaultCapable(false);
		btn_Destination.setContentAreaFilled(false);
		btn_Destination.setBorderPainted(false);
		btn_Destination.addMouseListener(new MouseAdapter() {

			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_Destination.setIcon(new ImageIcon(Main_Frame.class.getResource("/path-to-save-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent arg0) {
				btn_Destination.setIcon(new ImageIcon(Main_Frame.class.getResource("/path-to-save-btn.png")));
			}
		});

		btn_Destination.setIcon(new ImageIcon(Main_Frame.class.getResource("/path-to-save-btn.png")));

		btn_Destination.setBounds(886, 8, 130, 30);
		panel_3_2.add(btn_Destination);
		btn_Destination.setFont(new Font("Tahoma", Font.BOLD, 12));

		btn_converter = new JButton("");
		btn_converter.setRolloverEnabled(false);
		btn_converter.setEnabled(false);
		btn_converter.setRequestFocusEnabled(false);
		btn_converter.setOpaque(false);
		btn_converter.setFocusTraversalKeysEnabled(false);
		btn_converter.setFocusable(false);
		btn_converter.setFocusPainted(false);
		btn_converter.setContentAreaFilled(false);
		btn_converter.setBorderPainted(false);
		btn_converter.setDefaultCapable(false);
		btn_converter.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btn_converter.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btn_converter.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-btn.png")));
			}
		});

		btn_converter.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-btn.png")));
		btn_converter.setBounds(931, 562, 134, 39);
		panel_3.add(btn_converter);
		btn_converter.setFont(new Font("Tahoma", Font.BOLD, 12));

		panel_progress = new JPanel();
		panel_progress.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_progress.setBackground(Color.WHITE);
		panel_progress.setBounds(12, 454, 1053, 87);
		panel_3.add(panel_progress);
		panel_progress.setLayout(null);

		btnStop = new JButton("");
		btnStop.setBounds(899, 8, 118, 39);
		panel_progress.add(btnStop);
		btnStop.setContentAreaFilled(false);
		btnStop.setBorderPainted(false);
		btnStop.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnStop.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnStop.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-btn.png")));
			}
		});

		btnStop.setIcon(new ImageIcon(Main_Frame.class.getResource("/stop-btn.png")));
		btnStop.setRolloverEnabled(false);
		btnStop.setRequestFocusEnabled(false);
		btnStop.setOpaque(false);
		btnStop.setFocusable(false);
		btnStop.setFocusTraversalKeysEnabled(false);
		btnStop.setFocusPainted(false);
		btnStop.setDefaultCapable(false);

		Progressbar = new JLabel("");
		Progressbar.setBounds(10, 15, 891, 24);
		panel_progress.add(Progressbar);
		Progressbar.setVisible(false);
		Progressbar.setIcon(new ImageIcon(Main_Frame.class.getResource("/progress-bar.gif")));
		JLabel lblNamingConvention = new JLabel("Naming Convention :");
		lblNamingConvention.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNamingConvention.setBounds(10, 14, 134, 23);
		panel_5.add(lblNamingConvention);
		SwingUtilities.invokeLater(new Runnable() {

			public void run() {

				comboBox.setSelectedItem("Subject");

				comboBox_fileDestination_type.setSelectedItem("PST");
			}
		});

		lbl_progressreport = new JLabel("");
		lbl_progressreport.setBounds(20, 50, 882, 24);
		panel_progress.add(lbl_progressreport);

		JLabel label_18 = new JLabel("");
		label_18.setToolTipText(
				"Save the mailbox data by providing Date, Subject, To, From in different combinations as a naming base. "
						+ System.lineSeparator()
						+ " It makes the manoeuvring of the resultant mailbox data easy and convenient.");
		label_18.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_18.setBounds(462, 15, 26, 20);
		panel_5.add(label_18);

		JPanel panel_duplicacy = new JPanel();
		panel_duplicacy.setBounds(7, 5, 500, 303);
		panel_duplicacy.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_duplicacy.setBackground(Color.WHITE);
		panel_3_1_2.add(panel_duplicacy);
		panel_duplicacy.setLayout(null);

		label_convert_pdf_attch_pdf = new JLabel("");
		label_convert_pdf_attch_pdf.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_convert_pdf_attch_pdf
				.setToolTipText("It helps to convert the contained attachments of PDF files to PDF format.");
		label_convert_pdf_attch_pdf.setBounds(464, 83, 26, 23);
		panel_duplicacy.add(label_convert_pdf_attch_pdf);

		chckbxRemoveDuplicacy = new JCheckBox("Remove Duplicate Mail On basis of To, From, Subject, Bcc, Body");
//		chckbxRemoveDuplicacy.setToolTipText("All the replicated or duplicated emails will get deleted "
//				+ System.lineSeparator() + "on the basis of To, From, Subject, Bcc, and Body.");
		chckbxRemoveDuplicacy.setRolloverEnabled(false);
		chckbxRemoveDuplicacy.setRequestFocusEnabled(false);
		chckbxRemoveDuplicacy.setOpaque(false);
		chckbxRemoveDuplicacy.setFocusable(false);
		chckbxRemoveDuplicacy.setFocusPainted(false);
		chckbxRemoveDuplicacy.setContentAreaFilled(false);
		chckbxRemoveDuplicacy.setForeground(Color.RED);
		chckbxRemoveDuplicacy.setFont(new Font("Tahoma", Font.PLAIN, 10));
		chckbxRemoveDuplicacy.setBackground(Color.WHITE);
		chckbxRemoveDuplicacy.setBounds(6, 27, 396, 16);
		panel_duplicacy.add(chckbxRemoveDuplicacy);

		chckbxMaintainFolderHeirachy = new JCheckBox("Maintain Folder Hierarchy");
		// chckbxMaintainFolderHeirachy.setToolTipText("Maintain the folder hierarchy of
		// your mailbox.");
		chckbxMaintainFolderHeirachy.setRolloverEnabled(false);
		chckbxMaintainFolderHeirachy.setRequestFocusEnabled(false);
		chckbxMaintainFolderHeirachy.setOpaque(false);
		chckbxMaintainFolderHeirachy.setFocusable(false);
		chckbxMaintainFolderHeirachy.setFocusPainted(false);
		chckbxMaintainFolderHeirachy.setContentAreaFilled(false);
		chckbxMaintainFolderHeirachy.setFont(new Font("Tahoma", Font.PLAIN, 10));
		chckbxMaintainFolderHeirachy.setBounds(6, 46, 186, 23);
		panel_duplicacy.add(chckbxMaintainFolderHeirachy);
		chckbxMaintainFolderHeirachy.setForeground(Color.RED);
		chckbxMaintainFolderHeirachy.setBackground(Color.WHITE);

		chckbxDeleteEmailFrom = new JCheckBox("Free up Server Space");
		// chckbxDeleteEmailFrom.setToolTipText("To free up space, all the emails will
		// get deleted from. ");
		chckbxDeleteEmailFrom.setRolloverEnabled(false);
		chckbxDeleteEmailFrom.setRequestFocusEnabled(false);
		chckbxDeleteEmailFrom.setOpaque(false);
		chckbxDeleteEmailFrom.setFocusable(false);
		chckbxDeleteEmailFrom.setFocusPainted(false);
		chckbxDeleteEmailFrom.setContentAreaFilled(false);
		chckbxDeleteEmailFrom.setFont(new Font("Tahoma", Font.PLAIN, 10));
		chckbxDeleteEmailFrom.setBounds(6, 110, 448, 14);
		panel_duplicacy.add(chckbxDeleteEmailFrom);
		chckbxDeleteEmailFrom.setForeground(Color.RED);
		chckbxDeleteEmailFrom.setBackground(Color.WHITE);

		chckbxAutoIncrementBackup = new JCheckBox("Skip Previously Migrated Items");
		// chckbxAutoIncrementBackup.setToolTipText("Suppose, previously you\u2019ve
		// saved/backup the mailbox. ");
		chckbxAutoIncrementBackup.setRolloverEnabled(false);
		chckbxAutoIncrementBackup.setRequestFocusEnabled(false);
		chckbxAutoIncrementBackup.setOpaque(false);
		chckbxAutoIncrementBackup.setFocusable(false);
		chckbxAutoIncrementBackup.setFocusPainted(false);
		chckbxAutoIncrementBackup.setContentAreaFilled(false);
		chckbxAutoIncrementBackup.setFont(new Font("Tahoma", Font.PLAIN, 10));
		chckbxAutoIncrementBackup.setForeground(Color.RED);
		chckbxAutoIncrementBackup.setBackground(Color.WHITE);
		chckbxAutoIncrementBackup.setBounds(6, 128, 175, 23);
		panel_duplicacy.add(chckbxAutoIncrementBackup);

		chckbxSaveInSame = new JCheckBox("Save in the Same Folder (Source and Destination Folder are same)");
//		chckbxSaveInSame.setToolTipText("All the resultant data will get saved at the " + System.lineSeparator()
//				+ "location of the source file.");

		chckbxSaveInSame.setRolloverEnabled(false);
		chckbxSaveInSame.setRequestFocusEnabled(false);
		chckbxSaveInSame.setOpaque(false);
		chckbxSaveInSame.setFocusable(false);
		chckbxSaveInSame.setFocusPainted(false);
		chckbxSaveInSame.setContentAreaFilled(false);
		chckbxSaveInSame.setBackground(Color.WHITE);
		chckbxSaveInSame.setVisible(false);
		chckbxSaveInSame.setForeground(Color.RED);
		chckbxSaveInSame.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {

				if (arg0.getStateChange() == ItemEvent.SELECTED) {

					destination = tf_Destination_Location.getText();
					btn_Destination.setEnabled(false);
					tf_Destination_Location.setText(filepath(new File(filepath)));
				}

				else {
					btn_Destination.setEnabled(true);
					tf_Destination_Location.setText(destination);

				}
			}
		});
		chckbxSaveInSame.setFont(new Font("Tahoma", Font.PLAIN, 10));
		chckbxSaveInSame.setBounds(6, 8, 452, 16);

		panel_duplicacy.add(chckbxSaveInSame);

		chckbxSavePdfAttachment = new JCheckBox("Save attachment separately");
//		chckbxSavePdfAttachment.setToolTipText(
//				"Save all the email attachments separately in a  " + System.lineSeparator() + "folder.");
		chckbxSavePdfAttachment.setRolloverEnabled(false);
		chckbxSavePdfAttachment.setRequestFocusEnabled(false);
		chckbxSavePdfAttachment.setOpaque(false);
		chckbxSavePdfAttachment.setFocusable(false);
		chckbxSavePdfAttachment.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (chckbxMigrateOrBackup.isSelected() && chckbxSavePdfAttachment.isSelected()) {

					JOptionPane.showMessageDialog(frame,
							"Please Unselect Migrate or Backup Emails Without Attachment files checkbox before continuing.",
							messageboxtitle, JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));

					chckbxSavePdfAttachment.setSelected(false);
				}
			}
		});
		chckbxSavePdfAttachment.setFocusPainted(false);
		chckbxSavePdfAttachment.setContentAreaFilled(false);
		chckbxSavePdfAttachment.setForeground(Color.RED);
		chckbxSavePdfAttachment.setVisible(false);
		chckbxSavePdfAttachment.setFont(new Font("Tahoma", Font.PLAIN, 10));
		chckbxSavePdfAttachment.setBackground(Color.WHITE);
		chckbxSavePdfAttachment.setBounds(6, 64, 454, 23);
		panel_duplicacy.add(chckbxSavePdfAttachment);

		lblSaveintheSameFolder = new JLabel("");

		lblSaveintheSameFolder.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		lblSaveintheSameFolder.setToolTipText("All the resultant data will get saved at the " + System.lineSeparator()
				+ "location of the source file.");
		lblSaveintheSameFolder.setVisible(false);
		lblSaveintheSameFolder.setBounds(464, 8, 26, 23);
		panel_duplicacy.add(lblSaveintheSameFolder);

		label_Remove_Duplicate = new JLabel("");
		label_Remove_Duplicate.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_Remove_Duplicate.setToolTipText("All the replicated or duplicated emails will get deleted "
				+ System.lineSeparator() + "on the basis of To, From, Subject, Bcc, and Body.");

		label_Remove_Duplicate.setBounds(464, 27, 26, 23);
		panel_duplicacy.add(label_Remove_Duplicate);

		label_Maintain_Folder_Hierarchy = new JLabel("");
		label_Maintain_Folder_Hierarchy.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_Maintain_Folder_Hierarchy.setToolTipText("Maintain the folder hierarchy of your mailbox.");

		label_Maintain_Folder_Hierarchy.setBounds(464, 46, 27, 23);
		panel_duplicacy.add(label_Maintain_Folder_Hierarchy);

		label_Save_PDF_Attachments_Separately = new JLabel("");
		label_Save_PDF_Attachments_Separately.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_Save_PDF_Attachments_Separately.setToolTipText(
				"Save all the email attachments separately in a  " + System.lineSeparator() + "folder.");
		label_Save_PDF_Attachments_Separately.setVisible(false);
		label_Save_PDF_Attachments_Separately.setBounds(464, 64, 26, 23);
		panel_duplicacy.add(label_Save_PDF_Attachments_Separately);

		label_Free_up_Server_Space = new JLabel("");
		label_Free_up_Server_Space.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_Free_up_Server_Space.setToolTipText(
				"<html>To free up space, all the emails will get deleted from <br/> the server as soon as the process ends.</html>");

		label_Free_up_Server_Space.setBounds(464, 102, 26, 23);
		panel_duplicacy.add(label_Free_up_Server_Space);

		lblSkip_Previously_Migrated_Items = new JLabel("");
		lblSkip_Previously_Migrated_Items.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		lblSkip_Previously_Migrated_Items.setToolTipText(
				"<html>Suppose, previously you’ve saved/backup the mailbox <br/> items to a particular email client or premise system using our utility.<br/>  In that case, it will ensure to skip those email items and save the fresh ones only.</html>");

		lblSkip_Previously_Migrated_Items.setBounds(463, 123, 26, 23);
		panel_duplicacy.add(lblSkip_Previously_Migrated_Items);

		chckbxSetBackupSchedule = new JCheckBox("Set Backup Schedule ");
		chckbxSetBackupSchedule.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					runNextTimeComponents(true);
				} else {
					rdbtnOnce.setSelected(false);
					runNextTimeComponents(false);
				}

			}
		});
		chckbxSetBackupSchedule.setFont(new Font("Tahoma", Font.PLAIN, 10));
		chckbxSetBackupSchedule.setForeground(Color.RED);
		chckbxSetBackupSchedule.setBackground(Color.WHITE);
		chckbxSetBackupSchedule.setBounds(6, 205, 251, 20);
		if (versiontype > 2 || demo)
			panel_duplicacy.add(chckbxSetBackupSchedule);

		spinner = new JSpinner();
		spinner.setVisible(false);
		spinner.setModel(new SpinnerDateModel());

		spinner.setEditor(new JSpinner.DateEditor(spinner, "HH:mm:ss"));

		spinner.setBounds(247, 230, 82, 23);
		panel_duplicacy.add(spinner);

		dateChooserNextSchedular = new JDateChooser();
		dateChooserNextSchedular.setVisible(false);
		Date startdate = java.util.Calendar.getInstance().getTime();
		dateChooserNextSchedular.setMinSelectableDate(startdate);
		dateChooserNextSchedular.setDate(startdate);

		dateChooserNextSchedular.setBounds(104, 232, 116, 20);
		panel_duplicacy.add(dateChooserNextSchedular);

		label_14 = new Label("Task run time:");
		label_14.setVisible(false);
		label_14.setBounds(5, 230, 94, 22);
		panel_duplicacy.add(label_14);

		rdbtnOnce = new JRadioButton("Once");
		rdbtnOnce.setVisible(false);
		rdbtnOnce.setBackground(Color.WHITE);
		rdbtnOnce.setFont(new Font("Tahoma", Font.PLAIN, 10));
		rdbtnOnce.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					dateChooserNextSchedular.setEnabled(true);
					Date startdate = java.util.Calendar.getInstance().getTime();
					dateChooserNextSchedular.setDate(startdate);

				}

			}
		});
		rdbtnOnce.setSelected(true);
		buttonGroup_Schedulling.add(rdbtnOnce);
		rdbtnOnce.setBounds(6, 260, 62, 16);
		panel_duplicacy.add(rdbtnOnce);

		rdbtnEveryday = new JRadioButton("Every Day :\r\n");
		rdbtnEveryday.setVisible(false);
		rdbtnEveryday.setBackground(Color.WHITE);
		rdbtnEveryday.setFont(new Font("Tahoma", Font.PLAIN, 10));
		rdbtnEveryday.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {

					dateChooserNextSchedular.setEnabled(true);
					Date startdate = java.util.Calendar.getInstance().getTime();
					dateChooserNextSchedular.setDate(startdate);
					Calendar cal = dateChooserNextSchedular.getCalendar();
					cal.add(Calendar.DATE, 1);

					dateChooserNextSchedular.setDate(cal.getTime());
					dateChooserNextSchedular.setEnabled(false);

				}

			}
		});
		buttonGroup_Schedulling.add(rdbtnEveryday);
		rdbtnEveryday.setBounds(107, 259, 99, 16);
		panel_duplicacy.add(rdbtnEveryday);

		rdbtnEveryWeek = new JRadioButton("Every Week");
		rdbtnEveryWeek.setVisible(false);
		rdbtnEveryWeek.setFont(new Font("Tahoma", Font.PLAIN, 10));
		rdbtnEveryWeek.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {

					Date startdate = java.util.Calendar.getInstance().getTime();
					dateChooserNextSchedular.setDate(startdate);
					Calendar cal = dateChooserNextSchedular.getCalendar();
					cal.add(Calendar.DATE, 7);

					dateChooserNextSchedular.setDate(cal.getTime());

				}

			}
		});
		rdbtnEveryWeek.setBackground(Color.WHITE);
		buttonGroup_Schedulling.add(rdbtnEveryWeek);
		rdbtnEveryWeek.setBounds(6, 279, 93, 16);
		panel_duplicacy.add(rdbtnEveryWeek);

		comboBox_weekdays = new JComboBox();
		comboBox_weekdays.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				dateChooserNextSchedular.setEnabled(true);

				Date startdate = java.util.Calendar.getInstance().getTime();
				dateChooserNextSchedular.setDate(startdate);

				Calendar cal = dateChooserNextSchedular.getCalendar();
				int in = (comboBox_weekdays.getSelectedIndex() + 1) - cal.get(Calendar.DAY_OF_WEEK);

				if (in <= 0) {
					in += 7;
				}

				cal.add(Calendar.DAY_OF_MONTH, in);

				System.out.println(cal.getTime());

				dateChooserNextSchedular.setDate(cal.getTime());

				if (!rdbtnOnce.isSelected()) {
					dateChooserNextSchedular.setEnabled(false);
				}
			}
		});

		comboBox_weekdays.setVisible(false);
		comboBox_weekdays.addItem("Sunday");
		comboBox_weekdays.addItem("Monday");
		comboBox_weekdays.addItem("Tuesday");
		comboBox_weekdays.addItem("Wednesday");
		comboBox_weekdays.addItem("Thursday");
		comboBox_weekdays.addItem("Friday");
		comboBox_weekdays.addItem("Saturday");
		comboBox_weekdays.setBounds(343, 258, 111, 20);
		panel_duplicacy.add(comboBox_weekdays);

		rdbtnOnWeekDay = new JRadioButton(" On Week Day : ");
		rdbtnOnWeekDay.setVisible(false);
		buttonGroup_Schedulling.add(rdbtnOnWeekDay);
		rdbtnOnWeekDay.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					comboBox_weekdays.setVisible(true);
					Date startdate = java.util.Calendar.getInstance().getTime();
					dateChooserNextSchedular.setDate(startdate);

					Calendar cal = dateChooserNextSchedular.getCalendar();
					int in = (comboBox_weekdays.getSelectedIndex() + 1) - cal.get(Calendar.DAY_OF_WEEK);

					if (in <= 0) {
						in += 7;
					}

					cal.add(Calendar.DAY_OF_MONTH, in);

					dateChooserNextSchedular.setDate(cal.getTime());
					dateChooserNextSchedular.setEnabled(false);

				} else {
					comboBox_weekdays.setVisible(false);
				}

			}
		});
		rdbtnOnWeekDay.setFont(new Font("Tahoma", Font.PLAIN, 10));
		rdbtnOnWeekDay.setBackground(Color.WHITE);
		rdbtnOnWeekDay.setBounds(232, 260, 107, 16);
		panel_duplicacy.add(rdbtnOnWeekDay);

		rdbtnOnmonthDay = new JRadioButton("On Month Day :");
		rdbtnOnmonthDay.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					dateChooserNextSchedular.setEnabled(true);
					comboBox_MonthDay.setVisible(true);
					Date startdate = java.util.Calendar.getInstance().getTime();
					dateChooserNextSchedular.setDate(startdate);

					Calendar cal = dateChooserNextSchedular.getCalendar();
					int calenderdate = comboBox_MonthDay.getSelectedIndex() + 1;
					cal.set(Calendar.DAY_OF_MONTH, calenderdate);
					if (cal.get(Calendar.DAY_OF_MONTH) >= calenderdate) {
						cal.add(Calendar.MONTH, 1);
					}

					dateChooserNextSchedular.setDate(cal.getTime());
					dateChooserNextSchedular.setEnabled(false);
				} else {
					comboBox_MonthDay.setVisible(false);
				}

			}
		});
		buttonGroup_Schedulling.add(rdbtnOnmonthDay);
		rdbtnOnmonthDay.setFont(new Font("Tahoma", Font.PLAIN, 10));
		rdbtnOnmonthDay.setBackground(Color.WHITE);
		rdbtnOnmonthDay.setBounds(232, 279, 107, 16);
		rdbtnOnmonthDay.setVisible(false);
		panel_duplicacy.add(rdbtnOnmonthDay);

		comboBox_MonthDay = new JComboBox();
		comboBox_MonthDay.setVisible(false);
		comboBox_MonthDay.addItem("1");
		comboBox_MonthDay.addItem("2");
		comboBox_MonthDay.addItem("3");
		comboBox_MonthDay.addItem("4");
		comboBox_MonthDay.addItem("5");
		comboBox_MonthDay.addItem("6");
		comboBox_MonthDay.addItem("7");
		comboBox_MonthDay.addItem("8");
		comboBox_MonthDay.addItem("9");
		comboBox_MonthDay.addItem("10");
		comboBox_MonthDay.addItem("11");
		comboBox_MonthDay.addItem("12");
		comboBox_MonthDay.addItem("13");
		comboBox_MonthDay.addItem("14");
		comboBox_MonthDay.addItem("15");
		comboBox_MonthDay.addItem("16");
		comboBox_MonthDay.addItem("17");
		comboBox_MonthDay.addItem("18");
		comboBox_MonthDay.addItem("19");
		comboBox_MonthDay.addItem("20");
		comboBox_MonthDay.addItem("21");
		comboBox_MonthDay.addItem("22");
		comboBox_MonthDay.addItem("23");
		comboBox_MonthDay.addItem("24");
		comboBox_MonthDay.addItem("25");
		comboBox_MonthDay.addItem("26");
		comboBox_MonthDay.addItem("27");
		comboBox_MonthDay.addItem("28");
		comboBox_MonthDay.addItem("29");
		comboBox_MonthDay.addItem("30");
		comboBox_MonthDay.addItem("31");
		comboBox_MonthDay.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				dateChooserNextSchedular.setEnabled(true);

				Date startdate = java.util.Calendar.getInstance().getTime();
				dateChooserNextSchedular.setDate(startdate);
				comboBox_MonthDay.setVisible(true);
				Calendar cal = dateChooserNextSchedular.getCalendar();
				int calenderdate = comboBox_MonthDay.getSelectedIndex() + 1;

				if (cal.get(Calendar.DAY_OF_MONTH) >= calenderdate) {
					cal.add(Calendar.MONTH, 1);
				}
				cal.set(Calendar.DAY_OF_MONTH, calenderdate);
				dateChooserNextSchedular.setDate(cal.getTime());

				if (!rdbtnOnce.isSelected()) {
					dateChooserNextSchedular.setEnabled(false);
				}
			}
		});
		comboBox_MonthDay.setBounds(345, 277, 57, 20);
		panel_duplicacy.add(comboBox_MonthDay);

		rdbtnEveryMonth = new JRadioButton("Every Month");
		buttonGroup_Schedulling.add(rdbtnEveryMonth);
		rdbtnEveryMonth.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					dateChooserNextSchedular.setEnabled(true);
					Date startdate = java.util.Calendar.getInstance().getTime();
					dateChooserNextSchedular.setDate(startdate);
					Calendar cal = dateChooserNextSchedular.getCalendar();
					cal.add(Calendar.MONTH, 1);

					dateChooserNextSchedular.setDate(cal.getTime());
					dateChooserNextSchedular.setEnabled(false);
				}

			}
		});
		rdbtnEveryMonth.setBackground(Color.WHITE);
		rdbtnEveryMonth.setFont(new Font("Tahoma", Font.PLAIN, 10));
		rdbtnEveryMonth.setBounds(104, 279, 97, 16);
		rdbtnEveryMonth.setVisible(false);
		panel_duplicacy.add(rdbtnEveryMonth);

		chckbxMigrateOrBackup = new JCheckBox("Migrate or Backup Emails Without Attachment files");
		chckbxMigrateOrBackup.setToolTipText(
				"Check the option, If you want to migrate or backup emails without their attachment files.");
		chckbxMigrateOrBackup.setRolloverEnabled(false);
		chckbxMigrateOrBackup.setRequestFocusEnabled(false);
		chckbxMigrateOrBackup.setOpaque(false);
		chckbxMigrateOrBackup.setFocusable(false);
		chckbxMigrateOrBackup.setFocusPainted(false);
		chckbxMigrateOrBackup.setContentAreaFilled(false);
		chckbxMigrateOrBackup.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (chckbxMigrateOrBackup.isSelected()) {

					chckbxSavePdfAttachment.setSelected(false);
					chckbx_convert_pdf_to_pdf.setSelected(false);
				} else {
					if (filetype.equalsIgnoreCase("pdf"))
						chckbxSavePdfAttachment.setVisible(true);
				}
			}
		});
		chckbxMigrateOrBackup.setFont(new Font("Tahoma", Font.PLAIN, 10));
		chckbxMigrateOrBackup.setForeground(Color.RED);
		chckbxMigrateOrBackup.setBackground(Color.WHITE);
		chckbxMigrateOrBackup.setBounds(6, 153, 448, 23);
		panel_duplicacy.add(chckbxMigrateOrBackup);

		JLabel lbl_migratewithoutatt = new JLabel("");
		lbl_migratewithoutatt.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		lbl_migratewithoutatt.setToolTipText(
				"Check the option, If you want to migrate or backup emails without their attachment files.");

		lbl_migratewithoutatt.setBounds(464, 157, 26, 23);
		panel_duplicacy.add(lbl_migratewithoutatt);

		label_backupschedul = new JLabel("");
		label_backupschedul.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		label_backupschedul.setToolTipText(
				"Check the option, If you want to execute the backup process at regular intervals like Day/Week/Month.");

		label_backupschedul.setBounds(464, 202, 26, 23);
		if (versiontype > 2 || demo)
			panel_duplicacy.add(label_backupschedul);

		chckbx_convert_pdf_to_pdf = new JCheckBox("Convert Attachments to PDF Format");
		chckbx_convert_pdf_to_pdf
				.setToolTipText("It helps to convert the contained attachments of Attachment files to PDF format.");
		chckbx_convert_pdf_to_pdf.setRolloverEnabled(false);
		chckbx_convert_pdf_to_pdf.setRequestFocusEnabled(false);
		chckbx_convert_pdf_to_pdf.setOpaque(false);
		chckbx_convert_pdf_to_pdf.setFocusable(false);
		chckbx_convert_pdf_to_pdf.setFocusPainted(false);
		chckbx_convert_pdf_to_pdf.setContentAreaFilled(false);
		chckbx_convert_pdf_to_pdf.setForeground(Color.RED);
		chckbx_convert_pdf_to_pdf.setFont(new Font("Tahoma", Font.PLAIN, 10));
		chckbx_convert_pdf_to_pdf.setBackground(Color.WHITE);
		chckbx_convert_pdf_to_pdf.setBounds(7, 86, 451, 23);
		chckbx_convert_pdf_to_pdf.setVisible(false);
		chckbx_convert_pdf_to_pdf.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (chckbxMigrateOrBackup.isSelected() && chckbx_convert_pdf_to_pdf.isSelected()) {

					JOptionPane.showMessageDialog(frame,
							"Please Unselect Migrate or Backup Emails Without Attachment files checkbox before continuing.",
							messageboxtitle, JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));

					chckbx_convert_pdf_to_pdf.setSelected(false);
				} else if (chckbx_convert_pdf_to_pdf.isSelected()) {
					chckbxSavePdfAttachment.setSelected(false);
				}
			}
		});

		panel_duplicacy.add(chckbx_convert_pdf_to_pdf);
		label_convert_pdf_attch_pdf.setVisible(false);
		chckbx_splitpst = new JCheckBox("Split Resultant PST\r\n");
		chckbx_splitpst.setToolTipText("Split resultant PST file according to size.");
		chckbx_splitpst.setVisible(false);
		chckbx_splitpst.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent e) {
				if (chckbx_splitpst.isSelected()) {

					spinner_sizespinner.setVisible(true);
					comboBox_setsize.setVisible(true);
					spinner_sizespinner.setEnabled(true);
					comboBox_setsize.setEnabled(true);

				} else {

					spinner_sizespinner.setVisible(false);
					comboBox_setsize.setVisible(false);
					spinner_sizespinner.setEnabled(false);
					comboBox_setsize.setEnabled(false);

				}

			}
		});
		chckbx_splitpst.setForeground(Color.RED);
		chckbx_splitpst.setFont(new Font("Tahoma", Font.PLAIN, 10));
		chckbx_splitpst.setBackground(Color.WHITE);
		chckbx_splitpst.setBounds(6, 179, 147, 23);
		panel_duplicacy.add(chckbx_splitpst);

		spinner_sizespinner = new JSpinner();
		spinner_sizespinner.setVisible(false);
		spinner_sizespinner.setFont(new Font("Calibri", Font.PLAIN, 14));
		spinner_sizespinner.setBackground(Color.WHITE);
		spinner_sizespinner.setFont(new Font("Calibri", Font.PLAIN, 14));
		spinner_sizespinner.setEditor(new JSpinner.DefaultEditor(spinner));
		spinner_sizespinner.setBackground(Color.WHITE);
		spinner_sizespinner.setBounds(180, 183, 52, 20);
		SpinnerModel sm = new SpinnerNumberModel(5, 1, 900, 1);

		spinner_sizespinner.setModel(sm);
		spinner_sizespinner.setValue(1);

		spinner_sizespinner.setEditor(new JSpinner.DefaultEditor(spinner_sizespinner));
		panel_duplicacy.add(spinner_sizespinner);

		comboBox_setsize = new JComboBox();
		comboBox_setsize.setVisible(false);
		comboBox_setsize.setFont(new Font("Calibri", Font.PLAIN, 14));
		comboBox_setsize.setBackground(Color.WHITE);
		comboBox_setsize.setFont(new Font("Calibri", Font.PLAIN, 14));
		comboBox_setsize.setBackground(Color.WHITE);
		comboBox_setsize.addItem("MB");
		comboBox_setsize.addItem("GB");
		comboBox_setsize.setSelectedItem(0);
		comboBox_setsize.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (comboBox_setsize.getSelectedIndex() == 0) {
					spinner_sizespinner.setEditor(new JSpinner.NumberEditor(spinner_sizespinner));
					SpinnerModel sm = new SpinnerNumberModel(1, 1, 900, 1);
					spinner_sizespinner.setModel(sm);

					spinner_sizespinner.setValue(sm.getValue());
					spinner_sizespinner.setEditor(new JSpinner.DefaultEditor(spinner_sizespinner));
				} else if (comboBox_setsize.getSelectedIndex() == 1) {
					spinner_sizespinner.setEditor(new JSpinner.NumberEditor(spinner_sizespinner));
					SpinnerModel sm = new SpinnerNumberModel(1, 1, 20, 1);
					spinner_sizespinner.setModel(sm);
					spinner_sizespinner.setValue(sm.getValue());
					spinner_sizespinner.setEditor(new JSpinner.DefaultEditor(spinner_sizespinner));
				}
			}
		});
		// bysizesplitcomboBox.setBounds(249, 16, 74, 20);
		comboBox_setsize.setBounds(242, 182, 87, 20);
		panel_duplicacy.add(comboBox_setsize);

		lbl_splitpst = new JLabel("");
		lbl_splitpst.setVisible(false);
		lbl_splitpst.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		lbl_splitpst.setToolTipText("Split resultant PST file according to size.");
		lbl_splitpst.setBounds(464, 179, 26, 23);
		panel_duplicacy.add(lbl_splitpst);

		JPanel panel_6 = new JPanel();
		panel_6.setBounds(518, 152, 500, 44);
		panel_6.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_6.setBackground(Color.WHITE);
		panel_3_1_2.add(panel_6);
		panel_6.setLayout(null);

		textField_customfolder = new JTextField();
		textField_customfolder.addKeyListener(new KeyAdapter() {
			public void keyReleased(KeyEvent event) {

				String content = textField_customfolder.getText();
				if (content.contains(":") || content.contains(":") || content.contains("\\") || content.contains("?")
						|| content.contains("/") || content.contains("|") || content.contains("*")
						|| content.contains("<") || content.contains(">") || content.contains("\t")
						|| content.contains("//s") || content.contains("\"")) {
					textField_customfolder.setText(getRidOfIllegalFileNameCharacters(content).trim());
				}
			}
		});
		textField_customfolder.setBounds(181, 10, 237, 23);
		panel_6.add(textField_customfolder);
		textField_customfolder.setEditable(false);
		textField_customfolder.setColumns(10);

		chckbxCustomFolderName = new JCheckBox("Custom Folder Name :");
		chckbxCustomFolderName.setRolloverEnabled(false);
		chckbxCustomFolderName.setRequestFocusEnabled(false);
		chckbxCustomFolderName.setOpaque(false);
		chckbxCustomFolderName.setFocusable(false);
		chckbxCustomFolderName.setFocusPainted(false);
		chckbxCustomFolderName.setContentAreaFilled(false);
		chckbxCustomFolderName.setBounds(6, 10, 169, 23);
		panel_6.add(chckbxCustomFolderName);
		chckbxCustomFolderName.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbxCustomFolderName.addItemListener(new ItemListener() {
			public void itemStateChanged(ItemEvent arg0) {
				if (arg0.getStateChange() == ItemEvent.SELECTED) {
					textField_customfolder.setEditable(true);
				}

				else {
					textField_customfolder.setEditable(false);
				}
			}
		});

		chckbxCustomFolderName.setBackground(Color.WHITE);

		JLabel lblNewLabel_10 = new JLabel("");
		lblNewLabel_10.setToolTipText(
				"Provide a name of your choice to the destination folder or where your mailbox data will get saved. "
						+ System.lineSeparator() + " it must not contain these characters :\\?/|*<>\t.");
		lblNewLabel_10.setIcon(new ImageIcon(Main_Frame.class.getResource("/infolabel.png")));
		lblNewLabel_10.setBounds(464, 10, 26, 23);
		panel_6.add(lblNewLabel_10);

		panel_9 = new JPanel();
		panel_9.setBounds(517, 204, 501, 48);
		panel_9.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		panel_9.setBackground(Color.WHITE);
		panel_3_1_2.add(panel_9);
		panel_9.setLayout(null);

		chckbxRestoreToDefault = new JCheckBox("Restore to Default Folder");
		chckbxRestoreToDefault.setRolloverEnabled(false);
		chckbxRestoreToDefault.setRequestFocusEnabled(false);
		chckbxRestoreToDefault.setOpaque(false);
		chckbxRestoreToDefault.setFocusable(false);
		chckbxRestoreToDefault.setFocusPainted(false);
		chckbxRestoreToDefault.setContentAreaFilled(false);
		chckbxRestoreToDefault.setBackground(Color.WHITE);
		chckbxRestoreToDefault.setVisible(false);
		chckbxRestoreToDefault.setFont(new Font("Tahoma", Font.BOLD, 10));
		chckbxRestoreToDefault.setBounds(6, 7, 153, 34);
		panel_9.add(chckbxRestoreToDefault);

		JLabel lblNewLabel_8 = new JLabel("(PST folder to office 365  inbox , calendar , contact , sent item etc)");
		lblNewLabel_8.setBounds(169, 7, 322, 30);
		panel_9.add(lblNewLabel_8);

		scrollPane_2 = new JScrollPane();
		scrollPane_2.setBackground(Color.WHITE);
		scrollPane_2.setBounds(520, 63, 393, 70);
		panel_3_1_2.add(scrollPane_2);

		table = new JTable();
		scrollPane_2.setViewportView(table);
		table.setModel(new DefaultTableModel(new Object[][] {}, new String[] { "From", "To" }));

		JPanel panel_11 = new JPanel();
		panel_11.setBackground(Color.WHITE);
		panel_11.setBorder(new TitledBorder(
				new EtchedBorder(EtchedBorder.LOWERED, new Color(255, 255, 255), new Color(160, 160, 160)),
				"                   ", TitledBorder.LEADING, TitledBorder.TOP, null, new Color(0, 0, 0)));
		panel_11.setBounds(507, 5, 546, 142);
		panel_3_1_2.add(panel_11);
		panel_11.setLayout(null);

		JLabel lblNewLabel_9 = new JLabel("Start Date :");
		lblNewLabel_9.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNewLabel_9.setBounds(10, 23, 80, 25);
		panel_11.add(lblNewLabel_9);
		lblNewLabel_9.setBackground(Color.WHITE);

		dateChooser_newFrom = new JDateChooser();
		dateChooser_newFrom.getCalendarButton().setBackground(Color.WHITE);
		dateChooser_newFrom.addMouseListener(new MouseAdapter() {

			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_newFrom.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_newFrom.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_newFrom.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_newFrom.setBackground(Color.WHITE);
		dateChooser_newFrom.setBounds(100, 23, 102, 25);
		panel_11.add(dateChooser_newFrom);
		dateChooser_newFrom.setEnabled(false);
		dateChooser_newFrom.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				Calendar cal2 = Calendar.getInstance();

				cal2.set(Calendar.HOUR_OF_DAY, 00);
				cal2.set(Calendar.MINUTE, 00);
				cal2.set(Calendar.SECOND, 00);
				Date startdate = cal2.getTime();
				dateChooser_newFrom.setMaxSelectableDate(startdate);

			}
		});
		dateChooser_newFrom.setDateFormatString("dd-MMM-yyyy");

		JLabel lblNewLabel_11 = new JLabel("End Date :");
		lblNewLabel_11.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNewLabel_11.setBounds(222, 23, 70, 25);
		panel_11.add(lblNewLabel_11);

		dateChooser_newTo = new JDateChooser();
		dateChooser_newTo.getCalendarButton().setBackground(Color.WHITE);
		dateChooser_newTo.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				dateChooser_newTo.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				dateChooser_newTo.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
			}
		});
		dateChooser_newTo.setIcon(new ImageIcon(Main_Frame.class.getResource("/cal-btn.png")));
		dateChooser_newTo.setBackground(Color.WHITE);
		dateChooser_newTo.setBounds(302, 23, 102, 25);
		panel_11.add(dateChooser_newTo);
		dateChooser_newTo.setEnabled(false);
		dateChooser_newTo.getCalendarButton().addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				Calendar cal3 = Calendar.getInstance();
				cal3.set(Calendar.HOUR_OF_DAY, 23);
				cal3.set(Calendar.MINUTE, 59);
				cal3.set(Calendar.SECOND, 59);
				Date enddate = cal3.getTime();
				dateChooser_newTo.setMaxSelectableDate(enddate);
				try {
					Calendar calendarstartdate = dateChooser_newFrom.getCalendar();
					calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
					calendarstartdate.set(Calendar.MINUTE, 00);
					calendarstartdate.set(Calendar.SECOND, 00);
					dateChooser_newTo.setMinSelectableDate(calendarstartdate.getTime());
				} catch (Error e1) {
					logger.warning("Error : " + e1.getMessage() + System.lineSeparator());
				} catch (Exception e1) {
					logger.warning("Exception : " + e1.getMessage() + System.lineSeparator());
					return;
				}

			}
		});
		dateChooser_newTo.setDateFormatString("dd-MMM-yyyy");

		addButton = new JButton("Add");
		addButton.setBounds(417, 54, 89, 23);
		panel_11.add(addButton);

		removeButton = new JButton("Remove");
		removeButton.setBounds(417, 99, 89, 23);
		panel_11.add(removeButton);
		removeButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				DefaultTableModel model = (DefaultTableModel) table.getModel();
				int[] rows = table.getSelectedRows();

				if (table.getRowCount() < 1) {
					JOptionPane.showMessageDialog(Main_Frame.this, "Please add date to remove", messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE, null);
					return;
				}
				if (rows.length < 1) {
					JOptionPane.showMessageDialog(Main_Frame.this, "Please add date to remove", messageboxtitle,
							JOptionPane.INFORMATION_MESSAGE, null);
					return;
				}
				for (int i = 0; i < rows.length; i++) {
					model.removeRow(rows[i] - i);
				}

			}
		});
		addButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				new Date();

				if (!fromdateeditor.getText().isEmpty() && !todateeditor.getText().isEmpty()) {
					from = fromdateeditor.getText();
					to = todateeditor.getText();
					String tempFrom, tempTo;
					boolean flag = true;

					if (dateChooser_newTo.getDate().after(dateChooser_newFrom.getDate())) {
						for (int i = 0; i < table.getModel().getRowCount(); i++) {
							tempFrom = table.getModel().getValueAt(i, 0) + "";
							tempTo = table.getModel().getValueAt(i, 1) + "";
							if (from.equals(tempFrom) && to.equals(tempTo))
								flag = false;
						}
						if (flag)
							((DefaultTableModel) table.getModel()).addRow(new Object[] { from, to });
						else
							JOptionPane.showMessageDialog(Main_Frame.this,
									"Duplicate date cannot be added in the table ,Please add valid date",
									messageboxtitle, JOptionPane.INFORMATION_MESSAGE);
					} else {
						JOptionPane.showMessageDialog(Main_Frame.this, "Invalid Date, Please add valid date",
								messageboxtitle, JOptionPane.INFORMATION_MESSAGE);
					}

				}

			}
		});
		todateeditor = (JTextFieldDateEditor) dateChooser_newTo.getDateEditor();
		fromdateeditor = (JTextFieldDateEditor) dateChooser_newFrom.getDateEditor();

		label_12 = new JLabel("");
		label_12.setBounds(936, 48, 81, 28);
		panel_progress.add(label_12);

		label_9 = new JLabel("");
		label_9.setIcon(new ImageIcon(Main_Frame.class.getResource("/bottom.png")));
		label_9.setBounds(0, 552, 1075, 75);
		panel_3.add(label_9);
		btnStop.setToolTipText("Click here to Stop the Conversion. ");
		btnStop.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				String warn = "Do you want to stop the process?";
				int ans = JOptionPane.showConfirmDialog(Main_Frame.this, warn, messageboxtitle,
						JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {

					stop = true;
				}
			}
		});
		btn_converter.setToolTipText("Click here to Begin the Conversion.");
		btn_converter.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {

				count_destination = 0;

				// New Date Filter set By Rahul
				SimpleDateFormat df2 = new SimpleDateFormat("dd-MMM-yyyy HH:mm");

				fromList.clear();
				toList.clear();
				if (DateFilter.isSelected()) {
					if (table.getModel().getRowCount() < 1) {
						JOptionPane.showMessageDialog(Main_Frame.this, "Please Add date in the table", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));
						return;
					} else {
						for (int i = 0; i < table.getModel().getRowCount(); i++) {
							Date frmdate, todate;
							try {

								frmdate = df2.parse(table.getValueAt(i, 0).toString() + " 00:00");
								System.out.println(frmdate);
								fromList.add(frmdate);
								todate = df2.parse(table.getValueAt(i, 1).toString() + " 23:59");
								System.out.println(todate);
								toList.add(todate);
							} catch (ParseException e1) {
								logger.severe(e1.getMessage());
								logger.warning("Exception : " + e1.getMessage() + System.lineSeparator());
								e1.printStackTrace();
							}
						}

//					for (int i = 0; i < fromList.size(); i++) {
//						System.out.println("From :" + fromList.get(i));
//						System.out.println("To :" + toList.get(i));
//					}
					}
				}
				if (chckbxSetBackupSchedule.isSelected()) {
					table_fileConvertionreport_panel4
							.setModel(new DefaultTableModel(new Object[][] {}, new String[] { "From", "To", "Status",
									"Duration", "Message Count", "Path", "Last Runtime", "Next RunTime" }));

				}
				if (chckbx_calender_box.isSelected()) {
					try {
						Calendar calendarstartdate = dateChooser_calender_start.getCalendar();
						Calendar calendarenddate = dateChooser_calendar_end.getCalendar();
						calendarstartdate.set(Calendar.HOUR_OF_DAY, 00);
						calendarstartdate.set(Calendar.MINUTE, 00);
						calendarstartdate.set(Calendar.SECOND, 00);
						calendarenddate.set(Calendar.HOUR_OF_DAY, 23);
						calendarenddate.set(Calendar.MINUTE, 59);
						calendarenddate.set(Calendar.SECOND, 59);
						Calenderfilterstartdate = calendarstartdate.getTime();
						Calenderfilterenddate = calendarenddate.getTime();

					} catch (Exception e) {
						JOptionPane.showMessageDialog(frame,
								"Please Enter the date in Calendar filter before Continuing", messageboxtitle,
								JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

						throw new NullPointerException();
					}

				}
				if (chckbx_Mail_Filter.isSelected()) {
					try {
						Calendar mailstartdate = dateChooser_mail_fromdate.getCalendar();
						Calendar mailenddate = dateChooser_mail_tilldate.getCalendar();
						mailstartdate.set(Calendar.HOUR_OF_DAY, 00);
						mailstartdate.set(Calendar.MINUTE, 00);
						mailstartdate.set(Calendar.SECOND, 00);
						mailenddate.set(Calendar.HOUR_OF_DAY, 23);
						mailenddate.set(Calendar.MINUTE, 59);
						mailenddate.set(Calendar.SECOND, 59);
						mailfilterstartdate = mailstartdate.getTime();

						mailfilterenddate = mailenddate.getTime();

					} catch (Exception e) {
						JOptionPane.showMessageDialog(frame, "Please Enter the date in Mail filter before Continuing",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

						throw new NullPointerException();
					}
				}
				if (task_box.isSelected()) {
					try {
						Calendar taskstartdate = dateChooser_task_start_date.getCalendar();
						Calendar taskenddate = dateChooser_task_end_date.getCalendar();

						taskstartdate.set(Calendar.HOUR_OF_DAY, 00);
						taskstartdate.set(Calendar.MINUTE, 00);
						taskstartdate.set(Calendar.SECOND, 00);
						taskenddate.set(Calendar.HOUR_OF_DAY, 23);
						taskenddate.set(Calendar.MINUTE, 59);
						taskenddate.set(Calendar.SECOND, 59);

						taskfilterstartdate = taskstartdate.getTime();
						taskfilterenddate = taskenddate.getTime();
					} catch (Exception e) {
						JOptionPane.showMessageDialog(frame, "Please Enter the date in task filter before Continuing",
								messageboxtitle, JOptionPane.ERROR_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));

						throw new NullPointerException();
					}
				}

				if (chckbxCustomFolderName.isSelected() && textField_customfolder.getText().isEmpty()) {

					JOptionPane.showMessageDialog(frame, "Please Enter the name of folder before Continuing",
							messageboxtitle, JOptionPane.ERROR_MESSAGE,
							new ImageIcon(Main_Frame.class.getResource("/information.png")));

					throw new NullPointerException();

				}

				Desktop desktop = Desktop.getDesktop();
				th = new Thread(new Runnable() {

					@SuppressWarnings("resource")
					@Override
					public void run() {
						btn_Destination.setEnabled(false);
						btn_previous_p3.setEnabled(false);
						btn_converter.setEnabled(false);
						btnStop.setVisible(true);
						chckbxSavePdfAttachment.setEnabled(false);
						chckbx_convert_pdf_to_pdf.setEnabled(false);
						panel_5.setEnabled(false);
						chckbxSaveInSame.setEnabled(false);
						// chckbxMigrateOrBackup.setVisible(false);
						textField_customfolder.setEditable(false);
						chckbxCustomFolderName.setEnabled(false);
						dateChooser_calender_start.setEnabled(false);
						dateChooser_calendar_end.setEnabled(false);
						chckbxMaintainFolderHeirachy.setEnabled(false);
						chckbxRestoreToDefault.setEnabled(false);
						chckbxDeleteEmailFrom.setEnabled(false);
						chckbxMigrateOrBackup.setEnabled(false);
						chckbxAutoIncrementBackup.setEnabled(false);
						dateChooser_mail_fromdate.setEnabled(false);
						dateChooser_mail_tilldate.setEnabled(false);
						comboBox.setEnabled(false);
						chckbx_splitpst.setEnabled(false);
						lbl_progressreport.setText("");
						task_box.setEnabled(false);
						chckbxRemoveDuplicacy.setEnabled(false);
						chckbx_Mail_Filter.setEnabled(false);
						dateChooser_task_start_date.setEnabled(false);
						dateChooser_task_end_date.setEnabled(false);
						chckbx_Mail_Filter.setEnabled(false);
						chckbx_calender_box.setEnabled(false);
						task_box.setEnabled(false);
						chckbx_calender_box.setEnabled(false);
						btn_signout_p3.setVisible(false);
						comboBox_fileDestination_type.setEnabled(false);
						long starttime = System.currentTimeMillis();
						textField_customfolder.setEditable(false);
						chckbxSetBackupSchedule.setEnabled(false);
						label_14.setEnabled(false);
						dateChooserNextSchedular.setEnabled(false);
						spinner.setEnabled(false);
						rdbtnOnce.setEnabled(false);
						rdbtnEveryWeek.setEnabled(false);
						rdbtnEveryday.setEnabled(false);
						rdbtnOnWeekDay.setEnabled(false);
						rdbtnOnmonthDay.setEnabled(false);
						rdbtnEveryMonth.setEnabled(false);
						spinner_sizespinner.setEnabled(false);
						comboBox_setsize.setEnabled(false);

						String destinationfile = "";
						Progressbar.setVisible(true);

						if (fileoption.equalsIgnoreCase("MBOX")) {
							String fnam = file.getName();
							fnam = fnam.replaceAll("\\p{C}", "_");
							fnam = fnam.replace(".mbox", "").replace(".mbx", "");
							fname = fnam.trim();
							fname = getRidOfIllegalFileNameCharacters(fname);

						} else if (fileoption.equalsIgnoreCase("DBX")) {
							String fnam = file.getName();
							fnam = getRidOfIllegalFileNameCharacters(fnam);
							fnam = fnam.replace(".dbx", "");
							fnam = fnam.replaceAll("\\p{C}", "_");

							fname = fnam.trim();

						} else if (fileoption.equalsIgnoreCase("Zimbra files (.tgz)")) {
							String fnam = file.getName();
							fnam = getRidOfIllegalFileNameCharacters(fnam);
							fnam = fnam.replace(".tgz", "");
							fnam = fnam.replaceAll("\\p{C}", "_");

							fname = fnam.trim();

						} else if (fileoption.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
							String fnam = file.getName();
							fnam = getRidOfIllegalFileNameCharacters(fnam);
							fnam = fnam.replace(".pst", "");
							fnam = fnam.replaceAll("\\p{C}", "_");

							fname = fnam.trim();

						} else if (fileoption.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {
							String fnam = file.getName();
							fnam = fnam.replaceAll("\\p{C}", "_");
							fnam = fnam.replace(".ost", "");
							fname = fnam.trim();
							fname = getRidOfIllegalFileNameCharacters(fname);
						} else if (fileoption.equalsIgnoreCase("EML File (.eml)")) {
							String fnam = file.getName();
							fname = fnam.replace(".eml", "").replace("//s", "");
							fname = getRidOfIllegalFileNameCharacters(fname);
						} else if (fileoption.equalsIgnoreCase("EMLX File (.emlx)")) {
							String fnam = file.getName();
							fnam = fnam.replaceAll("\\p{C}", "_");
							fnam = fnam.replace(".emlx", "");
							fname = fnam.trim();
							fname = getRidOfIllegalFileNameCharacters(fname);
						} else if (fileoption.equalsIgnoreCase("OFT File (.oft)")) {
							String fnam = file.getName();
							fnam = fnam.replaceAll("\\p{C}", "_");
							fnam = fnam.replace(".oft", "");
							fname = fnam.trim();
							fname = getRidOfIllegalFileNameCharacters(fname);
						} else if (fileoption.equalsIgnoreCase("Maildir")) {
							String fnam = file.getName();
							fnam = fnam.replaceAll("\\p{C}", "_");

							fname = fnam.trim();
							fname = getRidOfIllegalFileNameCharacters(fname);
						} else if (fileoption.equalsIgnoreCase("Message File (.msg)")) {
							String fnam = file.getName();
							fnam = fnam.replaceAll("\\p{C}", "_");
							fnam = fnam.replace(".msg", "");
							fname = fnam.trim();
							fname = getRidOfIllegalFileNameCharacters(fname);
						} else if (fileoption.equalsIgnoreCase("OLM File (.olm)")) {
							String fnam = file.getName();
							fnam = fnam.replace(".olm", "");
							fnam = fnam.replaceAll("\\p{C}", "_");
							fname = fnam.trim();
							fname = getRidOfIllegalFileNameCharacters(fname);
						}

						if (checkconvertagain) {
							if (!(filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("YAHOO MAIL")
									|| filetype.equalsIgnoreCase("AOL") || filetype.equalsIgnoreCase("Live Exchange")
									|| filetype.equalsIgnoreCase("Amazon WorkMail")
									|| filetype.equalsIgnoreCase("Zoho Mail")
									|| filetype.equalsIgnoreCase("Yandex Mail")
									|| filetype.equalsIgnoreCase("Hostgator email")
									|| filetype.equalsIgnoreCase("Icloud") || filetype.equalsIgnoreCase("GoDaddy email")
									|| filetype.equalsIgnoreCase("OFFICE 365") || filetype.equalsIgnoreCase("IMAP")
									|| filetype.equalsIgnoreCase("HOTMAIL"))) {

								if (fileoption.equalsIgnoreCase("Gmail") || fileoption.equalsIgnoreCase("Yahoo Mail")
										|| fileoption.equalsIgnoreCase("OFFICE 365")
										|| fileoption.equalsIgnoreCase("IMAP")
										|| fileoption.equalsIgnoreCase("Opera Mail")
										|| fileoption.equalsIgnoreCase("Amazon WorkMail")
										|| fileoption.equalsIgnoreCase("Yandex Mail")
										|| fileoption.equalsIgnoreCase("Thunderbird")
										|| fileoption.equalsIgnoreCase("Hostgator email")
										|| fileoption.equalsIgnoreCase("Icloud")
										|| fileoption.equalsIgnoreCase("GoDaddy email")
										|| fileoption.equalsIgnoreCase("Apple Mail")
										|| fileoption.equalsIgnoreCase("AOL")
										|| fileoption.equalsIgnoreCase("Zoho Mail")
										|| fileoption.equalsIgnoreCase("Live Exchange")
										|| fileoption.equalsIgnoreCase("Hotmail")) {

									if (chckbxCustomFolderName.isSelected()) {

										String customerfolder = textField_customfolder.getText().replace("//s", "");

										customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
										String servername = getRidOfIllegalFileNameCharacters(
												textField_username_p2.getText());
										f = new File(tf_Destination_Location.getText() + File.separator + customerfolder
												+ File.separator + servername + "_" + filetype + "_"
												+ comboBox.getSelectedItem().toString());

										if (filetype.equalsIgnoreCase("Thunderbird")) {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ customerfolder + ".sbd" + File.separator + servername + "_"
													+ filetype + "_" + comboBox.getSelectedItem().toString() + ".sbd");
										} else {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ customerfolder + File.separator + servername + "_" + filetype
													+ "_" + comboBox.getSelectedItem().toString());
										}

										if (!f.isFile()) {

											f.mkdirs();
											if (filetype.equalsIgnoreCase("Thunderbird")) {
												new MboxrdStorageWriter(f.getAbsolutePath() + f.getName(), false);

											}

											destination_path = f.getAbsolutePath();

											destinationfile = f.getAbsolutePath();

										} else {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ customerfolder + ".sbd" + "(" + calendertime + ")" + ".sbd"
													+ File.separator + servername + "_" + filetype + "_"
													+ comboBox.getSelectedItem().toString() + ".sbd");

											f.mkdirs();

											if (filetype.equalsIgnoreCase("Thunderbird")) {
												new MboxrdStorageWriter(f.getAbsolutePath() + f.getName(), false);

											}
											destination_path = f.getAbsolutePath();

											destinationfile = f.getAbsolutePath();

										}

									} else {
										if (filetype.equalsIgnoreCase("Thunderbird")) {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ calendertime + ".sbd" + File.separator
													+ textField_username_p2.getText() + filetype + ".sbd");

											f.mkdirs();
											if (filetype.equalsIgnoreCase("Thunderbird")) {

												new MboxrdStorageWriter(f.getAbsolutePath() + f.getName(), false);

											}
											destination_path = f.getAbsolutePath();
										} else {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ calendertime + File.separator + textField_username_p2.getText()
													+ filetype);

											f.mkdirs();
											destination_path = f.getAbsolutePath();
										}
									}

								} else {

									if (chckbxCustomFolderName.isSelected()) {

										String customerfolder = textField_customfolder.getText().replace("//s", "");

										customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);

										f = new File(tf_Destination_Location.getText() + File.separator + customerfolder
												+ filetype + "_" + filetype + "_"
												+ comboBox.getSelectedItem().toString());

										if (!f.isFile()) {

											f.mkdirs();

											destination_path = f.getAbsolutePath();

											destinationfile = f.getAbsolutePath();

										} else {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ customerfolder + "(" + calendertime + ")" + "_"
													+ comboBox.getSelectedItem().toString());

											f.mkdirs();

											destination_path = f.getAbsolutePath();

											destinationfile = f.getAbsolutePath();

										}

									} else {
										calendertime = getRidOfIllegalFileNameCharacters(calendertime);
										f = new File(tf_Destination_Location.getText() + File.separator + calendertime
												+ File.separator + fname + "_" + filetype + "_"
												+ comboBox.getSelectedItem().toString());
										f.mkdirs();

										destination_path = f.getAbsolutePath();

										destinationfile = f.getAbsolutePath();

									}

								}
							}

						} else {
							if (!(filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("YAHOO MAIL")
									|| filetype.equalsIgnoreCase("AOL") || filetype.equalsIgnoreCase("Live Exchange")
									|| filetype.equalsIgnoreCase("Amazon WorkMail")
									|| filetype.equalsIgnoreCase("Zoho Mail")
									|| filetype.equalsIgnoreCase("Yandex Mail")
									|| filetype.equalsIgnoreCase("Hostgator email")
									|| filetype.equalsIgnoreCase("Icloud") || filetype.equalsIgnoreCase("GoDaddy email")
									|| filetype.equalsIgnoreCase("OFFICE 365") || filetype.equalsIgnoreCase("IMAP")
									|| filetype.equalsIgnoreCase("HOTMAIL"))) {

								if (fileoption.equalsIgnoreCase("Gmail") || fileoption.equalsIgnoreCase("Yahoo Mail")
										|| fileoption.equalsIgnoreCase("OFFICE 365")
										|| fileoption.equalsIgnoreCase("Amazon WorkMail")
										|| fileoption.equalsIgnoreCase("IMAP")
										|| fileoption.equalsIgnoreCase("Hostgator email")
										|| fileoption.equalsIgnoreCase("Icloud")
										|| fileoption.equalsIgnoreCase("GoDaddy email")
										|| fileoption.equalsIgnoreCase("Zoho Mail")
										|| fileoption.equalsIgnoreCase("Yandex Mail")
										|| fileoption.equalsIgnoreCase("Opera Mail")
										|| fileoption.equalsIgnoreCase("Thunderbird")
										|| fileoption.equalsIgnoreCase("Apple Mail")
										|| fileoption.equalsIgnoreCase("AOL")
										|| fileoption.equalsIgnoreCase("Live Exchange")
										|| fileoption.equalsIgnoreCase("Hotmail")) {

									if (chckbxCustomFolderName.isSelected()) {

										String customerfolder = textField_customfolder.getText().replace("//s", "");

										customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
										String servername = getRidOfIllegalFileNameCharacters(
												textField_username_p2.getText());
										f = new File(tf_Destination_Location.getText() + File.separator + customerfolder
												+ File.separator + servername);

										if (!f.isFile()) {

											f.mkdirs();

											destination_path = f.getAbsolutePath();

											destinationfile = f.getAbsolutePath();

										} else {
											f = new File(
													tf_Destination_Location.getText() + File.separator + customerfolder
															+ "(" + calendertime + ")" + File.separator + servername);

											f.mkdirs();

											destination_path = f.getAbsolutePath();

											destinationfile = f.getAbsolutePath();

										}

									} else {

										f = new File(tf_Destination_Location.getText() + File.separator + calendertime
												+ File.separator
												+ getRidOfIllegalFileNameCharacters(textField_username_p2.getText()));
										f.mkdirs();
										destination_path = tf_Destination_Location.getText() + File.separator
												+ calendertime + File.separator
												+ getRidOfIllegalFileNameCharacters(textField_username_p2.getText());
										destinationfile = tf_Destination_Location.getText() + File.separator
												+ calendertime + File.separator
												+ getRidOfIllegalFileNameCharacters(textField_username_p2.getText());

									}

								} else {

									if (chckbxCustomFolderName.isSelected()) {

										String customerfolder = textField_customfolder.getText().replace("//s", "");

										customerfolder = getRidOfIllegalFileNameCharacters(customerfolder);
										f = new File(
												tf_Destination_Location.getText() + File.separator + customerfolder);

										if (!f.isFile()) {

											f.mkdirs();

											destination_path = f.getAbsolutePath();

											destinationfile = f.getAbsolutePath();

										} else {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ customerfolder + "(" + calendertime + ")");

											f.mkdirs();

											destination_path = f.getAbsolutePath();

											destinationfile = f.getAbsolutePath();

										}

									} else {

										if (filetype.equalsIgnoreCase("Thunderbird")) {

											File fd = new File(tf_Destination_Location.getText() + File.separator
													+ calendertime + ".sbd");

											System.out.println(fd.getAbsolutePath());
											fd.mkdirs();
											System.out.println(fd.exists());

											// if

											new MboxrdStorageWriter(fd.getAbsolutePath(), false);

											fd = new File(fd.getAbsolutePath() + File.separator + fname + ".sbd");

											System.out.println(fd.getAbsolutePath());
											fd.mkdirs();
											System.out.println(fd.exists());

											new MboxrdStorageWriter(fd.getAbsolutePath(), false);
											f = new File(tf_Destination_Location.getText() + File.separator
													+ calendertime + File.separator + fname + ".sbd");

										} else {
											f = new File(tf_Destination_Location.getText() + File.separator
													+ calendertime + File.separator + fname);
										}

										f.mkdirs();
										if (filetype.equalsIgnoreCase("Thunderbird")) {
											try {
												new MboxrdStorageWriter(
														tf_Destination_Location.getText() + File.separator
																+ calendertime + File.separator + fname + ".sbd",
														false);
											} catch (Exception e) {

											}

										}
										destination_path = f.getAbsolutePath();

										destinationfile = f.getAbsolutePath();

									}

								}

							}

						}

						try {
							if (chckbx_Mail_Filter.isSelected()
									&& (mailfilterenddate == null || mailfilterstartdate == null)) {
								JOptionPane.showMessageDialog(frame, "Please Select Start and End Date",
										messageboxtitle, JOptionPane.ERROR_MESSAGE,
										new ImageIcon(Main_Frame.class.getResource("/information.png")));
							} else if (chckbx_calender_box.isSelected()
									&& (Calenderfilterenddate == null || Calenderfilterstartdate == null)) {
								JOptionPane.showMessageDialog(frame, "Please Select Start and End Date",
										messageboxtitle, JOptionPane.ERROR_MESSAGE,
										new ImageIcon(Main_Frame.class.getResource("/information.png")));
							} else {
								if (nextTime != null) {

									if (filetype.equalsIgnoreCase("GMAIL") || filetype.equalsIgnoreCase("YAHOO MAIL")
											|| filetype.equalsIgnoreCase("AOL")
											|| filetype.equalsIgnoreCase("Amazon WorkMail")
											|| filetype.equalsIgnoreCase("Zoho Mail")
											|| filetype.equalsIgnoreCase("Yandex Mail")
											|| filetype.equalsIgnoreCase("Hostgator email")
											|| filetype.equalsIgnoreCase("Icloud")
											|| filetype.equalsIgnoreCase("GoDaddy email")
											|| filetype.equalsIgnoreCase("IMAP")) {

										path = nextPAth;

									} else {

										new File(destination_path).delete();
										new File(removefolder(destination_path)).delete();
										destination_path = nextPAth.replace("/", File.separator);
										new File(destination_path).mkdirs();
										f = new File(destination_path);

									}
//

									filetype = nextfiletype;
									table_fileConvertionreport_panel4.setModel(new DefaultTableModel(new Object[][] {},
											new String[] { "From", "To", "Status", "Duration", "Message Count", "Path",
													"Last Runtime", "Next RunTime" }));
									for (int i = 0; i < listnextendtime.size(); i++) {
										mode = (DefaultTableModel) table_fileConvertionreport_panel4.getModel();
										Calendar cs = Calendar.getInstance();
										cs.setTimeInMillis(listnextstarttime.get(i));
										Calendar cs1 = Calendar.getInstance();
										cs1.setTimeInMillis(listnextendtime.get(i));

										mode.addRow(new Object[] { fileoption, filetype, Status,
												listnextduration.get(i), listnextcount.get(i), destination_path,
												cs.getTime(), cs1.getTime() });

									}

								}
								spinner.updateUI();
								spinner_sizespinner.updateUI();
								maxsize = 0;

								if (chckbx_splitpst.isSelected()) {
									if (comboBox_setsize.getSelectedIndex() == 0) {
										Object o = spinner_sizespinner.getValue();
										Number n = (Number) o;
										maxsize = (n.longValue()) * (1000 * 1000);

									} else if (comboBox_setsize.getSelectedIndex() == 1) {
										Object o = spinner_sizespinner.getValue();
										Number n = (Number) o;
										maxsize = (n.longValue()) * (1000 * 1000 * 1000);
									}
								}
								if (fileoption.equalsIgnoreCase("MBOX")) {
								} else if (fileoption.equalsIgnoreCase("Zimbra files (.tgz)")) {
								} else if (fileoption.equalsIgnoreCase("Opera Mail")
										|| fileoption.equalsIgnoreCase("Thunderbird")
										|| fileoption.equalsIgnoreCase("Apple Mail")) {
								}
//Here 
								else if (fileoption.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")
										|| fileoption.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {
								} else if (fileoption.equalsIgnoreCase("Live Exchange")
										|| fileoption.equalsIgnoreCase("OFFICE 365")
										|| fileoption.equalsIgnoreCase("Hotmail")) {
								}

								else if (fileoption.equalsIgnoreCase("Yahoo Mail")
										|| fileoption.equalsIgnoreCase("Gmail") || fileoption.equalsIgnoreCase("AOL")
										|| fileoption.equalsIgnoreCase("IMAP")
										|| fileoption.equalsIgnoreCase("Hostgator email")
										|| fileoption.equalsIgnoreCase("Icloud")
										|| fileoption.equalsIgnoreCase("GoDaddy email")
										|| fileoption.equalsIgnoreCase("Amazon WorkMail")
										|| fileoption.equalsIgnoreCase("Yandex Mail")
										|| fileoption.equalsIgnoreCase("Zoho Mail")) {
								}

								else if (fileoption.equalsIgnoreCase("OLM File (.olm)")) {
								} else {
								}

							}

							if (filetype.equalsIgnoreCase("THUNDERBIRD")) {
								JOptionPane.showMessageDialog(frame,
										"Please open the converted file from " + destination_path + " Thunderbird",
										messageboxtitle, JOptionPane.INFORMATION_MESSAGE);
							}
							if (filetype.equalsIgnoreCase("YAHOO MAIL") || filetype.equalsIgnoreCase("GMAIL")
									|| filetype.equalsIgnoreCase("OFFICE 365") || filetype.equals("AOL")
									|| filetype.equalsIgnoreCase("Yandex Mail")
									|| filetype.equalsIgnoreCase("Amazon WorkMail")
									|| filetype.equalsIgnoreCase("Hostgator email")
									|| filetype.equalsIgnoreCase("Icloud") || filetype.equalsIgnoreCase("GoDaddy email")
									|| filetype.equalsIgnoreCase("Live Exchange")
									|| filetype.equalsIgnoreCase("Yandex Mail") || filetype.equalsIgnoreCase("hotmail")
									|| filetype.equalsIgnoreCase("IMAP") || filetype.equalsIgnoreCase("Zoho Mail")) {

								if (filetype.equalsIgnoreCase("YAHOO MAIL")) {

									reportpath = "http://login.yahoo.com";
									openBrowser(reportpath);
								} else if (filetype.equalsIgnoreCase("Yandex Mail")) {

									reportpath = "https://mail.yandex.com/?uid=1213147137#tabs/relevant";
									openBrowser(reportpath);
								} else if (filetype.equalsIgnoreCase("GMAIL")) {

									reportpath = "https://mail.google.com";
									openBrowser(reportpath);
								} else if (filetype.equalsIgnoreCase("GoDaddy email")) {

									reportpath = "https://sso.godaddy.com/login?app=email&realm=pass";
									openBrowser(reportpath);
								} else if (filetype.equalsIgnoreCase("Icloud")) {

									reportpath = "https://www.icloud.com/mail";
									openBrowser(reportpath);
								} else if (filetype.equalsIgnoreCase("Hostgator email")) {

									reportpath = "https://www.hostgator.in/login.php";
									openBrowser(reportpath);
								} else if (filetype.equalsIgnoreCase("Zoho Mail")) {

									reportpath = "https://accounts.zoho.in/signin?servicename=VirtualOffice&signupurl=https://www.zoho.in/mail/zohomail-pricing.html&serviceurl=https://mail.zoho.in";
									openBrowser(reportpath);
								} else if (filetype.equals("AOL")) {

									reportpath = "https://login.aol.com";
									openBrowser(reportpath);
								} else if (filetype.equalsIgnoreCase("Live Exchange")) {

									JOptionPane.showMessageDialog(frame,
											"Please open the converted file from Live Exchange", messageboxtitle,
											JOptionPane.INFORMATION_MESSAGE);
								} else if (filetype.equalsIgnoreCase("Hotmail")) {

									reportpath = "https://outlook.live.com";
									openBrowser(reportpath);
								} else if (filetype.equalsIgnoreCase("Amazon WorkMail")) {

									JOptionPane.showMessageDialog(frame,
											"Please open the converted file from  Amazon WorkMail", messageboxtitle,
											JOptionPane.INFORMATION_MESSAGE);
								} else if (filetype.equalsIgnoreCase("IMAP")) {

									JOptionPane.showMessageDialog(frame, "Please open the converted file from  IMAP",
											messageboxtitle, JOptionPane.INFORMATION_MESSAGE);
								} else {

									reportpath = "https://outlook.office365.com";
									openBrowser(reportpath);
								}

							} else {
								reportpath = f.getAbsolutePath();
								desktop.open(f);

							}

						} catch (Error e) {
							logger.warning("Error : " + e.getMessage() + System.lineSeparator());
						} catch (Exception e) {
							logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
							// //System.out.println(e.getMessage());
							e.printStackTrace();

						} finally {

							// progressBar_message_p3.setVisible(false);
							Progressbar.setVisible(false);
							Calendar cio = Calendar.getInstance();

							String duration = Duration(starttime).toString();
							CardLayout card = (CardLayout) Cardlayout.getLayout();
							card.show(Cardlayout, "panel_4");

							logger.info("File Saved " + count_destination + System.lineSeparator() + "End Time : "
									+ cio.getTime() + System.lineSeparator()
									+ "**********************************************************");

							mode = (DefaultTableModel) table_fileConvertionreport_panel4.getModel();
							if (chckbxSetBackupSchedule.isSelected()) {
								btnConvertAgain.setVisible(false);
								btnDowloadReport.setVisible(false);

								if (rdbtnEveryday.isSelected()) {
									Date startdate = java.util.Calendar.getInstance().getTime();
									dateChooserNextSchedular.setDate(startdate);

									Calendar cal = dateChooserNextSchedular.getCalendar();
									cal.add(Calendar.DATE, 1);

									dateChooserNextSchedular.setDate(cal.getTime());

								} else if (rdbtnEveryMonth.isSelected()) {
									Date startdate = java.util.Calendar.getInstance().getTime();
									dateChooserNextSchedular.setDate(startdate);
									Calendar cal = dateChooserNextSchedular.getCalendar();
									cal.add(Calendar.MONTH, 1);

									dateChooserNextSchedular.setDate(cal.getTime());

								} else if (rdbtnEveryWeek.isSelected()) {
									Date startdate = java.util.Calendar.getInstance().getTime();
									dateChooserNextSchedular.setDate(startdate);
									Calendar cal = dateChooserNextSchedular.getCalendar();
									cal.add(Calendar.DATE, 7);
									dateChooserNextSchedular.setDate(cal.getTime());
								} else if (rdbtnOnmonthDay.isSelected()) {
									Date startdate = java.util.Calendar.getInstance().getTime();
									dateChooserNextSchedular.setDate(startdate);

									Calendar cal = dateChooserNextSchedular.getCalendar();
									int calenderdate = comboBox_MonthDay.getSelectedIndex() + 1;
									cal.set(Calendar.DAY_OF_MONTH, calenderdate);
									if (cal.get(Calendar.DAY_OF_MONTH) >= calenderdate) {
										cal.add(Calendar.MONTH, 1);
									}

									dateChooserNextSchedular.setDate(cal.getTime());
								} else if (rdbtnOnWeekDay.isSelected()) {
									Date startdate = java.util.Calendar.getInstance().getTime();
									dateChooserNextSchedular.setDate(startdate);

									Calendar cal = dateChooserNextSchedular.getCalendar();
									int in = (comboBox_weekdays.getSelectedIndex() + 1) - cal.get(Calendar.DAY_OF_WEEK);

									if (in <= 0) {
										in += 7;
									}

									cal.add(Calendar.DAY_OF_MONTH, in);

									dateChooserNextSchedular.setDate(cal.getTime());
								}

								Date nextdate = (Date) spinner.getValue();

								caltime = Calendar.getInstance();

								Calendar cll = dateChooserNextSchedular.getCalendar();

								caltime.setTime(nextdate);
								caltime.set(Calendar.DAY_OF_MONTH, cll.get(Calendar.DAY_OF_MONTH));
								caltime.set(Calendar.YEAR, cll.get(Calendar.YEAR));
								caltime.set(Calendar.MONTH, cll.get(Calendar.MONTH));
								Date dnext = caltime.getTime();
								System.out.println(dnext);
								for (int i = 0; i < listnextendtime.size(); i++) {
									mode = (DefaultTableModel) table_fileConvertionreport_panel4.getModel();
									Calendar cs = Calendar.getInstance();
									cs.setTimeInMillis(listnextstarttime.get(i));
									Calendar cs1 = Calendar.getInstance();
									cs1.setTimeInMillis(listnextendtime.get(i));

									mode.addRow(new Object[] { fileoption, filetype, Status, listnextduration.get(i),
											listnextcount.get(i), destination_path, cs.getTime(), cs1.getTime() });

								}

								mode.addRow(new Object[] { fileoption, filetype, Status, duration, count_destination,
										reportpath, calendertime, dnext.toString() });
								nextPAth = destination_path;

								try {
									once = String.valueOf(rdbtnOnce.isSelected());
									everyday = String.valueOf(rdbtnEveryday.isSelected());
									everyweek = String.valueOf(rdbtnEveryWeek.isSelected());
									OnWeekday = String.valueOf(rdbtnOnWeekDay.isSelected());
									everymonth = String.valueOf(rdbtnEveryMonth.isSelected());
									OnMonthday = String.valueOf(rdbtnOnmonthDay.isSelected());
									removeduplica = String.valueOf(chckbxRemoveDuplicacy.isSelected());
									maintainfolderh = String.valueOf(chckbxMaintainFolderHeirachy.isSelected());
									savepdfattac = String.valueOf(chckbxSavePdfAttachment.isSelected());
									freeupserverspace = String.valueOf(chckbxDeleteEmailFrom.isSelected());

									String sql = "Insert INTO Schedule_Detail" + "(Username," + "Password,"
											+ "Startime," + "NextEndTime," + "BtnOnce," + "BtnEveryDay,"
											+ "BtneveryWeek," + "BtnOnweekday," + "BtnEveryMonth," + "username_output,"
											+ "password_output," + "domain_p2," + "domain_p3," + "BtnOnmonthday,"
											+ "fileoption," + "nextfiletype," + "nextPAth," + "removeduplica,"
											+ "maintainfolderh," + "savepdfattac," + "freeupserverspace," + "nextcount,"
											+ "Duration)" +

											"VALUES ('" + username_p2 + "', '" + password_p2 + "',"
											+ cal.getTimeInMillis() + "," + caltime.getTimeInMillis() + ",'" + once
											+ "','" + everyday + "','" + everyweek + "','" + OnWeekday + "','"
											+ everymonth + "','" + username_p3 + "','" + password_p3 + "','" + domain_p2
											+ "','" + domain_p3 + "','" + OnMonthday + "','" + fileoption + "','"
											+ filetype + "','" + destination_path + "','" + removeduplica + "','"
											+ maintainfolderh + "','" + savepdfattac + "','" + freeupserverspace + "',"
											+ count_destination + ",'" + duration + "');";
									schsqlstmt.executeUpdate(sql);
									schsqlstmt.close();
								} catch (SQLException e1) {

									e1.printStackTrace();
								}
								System.out.println("Records created successfully");

								long start_time = System.currentTimeMillis();
								System.out.println(start_time);
								long end_time = caltime.getTimeInMillis();
								long difference = end_time - start_time;

								System.out.println("waiting");

								lblNextMigrationStart.setVisible(true);
								long i1 = TimeUnit.MILLISECONDS.toSeconds(difference);
								if (trayIcon != null) {
									trayIcon.displayMessage(
											"Next Migration Start On " + Long.valueOf(i1).toString() + " seconds ", " ",
											TrayIcon.MessageType.NONE);
								}

								while (i1 > 0) {
									if (nextstart) {
										nextstart = false;
										break;
									}
									Long day = TimeUnit.SECONDS.toDays(i1);

									long p1 = i1 % 60;
									long p2 = i1 / 60;
									long p3 = p2 % 60;
									p2 = p2 / 60;

									String s = null;
									if (day > 0) {
										s = "Remaining Days for Next Migration: " + day;
										System.out.print(s);
										lblNextMigrationStart.setText(s);
									} else {
										s = "Remaining Time  for Next Migration " + p2 + " Hrs:" + p3 + " Min:" + p1
												+ " Sec";
										System.out.print(s);
										lblNextMigrationStart.setText(s);
									}

									try {

										Thread.sleep(1000L);
										start_time = System.currentTimeMillis();

										difference = end_time - start_time;

										System.out.println(difference);

										lblNextMigrationStart.setVisible(true);
										i1 = TimeUnit.MILLISECONDS.toSeconds(difference);
									} catch (InterruptedException e) {

									}
								}

								lblNextMigrationStart.setText("Backup about to start ");
								if (trayIcon != null) {
									trayIcon.displayMessage(" ", messageboxtitle + " Backup about to start",
											TrayIcon.MessageType.NONE);
								}

								connectionHandle1();
								btn_converter.setEnabled(true);
								chckbxAutoIncrementBackup.setSelected(true);
								btn_converter.doClick();

								card.show(Cardlayout, "panel_3");

								if (rdbtnOnce.isSelected()) {
									chckbxSetBackupSchedule.setSelected(false);
									lblNextMigrationStart.setVisible(false);

								} else if (rdbtnEveryday.isSelected()) {
									rdbtnEveryday.doClick();
								} else if (rdbtnEveryMonth.isSelected()) {
									rdbtnEveryMonth.doClick();
								} else if (rdbtnEveryWeek.isSelected()) {
									rdbtnEveryWeek.doClick();
								} else if (rdbtnOnmonthDay.isSelected()) {
									rdbtnOnmonthDay.doClick();
								} else if (rdbtnOnWeekDay.isSelected()) {
									rdbtnOnWeekDay.doClick();
								}

							} else {

								chckbxSetBackupSchedule.setSelected(false);
								lblNextMigrationStart.setVisible(false);
								String url = "";
								String path = "";

								try {
									if (System.getProperty("os.name").toLowerCase().contains("windows")) {

										url = "jdbc:sqlite:" + System.getenv("APPDATA") + File.separator + projectTitle
												+ File.separator + hashcode("userdetails");

										path = System.getenv("APPDATA") + File.separator + projectTitle + File.separator
												+ hashcode("userdetails");

									} else {

										url = "jdbc:sqlite:" + System.getProperty("user.home") + File.separator
												+ "Library" + File.separator + "Application Support" + File.separator
												+ projectTitle + File.separator + hashcode("userdetails");

										path = System.getProperty("user.home") + File.separator + "Library"
												+ File.separator + "Application Support" + File.separator + projectTitle
												+ File.separator + hashcode("userdetails");

									}

									System.out.println("Database deleted successfully...");

									schsqlconnection = DriverManager.getConnection(url);

									System.out.println("Opened database successfully");
									schsqlstmt = (Statement) schsqlconnection.createStatement();
									String sql = "DROP TABLE Schedule_Detail";
									schsqlstmt.executeUpdate(sql);

								} catch (Exception e) {

									e.printStackTrace();
								} finally {

									try {
										if (schsqlstmt != null) {
											schsqlstmt.close();
											schsqlconnection.close();
										}
										if (schsqlconnection != null) {
											schsqlconnection.close();
										}

									} catch (Exception e) {
										System.out.println(e.getMessage());
									}

									File file = new File(path);
									recursiveDelete(file);
								}

								if (nexttime == false) {
									Calendar cal = Calendar.getInstance();
									cal.setTimeInMillis(nextTime);
									mode.addRow(new Object[] { fileoption, filetype, Status, duration,
											count_destination, reportpath, cal.getTime() });
								} else {

									JOptionPane.showMessageDialog(frame, "Process has Completed", messageboxtitle,
											JOptionPane.ERROR_MESSAGE,
											new ImageIcon(Main_Frame.class.getResource("/information.png")));

									mode.addRow(new Object[] { fileoption, filetype, Status, duration,
											count_destination, reportpath });
								}
								destination_path = destinationfile;
								btnStartTheMigration.setVisible(false);
								btnStopMigration.setVisible(false);
								btnConvertAgain.setVisible(true);
								btnDowloadReport.setVisible(true);
							}
						}
					}
				});

				th.start();

			}
		});
		btn_Destination.setToolTipText("Click here to Select the destination path. ");

		JLabel lblDestinationPath = new JLabel("Destination Path :");
		lblDestinationPath.setFont(new Font("Tahoma", Font.BOLD, 13));
		lblDestinationPath.setBounds(10, 8, 178, 30);
		panel_3_2.add(lblDestinationPath);

		lblSavesbackupmigrateAs = new JLabel(" Saves/Backup/Migrate As :");
		lblSavesbackupmigrateAs.setForeground(Color.BLUE);
		lblSavesbackupmigrateAs.setBackground(Color.WHITE);
		lblSavesbackupmigrateAs.setFont(new Font("Tahoma", Font.PLAIN, 13));
		lblSavesbackupmigrateAs.setBounds(29, 13, 197, 37);
		panel_3.add(lblSavesbackupmigrateAs);
		btn_Destination.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {

				try {

					destinationPath();
					checkdestination = false;

				} catch (Error e) {
					logger.warning("Error : " + e.getMessage() + System.lineSeparator());
				} catch (Exception e) {
					logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
					e.printStackTrace();

				}

			}
		});

		panel_4 = new JPanel();
		panel_4.setBackground(Color.WHITE);
		Cardlayout.add(panel_4, "panel_4");
		panel_4.setLayout(null);

		lblNextMigrationStart = new JLabel("");
		lblNextMigrationStart.setBackground(Color.WHITE);
		lblNextMigrationStart.setForeground(Color.RED);
		lblNextMigrationStart.setFont(new Font("Tahoma", Font.PLAIN, 13));
		lblNextMigrationStart.setVisible(false);
		lblNextMigrationStart.setBounds(340, 477, 432, 27);
		panel_4.add(lblNextMigrationStart);

		JScrollPane scrollPane_table_panel4 = new JScrollPane();
		scrollPane_table_panel4.setBackground(Color.WHITE);
		scrollPane_table_panel4.setBounds(0, 13, 1075, 397);
		panel_4.add(scrollPane_table_panel4);

		table_fileConvertionreport_panel4 = new JTable() {
			/**
			 *
			 */
			private static final long serialVersionUID = 1L;

			public boolean isCellEditable(int row, int column) {

				return false;
			}
		};
		table_fileConvertionreport_panel4.getTableHeader().setReorderingAllowed(false);

		table_fileConvertionreport_panel4.setModel(new DefaultTableModel(new Object[][] {},
				new String[] { "From", "To", "Status", "Duration", "Message Count", "Path" }));
		table_fileConvertionreport_panel4.getColumnModel().getColumn(0).setPreferredWidth(126);
		table_fileConvertionreport_panel4.getColumnModel().getColumn(1).setPreferredWidth(57);

		table_fileConvertionreport_panel4.getColumnModel().getColumn(2).setPreferredWidth(35);
		table_fileConvertionreport_panel4.getColumnModel().getColumn(3).setPreferredWidth(30);
		table_fileConvertionreport_panel4.getColumnModel().getColumn(4).setPreferredWidth(30);
		table_fileConvertionreport_panel4.getColumnModel().getColumn(5).setPreferredWidth(174);
		scrollPane_table_panel4.setViewportView(table_fileConvertionreport_panel4);

		btnStartTheMigration = new JButton("");
		btnStartTheMigration.setIcon(new ImageIcon(Main_Frame.class.getResource("/migration-early.png")));
		btnStartTheMigration.setRolloverEnabled(false);
		btnStartTheMigration.setRequestFocusEnabled(false);
		btnStartTheMigration.setOpaque(false);
		btnStartTheMigration.setFocusable(false);
		btnStartTheMigration.setFocusTraversalKeysEnabled(false);
		btnStartTheMigration.setFocusPainted(false);
		btnStartTheMigration.setDefaultCapable(false);
		btnStartTheMigration.setContentAreaFilled(false);
		btnStartTheMigration.setBorderPainted(false);
		btnStartTheMigration.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnStartTheMigration.setIcon(new ImageIcon(Main_Frame.class.getResource("/hvr-migration-early.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnStartTheMigration.setIcon(new ImageIcon(Main_Frame.class.getResource("/migration-early.png")));
			}
		});
		btnStartTheMigration.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				nextstart = true;
			}
		});
		btnStartTheMigration.setBounds(760, 469, 220, 39);
		panel_4.add(btnStartTheMigration);

		btnStopMigration = new JButton("");
		btnStopMigration.setIcon(new ImageIcon(Main_Frame.class.getResource("/can-migration-btn.png")));
		btnStopMigration.setRolloverEnabled(false);
		btnStopMigration.setRequestFocusEnabled(false);
		btnStopMigration.setOpaque(false);
		btnStopMigration.setFocusable(false);
		btnStopMigration.setFocusTraversalKeysEnabled(false);
		btnStopMigration.setFocusPainted(false);
		btnStopMigration.setDefaultCapable(false);
		btnStopMigration.setContentAreaFilled(false);
		btnStopMigration.setBorderPainted(false);
		btnStopMigration.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnStopMigration.setIcon(new ImageIcon(Main_Frame.class.getResource("/can-migration-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnStopMigration.setIcon(new ImageIcon(Main_Frame.class.getResource("/can-migration-btn.png")));
			}
		});

		btnStopMigration.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				String url = "";
				String path = "";

				try {
					if (System.getProperty("os.name").toLowerCase().contains("windows")) {

						url = "jdbc:sqlite:" + System.getenv("APPDATA") + File.separator + projectTitle + File.separator
								+ hashcode("userdetails");

						path = System.getenv("APPDATA") + File.separator + projectTitle + File.separator
								+ hashcode("userdetails");

					} else {

						url = "jdbc:sqlite:" + System.getProperty("user.home") + File.separator + "Library"
								+ File.separator + "Application Support" + File.separator + projectTitle
								+ File.separator + hashcode("userdetails");

						path = System.getProperty("user.home") + File.separator + "Library" + File.separator
								+ "Application Support" + File.separator + projectTitle + File.separator
								+ hashcode("userdetails");

					}

					System.out.println("Database deleted successfully...");

					schsqlconnection = DriverManager.getConnection(url);

					System.out.println("Opened database successfully");
					schsqlstmt = (Statement) schsqlconnection.createStatement();
					String sql = "DROP TABLE Schedule_Detail";
					schsqlstmt.executeUpdate(sql);

				} catch (Exception e) {

					e.printStackTrace();
				} finally {

					try {
						if (schsqlstmt != null) {
							schsqlstmt.close();
							schsqlconnection.close();
						}
						if (schsqlconnection != null) {
							schsqlconnection.close();
						}

					} catch (Exception e) {
						System.out.println(e.getMessage());
					}

					File file = new File(path);
					recursiveDelete(file);

					System.exit(0);
				}

			}
		});
		btnStopMigration.setBounds(454, 516, 242, 38);
		panel_4.add(btnStopMigration);

		btnDowloadReport = new JButton("");
		btnDowloadReport.setRolloverEnabled(false);
		btnDowloadReport.setRequestFocusEnabled(false);
		btnDowloadReport.setOpaque(false);
		btnDowloadReport.setFocusable(false);
		btnDowloadReport.setFocusTraversalKeysEnabled(false);
		btnDowloadReport.setFocusPainted(false);
		btnDowloadReport.setDefaultCapable(false);
		btnDowloadReport.setContentAreaFilled(false);
		btnDowloadReport.setBorderPainted(false);
		btnDowloadReport.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnDowloadReport.setIcon(new ImageIcon(Main_Frame.class.getResource("/download-report-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnDowloadReport.setIcon(new ImageIcon(Main_Frame.class.getResource("/download-report-btn.png")));
			}
		});

		btnDowloadReport.setIcon(new ImageIcon(Main_Frame.class.getResource("/download-report-btn.png")));
		btnDowloadReport.setToolTipText("Click here to Download the status report. ");
		btnDowloadReport.addActionListener(new ActionListener() {

			public void actionPerformed(ActionEvent arg0) {
				cal = Calendar.getInstance();
				calendertime = getRidOfIllegalFileNameCharacters(cal.getTime().toString());
				reportpath = textField_hi.getText();
				new File(reportpath + File.separator + messageboxtitle + " report").mkdirs();

				File file = new File(reportpath + File.separator + messageboxtitle + " report" + File.separator
						+ calendertime + "report.csv");

				try {
					int column_no = table_fileConvertionreport_panel4.getColumnCount();
					FileWriter outputfile = new FileWriter(file);

					CSVWriter writer = new CSVWriter(outputfile);
					if (column_no == 6) {
						String[] header = { "From", "To", "Status", "Duration", "Message Count", "Path" };

						writer.writeNext(header);

					} else {
						String[] header = { "From", "To", "Status", "Duration", "Message Count", "Path", "Last Runtime",
								"Next RunTime" };

						writer.writeNext(header);
					}

					for (int i = 0; i < table_fileConvertionreport_panel4.getRowCount(); i++) {
						String g1 = "";
						try {
							g1 = table_fileConvertionreport_panel4.getValueAt(i, 0).toString();
						} catch (Exception e) {

						}
						String g2 = "";
						try {
							g2 = table_fileConvertionreport_panel4.getValueAt(i, 1).toString();
						} catch (Exception e) {

						}
						String g3 = "";
						try {
							g3 = table_fileConvertionreport_panel4.getValueAt(i, 2).toString();
						} catch (Exception e) {

						}
						String g4 = "";
						try {
							g4 = table_fileConvertionreport_panel4.getValueAt(i, 3).toString();
						} catch (Exception e) {

						}
						String g5 = "";
						try {
							g5 = table_fileConvertionreport_panel4.getValueAt(i, 4).toString();
						} catch (Exception e) {

						}
						String g6 = "";
						try {
							g6 = table_fileConvertionreport_panel4.getValueAt(i, 5).toString();
						} catch (Exception e) {

						}

						if (column_no == 6) {
							String[] data1 = { g1, g2, g3, g4, g5, g6 };
							writer.writeNext(data1);
						}

						else {
							String g7 = "";
							try {
								g7 = table_fileConvertionreport_panel4.getValueAt(i, 6).toString();
							} catch (Exception e) {

							}
							String g8 = "";
							try {
								g8 = table_fileConvertionreport_panel4.getValueAt(i, 7).toString();
							} catch (Exception e) {

							}
							String[] data1 = { g1, g2, g3, g4, g5, g6, g7, g8 };
							writer.writeNext(data1);
						}
					}

					writer.close();
					file.setReadOnly();
					Desktop desktop = Desktop.getDesktop();
					desktop.open(file);

				} catch (Error e) {
					logger.warning("Error : " + e.getMessage() + System.lineSeparator());
				} catch (Exception e) {
					logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
				}
			}
		});
		btnDowloadReport.setFont(new Font("Tahoma", Font.BOLD, 15));
		btnDowloadReport.setBounds(484, 421, 163, 39);
		panel_4.add(btnDowloadReport);

		btnConvertAgain = new JButton("");
		btnConvertAgain.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseEntered(MouseEvent arg0) {
				btnConvertAgain.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-again-hvr-btn.png")));
			}

			@Override
			public void mouseExited(MouseEvent e) {
				btnConvertAgain.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-again-btn.png")));
			}
		});
		btnConvertAgain.setIcon(new ImageIcon(Main_Frame.class.getResource("/convert-again-btn.png")));
		btnConvertAgain.setBorderPainted(false);
		btnConvertAgain.setContentAreaFilled(false);
		btnConvertAgain.setDefaultCapable(false);
		btnConvertAgain.setFocusTraversalKeysEnabled(false);
		btnConvertAgain.setFocusable(false);
		btnConvertAgain.setOpaque(false);
		btnConvertAgain.setRolloverEnabled(false);
		btnConvertAgain.setRequestFocusEnabled(false);
		btnConvertAgain.setToolTipText("Click here to Convert again.  ");
		btnConvertAgain.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				filetype = "";
				path = "";
				checkconvertagain = true;
				stop = false;
				lbl_progressreport.setText("");
				btn_Destination.setEnabled(true);
				btn_previous_p3.setEnabled(true);
				btnStop.setVisible(false);
				comboBox.setEnabled(false);
				label_12.setVisible(false);
				chckbx_calender_box.setEnabled(true);
				chckbx_convert_pdf_to_pdf.setEnabled(true);
				listduplicacy.clear();
				chckbxSaveInSame.setEnabled(true);
				textField_customfolder.setEditable(true);
				btn_signout_p3.setVisible(false);
				chckbxSavePdfAttachment.setEnabled(true);
				chckbxMigrateOrBackup.setEnabled(true);
				// chckbxMigrateOrBackup.setEnabled(true);
				chckbxCustomFolderName.setEnabled(true);
				chckbxCustomFolderName.setSelected(false);
				dateChooser_calender_start.setEnabled(true);
				spinner_sizespinner.setEnabled(true);
				comboBox_setsize.setEnabled(true);
				chckbx_splitpst.setEnabled(true);
				task_box.setEnabled(true);
				chckbxMaintainFolderHeirachy.setEnabled(true);
				chckbxRestoreToDefault.setEnabled(true);
				chckbxDeleteEmailFrom.setEnabled(true);
				chckbxRemoveDuplicacy.setEnabled(true);
				chckbxSetBackupSchedule.setEnabled(true);
				chckbxAutoIncrementBackup.setEnabled(true);
				chckbx_Mail_Filter.setEnabled(true);
				chckbx_Mail_Filter.setSelected(false);
				task_box.setEnabled(true);
				task_box.setSelected(false);
				chckbxSetBackupSchedule.setEnabled(true);
				label_14.setEnabled(true);
				dateChooserNextSchedular.setEnabled(true);
				spinner.setEnabled(true);
				rdbtnOnce.setEnabled(true);
				rdbtnEveryWeek.setEnabled(true);
				rdbtnEveryday.setEnabled(true);
				rdbtnOnWeekDay.setEnabled(true);
				rdbtnOnmonthDay.setEnabled(true);
				rdbtnEveryMonth.setEnabled(true);

				comboBox_fileDestination_type.setEnabled(true);

				CardLayout card = (CardLayout) Cardlayout.getLayout();
				card.show(Cardlayout, "panel_3");

			}
		});
		btnConvertAgain.setFont(new Font("Tahoma", Font.BOLD, 14));
		btnConvertAgain.setBounds(484, 559, 163, 39);
		panel_4.add(btnConvertAgain);

		progressBar_message_p3 = new JProgressBar();
		progressBar_message_p3.setBounds(0, 0, 11, 0);
		panel_4.add(progressBar_message_p3);
		progressBar_message_p3.setBackground(Color.WHITE);

		label_10 = new JLabel("");
		label_10.setIcon(new ImageIcon(Main_Frame.class.getResource("/bottom.png")));
		label_10.setBounds(0, 556, 1075, 64);
		panel_4.add(label_10);
		btnActivate = new JButton("");
		if (demo) {
			btnActivate.setVisible(true);
		} else {
			btnActivate.setVisible(true);

		}
		if (demo) {
			btnActivate.setToolTipText("Click here to activate the software.");
			btnActivate.addMouseListener(new MouseAdapter() {
				@Override
				public void mouseEntered(MouseEvent arg0) {
					btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/key-act-hvr-btn.png")));
				}

				@Override
				public void mouseExited(MouseEvent e) {
					btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/key-act-btn.png")));
				}
			});

			btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/key-act-btn.png")));
		} else {
			btnActivate.setToolTipText("Click here to deactivate the software.");
			btnActivate.addMouseListener(new MouseAdapter() {
				public void mouseEntered(MouseEvent arg0) {
					btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/deactivate-hvr-btn.png")));
				}

				public void mouseExited(MouseEvent e) {
					btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/deactivate-btn.png")));
				}
			});
			btnActivate.setIcon(new ImageIcon(Main_Frame.class.getResource("/deactivate-btn.png")));

		}

		menuBar = new JMenuBar();
		menuBar.setBackground(Color.WHITE);

		JMenu menu = new JMenu("Menu");
		JMenuItem ActivateTool = new JMenuItem("ActivateTool",
				new ImageIcon(Main_Frame.class.getResource("/activate.png")));
		ActivateTool.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {
				btnActivate.doClick();
			}
		});

		JMenuItem Info = new JMenuItem("Info", new ImageIcon(Main_Frame.class.getResource("/info.png")));
		Info.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {
				btn_info.doClick();
			}
		});

		JMenuItem exitItem = new JMenuItem("Exit", new ImageIcon(Main_Frame.class.getResource("/exist.png")));
		exitItem.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				String warn = "Do you want to Close the application?";
				int ans = JOptionPane.showConfirmDialog(Main_Frame.this, warn, messageboxtitle,
						JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {

					System.exit(0);
				}

			}
		});
		JMenuItem uninstall = new JMenuItem("Deactivate", new ImageIcon(Main_Frame.class.getResource("/unins.png")));
		uninstall.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				String warn = "Do you want to Deactivate the application?";
				int ans = JOptionPane.showConfirmDialog(Main_Frame.this, warn, messageboxtitle,
						JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
				if (ans == JOptionPane.YES_OPTION) {

					Starting_Frame sf = new Starting_Frame();
					new Uninstall(Starting_Frame.ToolUri, Starting_Frame.messageboxtitle, Starting_Frame.activationKey,
							Starting_Frame.orderId);
					dispose();
					sf.setLocationRelativeTo(null);
					sf.setResizable(false);
					sf.setVisible(true);
				}

			}
		});

		if (demo) {
			menu.add(ActivateTool);
			menu.add(Info);
			menu.add(exitItem);
		} else {
			menu.add(Info);
			menu.add(uninstall);
			menu.add(exitItem);
		}

		JMenu buymenu = new JMenu("Buy");
		tools = new JMenu("Tools");
		new JMenu("About");
		JMenu helpmenu = new JMenu("Help");
		
		JMenuItem buy = new JMenuItem("Buy", new ImageIcon(Main_Frame.class.getResource("/buy.png")));
		buy.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {
				btn_buy.doClick();
			}
		});
		if (demo) {
			btn_buy.setVisible(true);
		} else {
			btn_buy.setVisible(false);
		}
		buymenu.add(buy);

		JMenuItem About = new JMenuItem("Software Guide", new ImageIcon(Main_Frame.class.getResource("/aboutM.png")));
		About.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				btn_help.doClick();
			}
		});

		helpmenu.add(About);

		JMenuItem website = new JMenuItem("Website", new ImageIcon(Main_Frame.class.getResource("/website.png")));
		website.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {
				openBrowser(All_Data.websiteurl);
			}
		});

		helpmenu.add(website);
		if (!projectTitle.contains("Email Migration")) {
			JMenuItem emailmigartion = new JMenuItem("Email Migration (All in One)",
					new ImageIcon(Main_Frame.class.getResource("/website.png")));
			tools.add(emailmigartion);
			emailmigartion.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent ev) {

					SwingUtilities.invokeLater(new Runnable() {

						public void run() {

							openBrowser(All_Data.mailMigration);

						}
					});

				}
			});
		}

		JMenu emailclient = new JMenu("Email-Backup");
		emailclient.setIcon(new ImageIcon(Main_Frame.class.getResource("/email-clinet.png")));
		tools.add(emailclient);

		JMenuItem Gmail = new JMenuItem("Gmail Backup ", new ImageIcon(Main_Frame.class.getResource("/gmail-b.png")));
		Gmail.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {

						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.gmail);
						} else {

							All_Data.input_default = "Gmail";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
						}

					}
				});

			}
		});

		JMenuItem GmailtoPDF = new JMenuItem("Gmail To Pdf Backup ",
				new ImageIcon(Main_Frame.class.getResource("/gmail-b.png")));
		GmailtoPDF.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {

						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.gmail);
						} else {

							All_Data.input_default = "Gmail";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
							label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/gmail-to-pdf.png")));
						}

					}
				});

			}
		});

		JMenuItem yahoo = new JMenuItem("Yahoo Backup ", new ImageIcon(Main_Frame.class.getResource("/yahoo-b.png")));
		yahoo.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {

						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.yahoo);
						} else {

							All_Data.input_default = "Yahoo Mail";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
						}

					}
				});

			}
		});

		JMenuItem yahoopdf = new JMenuItem("Yahoo To Pdf Backup ",
				new ImageIcon(Main_Frame.class.getResource("/yahoo-b.png")));
		yahoopdf.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {

						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.yahoo);
						} else {
							All_Data.input_default = "Yahoo Mail";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
							label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/yahoo-to-pdf.png")));
						}

					}
				});

			}
		});

		JMenuItem hotmail = new JMenuItem("HotMail Backup ",
				new ImageIcon(Main_Frame.class.getResource("/hotmail-b.png")));
		hotmail.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {

						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.hotmail);
						} else {
							All_Data.input_default = "Hotmail";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
							label_8.setIcon(new ImageIcon(Main_Frame.class.getResource("/hotmail.png")));
						}
					}
				});

			}
		});

		JMenuItem Aol = new JMenuItem("Aol Backup ", new ImageIcon(Main_Frame.class.getResource("/aol-b.png")));
		Aol.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {

						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.aol);
						} else {

							All_Data.input_default = "AOL";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
						}
					}
				});

			}
		});

		JMenuItem office = new JMenuItem("Office365 Backup and Restore ",
				new ImageIcon(Main_Frame.class.getResource("/office365-b.png")));
		office.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {

						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.officeBkprestore);
						} else {

							All_Data.input_default = "Office 365 Backup & Restore";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
						}

					}
				});

			}
		});

		JMenuItem zoho = new JMenuItem("Zoho Backup ", new ImageIcon(Main_Frame.class.getResource("/zoho-b.png")));
		zoho.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {

						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.zoho);
						} else {

							All_Data.input_default = "Zoho Mail";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
						}

					}
				});

			}
		});

		JMenuItem icloud = new JMenuItem("Icloud Backup ",
				new ImageIcon(Main_Frame.class.getResource("/icloud-b.png")));
		icloud.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {

						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.icloud);
						} else {

							All_Data.input_default = "Icloud";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
						}
					}
				});

			}
		});

		JMenuItem hostgator = new JMenuItem("HostGator Backup ",
				new ImageIcon(Main_Frame.class.getResource("/hostgator-b.png")));
		hostgator.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {

						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.hostgator);
						} else {

							All_Data.input_default = "Hostgator email";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
						}
					}
				});

			}
		});

		JMenuItem imap = new JMenuItem("Imap Backup ", new ImageIcon(Main_Frame.class.getResource("/imap-b.png")));
		imap.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {
						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.imap);
						} else {

							All_Data.input_default = "IMAP";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
						}
					}
				});

			}
		});

		JMenuItem godady = new JMenuItem("GoDaddy Backup ",
				new ImageIcon(Main_Frame.class.getResource("/godady-b.png")));
		godady.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {
						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.godaddy);
						} else {

							All_Data.input_default = "GoDaddy email";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
						}
					}
				});

			}
		});

		JMenuItem Yandex = new JMenuItem("Yandex Backup ",
				new ImageIcon(Main_Frame.class.getResource("/yandex-b.png")));
		Yandex.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {

						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.yandex);
						} else {

							All_Data.input_default = "Yandex Mail";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
						}
					}
				});

			}
		});

		JMenuItem ThunderBird = new JMenuItem("Thunderbird Backup ",
				new ImageIcon(Main_Frame.class.getResource("/thunderbird-b.png")));
		ThunderBird.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {

						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.thunderbird);
						} else {

							All_Data.input_default = "Thunderbird";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
						}
					}
				});

			}
		});

		JMenuItem amazon = new JMenuItem("AmazonWebmail Backup ",
				new ImageIcon(Main_Frame.class.getResource("/amazon-b.png")));
		amazon.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {
						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.amazon);
						} else {

							All_Data.input_default = "Amazon WorkMail";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
						}
					}
				});

			}
		});

		JMenuItem opera = new JMenuItem("OperaMail Backup ",
				new ImageIcon(Main_Frame.class.getResource("/opera-b.png")));
		opera.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {
						if (!projectTitle.contains("Email Migration")) {
							openBrowser(All_Data.opera);
						} else {

							All_Data.input_default = "Opera Mail";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
						}
					}
				});

			}
		});

		emailclient.add(Gmail);
		emailclient.add(GmailtoPDF);
		emailclient.add(yahoo);
		emailclient.add(yahoopdf);
		emailclient.add(Aol);
		emailclient.add(hotmail);
		emailclient.add(office);
		emailclient.add(zoho);
		emailclient.add(icloud);
		emailclient.add(hostgator);
		emailclient.add(imap);
		emailclient.add(godady);
		emailclient.add(Yandex);
		emailclient.add(ThunderBird);
		emailclient.add(amazon);
		emailclient.add(opera);

		JMenu FileFormat = new JMenu("Email-Converter");
		FileFormat.setIcon(new ImageIcon(Main_Frame.class.getResource("/file-format.png")));
		tools.add(FileFormat);

		if (projectTitle.contains("Mac")) {

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("PST Converter", new ImageIcon(Main_Frame.class.getResource("/pst-b.png"))),
					"MICROSOFT OUTLOOK (.pst)", null));

			FileFormat.add(MeniItemFileFormat(new JMenuItem("NSF to PST Converter (Upgrade)",
					new ImageIcon(Main_Frame.class.getResource("/nsfb.png"))), null, All_Data.nsftopstflink));

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("PST to OFFICE 365", new ImageIcon(Main_Frame.class.getResource("/pst-b.png"))), null,
					All_Data.pst2office));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("OST Converter", new ImageIcon(Main_Frame.class.getResource("/ost-b.png"))), null,
					All_Data.ost));

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("MBOX Converter", new ImageIcon(Main_Frame.class.getResource("/mbox-b.png"))), null,
					All_Data.mbox));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("MBOX to OFFICE 365", new ImageIcon(Main_Frame.class.getResource("/mbox-b.png"))),
					null, All_Data.mbox2office));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("MSG Converter", new ImageIcon(Main_Frame.class.getResource("/msg-b.png"))), null,
					All_Data.msg));

			JMenuItem msgtopdf = new JMenuItem("MSG to PDF Converter",
					new ImageIcon(Main_Frame.class.getResource("/msg-b.png")));
			msgtopdf.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent ev) {

					SwingUtilities.invokeLater(new Runnable() {

						public void run() {
							openBrowser(All_Data.msg);

						}
					});

				}
			});

			FileFormat.add(msgtopdf);

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("EML Converter", new ImageIcon(Main_Frame.class.getResource("/eml-b.png"))), null,
					All_Data.eml));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("EMLX Converter", new ImageIcon(Main_Frame.class.getResource("/emlx-b.png"))), null,
					All_Data.emlx));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("OLM Converter", new ImageIcon(Main_Frame.class.getResource("/olm-b.png"))), null,
					All_Data.olm));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("MAILDIR Converter", new ImageIcon(Main_Frame.class.getResource("/maildir-b.png"))),
					null, All_Data.mailldir));

		} else if (!projectTitle.contains("Email Migration")) {

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("PST Converter", new ImageIcon(Main_Frame.class.getResource("/pst-b.png"))),
					"MICROSOFT OUTLOOK (.pst)", null));

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("PST Recovery (Upgrade)", new ImageIcon(Main_Frame.class.getResource("/pst-b.png"))),
					null, All_Data.pstreccoverylink));

			FileFormat.add(MeniItemFileFormat(new JMenuItem("NSF to PST Converter (Upgrade)",
					new ImageIcon(Main_Frame.class.getResource("/nsfb.png"))), null, All_Data.nsftopstflink));

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("EDB To Pst Recovery (Upgrade) ",
							new ImageIcon(Main_Frame.class.getResource("/edbb.png"))),
					null, All_Data.pstreccoverylink));

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("PST to OFFICE 365", new ImageIcon(Main_Frame.class.getResource("/pst-b.png"))), null,
					All_Data.pst2office));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("OST Converter", new ImageIcon(Main_Frame.class.getResource("/ost-b.png"))), null,
					All_Data.ost));

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("OST Recovery  (Upgrade)", new ImageIcon(Main_Frame.class.getResource("/ost-b.png"))),
					null, All_Data.ostreccoverylink));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("MBOX Converter", new ImageIcon(Main_Frame.class.getResource("/mbox-b.png"))), null,
					All_Data.mbox));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("MBOX to OFFICE 365", new ImageIcon(Main_Frame.class.getResource("/mbox-b.png"))),
					null, All_Data.mbox2office));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("MSG Converter", new ImageIcon(Main_Frame.class.getResource("/msg-b.png"))), null,
					All_Data.msg));

			JMenuItem msgtopdf = new JMenuItem("MSG to PDF Converter",
					new ImageIcon(Main_Frame.class.getResource("/msg-b.png")));
			msgtopdf.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent ev) {

					SwingUtilities.invokeLater(new Runnable() {

						public void run() {
							openBrowser(All_Data.msg);

						}
					});

				}
			});

			FileFormat.add(msgtopdf);

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("EML Converter", new ImageIcon(Main_Frame.class.getResource("/eml-b.png"))), null,
					All_Data.eml));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("EMLX Converter", new ImageIcon(Main_Frame.class.getResource("/emlx-b.png"))), null,
					All_Data.emlx));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("OLM Converter", new ImageIcon(Main_Frame.class.getResource("/olm-b.png"))), null,
					All_Data.olm));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("MAILDIR Converter", new ImageIcon(Main_Frame.class.getResource("/maildir-b.png"))),
					null, All_Data.mailldir));

		} else {

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("PST Converter", new ImageIcon(Main_Frame.class.getResource("/pst-b.png"))),
					"MICROSOFT OUTLOOK (.pst)", null));

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("PST Recovery (Upgrade)", new ImageIcon(Main_Frame.class.getResource("/pst-b.png"))),
					null, All_Data.pstreccoverylink));

			FileFormat.add(MeniItemFileFormat(new JMenuItem("NSF to PST Converter (Upgrade)",
					new ImageIcon(Main_Frame.class.getResource("/nsfb.png"))), null, All_Data.nsftopstflink));

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("EDB To Pst Recovery (Upgrade) ",
							new ImageIcon(Main_Frame.class.getResource("/edbb.png"))),
					null, All_Data.pstreccoverylink));

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("PST to OFFICE 365", new ImageIcon(Main_Frame.class.getResource("/pst-b.png"))),
					"PST to Office 365", null));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("OST Converter", new ImageIcon(Main_Frame.class.getResource("/ost-b.png"))),
					"Exchange Offline Storage (.ost)", null));

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("OST Recovery  (Upgrade)", new ImageIcon(Main_Frame.class.getResource("/ost-b.png"))),
					null, All_Data.ostreccoverylink));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("MBOX Converter", new ImageIcon(Main_Frame.class.getResource("/mbox-b.png"))), "MBOX",
					null));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("MBOX to OFFICE 365", new ImageIcon(Main_Frame.class.getResource("/mbox-b.png"))),
					"MBOX to Office 365", null));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("MSG Converter", new ImageIcon(Main_Frame.class.getResource("/msg-b.png"))),
					"Message File (.msg)", null));

			JMenuItem msgtopdf = new JMenuItem("MSG to PDF Converter",
					new ImageIcon(Main_Frame.class.getResource("/amazon-b.png")));
			msgtopdf.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent ev) {

					SwingUtilities.invokeLater(new Runnable() {

						public void run() {
							All_Data.input_default = "Message File (.msg)";
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
							label_7.setIcon(new ImageIcon(Main_Frame.class.getResource("/msg-to-pdf.png")));

						}
					});

				}
			});

			FileFormat.add(msgtopdf);

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("MSG to PDF Converter", new ImageIcon(Main_Frame.class.getResource("/msg-b.png"))),
					"Message File (.msg)", null));

			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("EML Converter", new ImageIcon(Main_Frame.class.getResource("/eml-b.png"))),
					"EML File (.eml)", null));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("EMLX Converter", new ImageIcon(Main_Frame.class.getResource("/emlx-b.png"))),
					"EMLX File (.emlx)", null));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("OLM Converter", new ImageIcon(Main_Frame.class.getResource("/olm-b.png"))),
					"OLM File (.olm)", null));
			FileFormat.add(MeniItemFileFormat(
					new JMenuItem("MAILDIR Converter", new ImageIcon(Main_Frame.class.getResource("/maildir-b.png"))),
					"Maildir", null));

		}

		menuBar.add(menu);

		if (demo) {
			menuBar.add(buymenu);
		}
		menuBar.add(tools);
		menuBar.add(helpmenu);
		setJMenuBar(menuBar);
		
		JPanel panel_10 = new JPanel();
		panel_10.setBackground(new Color(0,	0 , 0));
		
				lblNewLabel_5 = new JLabel("");
				lblNewLabel_5.setRequestFocusEnabled(false);
				lblNewLabel_5.setInheritsPopupMenu(false);
				lblNewLabel_5.setFocusable(false);
				lblNewLabel_5.setFocusTraversalKeysEnabled(false);
				lblNewLabel_5.setIcon(new ImageIcon(Main_Frame.class.getResource("/topbar.png")));
						btn_buy.setToolTipText("Click here to purchase the software.");
						btn_buy.addActionListener(new ActionListener() {
							public void actionPerformed(ActionEvent arg0) {
								openBrowser(buyurl);

							}
						});
						
								btn_buy.addMouseListener(new MouseAdapter() {
						
									public void mouseEntered(MouseEvent arg0) {
										btn_buy.setIcon(new ImageIcon(Main_Frame.class.getResource("/buy-hvr-btn.png")));
									}
						
									public void mouseExited(MouseEvent e) {
										btn_buy.setIcon(new ImageIcon(Main_Frame.class.getResource("/buy-btn.png")));
									}
								});
										
												btnNewButton = new JButton("");
												btnNewButton.setToolTipText("Click here for Technical Support");
												btnNewButton.setRolloverEnabled(false);
												btnNewButton.setRequestFocusEnabled(false);
												btnNewButton.setOpaque(false);
												btnNewButton.setFocusable(false);
												btnNewButton.setFocusTraversalKeysEnabled(false);
												btnNewButton.setFocusPainted(false);
												btnNewButton.setDefaultCapable(false);
												btnNewButton.setBorderPainted(false);
												btnNewButton.setContentAreaFilled(false);
												
														btnNewButton.addActionListener(new ActionListener() {
															public void actionPerformed(ActionEvent arg0) {
																openBrowser(All_Data.helpuri);
															}
														});
														btnNewButton.addMouseListener(new MouseAdapter() {

															public void mouseEntered(MouseEvent arg0) {
																btnNewButton.setIcon(new ImageIcon(Main_Frame.class.getResource("/live-chat-hvr-btn1.png")));

															}

															public void mouseExited(MouseEvent e) {
																btnNewButton.setIcon(new ImageIcon(Main_Frame.class.getResource("/live-chat-btn1.png")));
															}
														});
														
//		JButton btnNewButton_2 = new JButton("New button");
//		btnNewButton_2.setBounds(693, 6, 89, 55);
//		contentPane.add(btnNewButton_2);
/////////////////
														updateBtn = new JButton("");
														updateBtn.setFont(new Font("Tahoma", Font.BOLD, 11));
														updateBtn.addMouseListener(new MouseAdapter() {
															public void mouseEntered(MouseEvent arg0) {
																updateBtn.setIcon(new ImageIcon(Main_Frame.class.getResource("/update-hvr-btn.png")));
															}

															public void mouseExited(MouseEvent e) {
																updateBtn.setIcon(new ImageIcon(Main_Frame.class.getResource("/update-btn.png")));
															}
														});
														updateBtn.setIcon(new ImageIcon(Main_Frame.class.getResource("/update-btn.png")));
														updateBtn.addActionListener(new ActionListener() {
															public void actionPerformed(ActionEvent e) {
																try {
																	Desktop.getDesktop().browse(new URL(All_Data.updateSoftware).toURI());
																} catch (IOException | URISyntaxException e1) {

																	e1.printStackTrace();
																	StringWriter sw = new StringWriter();
																	PrintWriter pw = new PrintWriter(sw);
																	e1.printStackTrace(pw);
																	logger.warning(sw.toString());
																}
															}
														});
														updateBtn.setFocusTraversalKeysEnabled(false);
														updateBtn.setFocusable(false);
														updateBtn.setOpaque(false);
														updateBtn.setRolloverEnabled(false);
														updateBtn.setRequestFocusEnabled(false);
														updateBtn.setFocusPainted(false);
														updateBtn.setDefaultCapable(false);
														updateBtn.setContentAreaFilled(false);
														updateBtn.setVisible(false);
														updateBtn.setBorderPainted(false);
														updateBtn.setVisible(false);
														
															
																btnActivate.setRolloverEnabled(false);
																btnActivate.setRequestFocusEnabled(false);
																btnActivate.setOpaque(false);
																btnActivate.setFocusable(false);
																btnActivate.setFocusTraversalKeysEnabled(false);
																btnActivate.setFocusPainted(false);
																btnActivate.setDefaultCapable(false);
																btnActivate.setContentAreaFilled(false);
																btnActivate.setBorderPainted(false);
																//		btnActivate.setToolTipText("Click here to activate the software.");
																		btnActivate.addActionListener(new ActionListener() {
																			public void actionPerformed(ActionEvent e) {
																				try {
																					if (demo) {
																						if (System.getProperty("os.name").toLowerCase().contains("windows")) {
																							licFileon = new File(System.getenv("APPDATA") + File.separator + projectTitle
																									+ File.separator + "licenseOnline.lic");
																							licFileonoff = new File(System.getenv("APPDATA") + File.separator + projectTitle
																									+ File.separator + "license.lic");
																						} else {
																							licFileon = new File(System.getProperty("user.home") + File.separator + "Library"
																									+ File.separator + "Application Support" + File.separator + projectTitle
																									+ File.separator + "licenseOnline.lic");
																							licFileonoff = new File(System.getenv("APPDATA") + File.separator + projectTitle
																									+ File.separator + "license.lic");
																						}
																						boolean activatefromdemo = true;
																						new Starting_Frame();
																						OnlineActivation mf = new OnlineActivation(Starting_Frame.mf, licFileon, activatefromdemo);
																						mf.setLocationRelativeTo(null);
																						mf.setVisible(true);
																						setEnabled(false);
																//						setVisible(false);
																						mf.btnBack.setVisible(false);
																						mf.addWindowListener(new WindowAdapter() {
																							@Override
																							public void windowClosing(WindowEvent arg0) {
																								String warn = "Do you want to close?";
																								int ans = JOptionPane.showConfirmDialog(mf, warn, All_Data.messageboxtitle,
																										JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE,
																										new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
																								if (ans == JOptionPane.YES_OPTION) {
																									setEnabled(true);
																//									setVisible(true);
																									mf.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
																								} else {
																									mf.setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);
																								}
																							}
																						});
																						All_Data.messageboxtitle = projectTitle;
																					} else {
																						String warn = "Do you want to Deactivate the Software?";
																						int ans = JOptionPane.showConfirmDialog(Main_Frame.this, warn, All_Data.messageboxtitle,
																								JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE,
																								new ImageIcon(Main_Frame.class.getResource("/about-icon-2.png")));
																						if (ans == JOptionPane.YES_OPTION) {
																							Starting_Frame sf = new Starting_Frame();
																							new Uninstall(Starting_Frame.ToolUri, All_Data.messageboxtitle,
																									Starting_Frame.activationKey, Starting_Frame.orderId);
																							dispose();
																							sf.setLocationRelativeTo(null);
																							sf.setResizable(false);
																							sf.setVisible(true);
																						}
																					}
																				} catch (Exception e1) {
																					e1.printStackTrace();
																					logger.warning(e1.getMessage() + System.lineSeparator());
																				} catch (Error e1) {
																					e1.printStackTrace();
																					logger.warning(e1.getMessage() + System.lineSeparator());
																				}
																
																			}
																		});
														updateBtn.setToolTipText("Click here to download the latest version of the software.");
														btnNewButton.setIcon(new ImageIcon(Main_Frame.class.getResource("/live-chat-btn1.png")));
								
										btn_buy.setIcon(new ImageIcon(Main_Frame.class.getResource("/buy-btn.png")));
										
												btn_buy.setOpaque(false);
												btn_buy.setRolloverEnabled(false);
												btn_buy.setRequestFocusEnabled(false);
												btn_buy.setFocusTraversalKeysEnabled(false);
												btn_buy.setFocusable(false);
												btn_buy.setFocusPainted(false);
												btn_buy.setDefaultCapable(false);
												btn_buy.setContentAreaFilled(false);
												btn_buy.setBorderPainted(false);
				GroupLayout gl_panel_10 = new GroupLayout(panel_10);
				gl_panel_10.setHorizontalGroup(
					gl_panel_10.createParallelGroup(Alignment.LEADING)
						.addComponent(lblNewLabel_5, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
				);
				gl_panel_10.setVerticalGroup(
					gl_panel_10.createParallelGroup(Alignment.LEADING)
						.addComponent(lblNewLabel_5, GroupLayout.PREFERRED_SIZE, 70, GroupLayout.PREFERRED_SIZE)
				);
				panel_10.setLayout(gl_panel_10);
				GroupLayout gl_contentPane = new GroupLayout(contentPane);
				gl_contentPane.setHorizontalGroup(
					gl_contentPane.createParallelGroup(Alignment.LEADING)
						.addGroup(Alignment.TRAILING, gl_contentPane.createSequentialGroup()
							.addGap(643)
							.addComponent(updateBtn, GroupLayout.PREFERRED_SIZE, 73, GroupLayout.PREFERRED_SIZE)
							.addGap(1)
							.addComponent(btnNewButton, GroupLayout.PREFERRED_SIZE, 64, GroupLayout.PREFERRED_SIZE)
							.addGap(2)
							.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
								.addGroup(gl_contentPane.createSequentialGroup()
									.addGap(171)
									.addComponent(btn_help, GroupLayout.PREFERRED_SIZE, 50, GroupLayout.PREFERRED_SIZE))
								.addGroup(gl_contentPane.createSequentialGroup()
									.addGap(45)
									.addComponent(btnActivate, GroupLayout.PREFERRED_SIZE, 127, GroupLayout.PREFERRED_SIZE))
								.addComponent(btn_buy, GroupLayout.PREFERRED_SIZE, 50, GroupLayout.PREFERRED_SIZE)
								.addGroup(gl_contentPane.createSequentialGroup()
									.addGap(210)
									.addComponent(btn_info, GroupLayout.PREFERRED_SIZE, 64, GroupLayout.PREFERRED_SIZE)))
							.addGap(18))
						.addComponent(panel_10, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
						.addComponent(Cardlayout, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
				);
				gl_contentPane.setVerticalGroup(
					gl_contentPane.createParallelGroup(Alignment.LEADING)
						.addGroup(gl_contentPane.createSequentialGroup()
							.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
								.addGroup(gl_contentPane.createSequentialGroup()
									.addGap(6)
									.addComponent(updateBtn, GroupLayout.PREFERRED_SIZE, 55, GroupLayout.PREFERRED_SIZE))
								.addGroup(gl_contentPane.createSequentialGroup()
									.addGap(6)
									.addComponent(btnNewButton, GroupLayout.PREFERRED_SIZE, 55, GroupLayout.PREFERRED_SIZE))
								.addGroup(gl_contentPane.createSequentialGroup()
									.addGap(6)
									.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
										.addComponent(btn_help, GroupLayout.PREFERRED_SIZE, 61, GroupLayout.PREFERRED_SIZE)
										.addGroup(gl_contentPane.createSequentialGroup()
											.addGap(14)
											.addComponent(btnActivate, GroupLayout.PREFERRED_SIZE, 32, GroupLayout.PREFERRED_SIZE))
										.addComponent(btn_buy, GroupLayout.PREFERRED_SIZE, 61, GroupLayout.PREFERRED_SIZE)
										.addComponent(btn_info, GroupLayout.PREFERRED_SIZE, 61, GroupLayout.PREFERRED_SIZE)))
								.addComponent(panel_10, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
							.addGap(2)
							.addComponent(Cardlayout, GroupLayout.PREFERRED_SIZE, 595, Short.MAX_VALUE))
				);
				contentPane.setLayout(gl_contentPane);
		
		
		
	}
	
	
	
	public static void main(String[] args) {
		
		
		
		
		
		File folder = null;
		if (System.getProperty("os.name").toLowerCase().contains("windows")) {
			folder = new File(System.getenv("APPDATA") + File.separator + All_Data.messageboxtitle);
			folder.mkdirs();
		} else {
			folder = new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
					+ "Application Support" + File.separator + All_Data.messageboxtitle);
			folder.mkdirs();
		}

		String fileKey = null;
		if (System.getProperty("os.name").toLowerCase().contains("windows")) {
			licFileon = new File(System.getenv("APPDATA") + File.separator + All_Data.messageboxtitle+ File.separator
					+ "licenseOnline.lic");

			licFileonoff = new File(System.getenv("APPDATA") + File.separator + All_Data.messageboxtitle+ File.separator
					+ "license.lic");

		} else {
			licFileon = new File(System.getProperty("user.home") + File.separator + "Library" + File.separator
					+ "Application Support" + File.separator + All_Data.messageboxtitle + File.separator
					+ "licenseOnline.lic");
			licFileonoff = new File(System.getenv("APPDATA") + File.separator + All_Data.messageboxtitle + File.separator
					+ "license.lic");
		}
		
		
		if (licFileonoff.exists()) {
			try {
				FileReader fr = new FileReader(licFileonoff);
				BufferedReader br = new BufferedReader(fr);
				fileKey = br.readLine();
				fr.close();
			} catch (Exception ex) {
			}

			if (fileKey != null) {

				strSerialNumber = ActivationFrame
						.getSerialNumber(System.getProperty("user.home").substring(0, 1));
				hashKey = new Hash().getHash(strSerialNumber);
				String licencetype = fileKey.substring(fileKey.length() - 1);
				fileKey = fileKey.substring(0, fileKey.length() - 1);

				int intlic = Integer.valueOf(licencetype);

				System.out.println(fileKey);

				if (hashKey.equals(fileKey)) {
					try {
						UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
					} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
							| UnsupportedLookAndFeelException e) {
						e.printStackTrace();
					}
				 mf1 = new Main_Frame(false, intlic);
					mf1.setLocationRelativeTo(null);
//					mf.setVisible(true);
					mf1.temppath = mf1.textField_1.getText();
					mf1.cal = Calendar.getInstance();

					mf1.calendertime = mf1.getRidOfIllegalFileNameCharacters(mf1.cal.getTime().toString());
					main_multiplefile multi=new main_multiplefile( mf1,  mf1.demo,  mf1.messageboxtitle);
					multi.setLocationRelativeTo(null);
					multi.setVisible(true);
				} else {
					try {
						UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
					} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
							| UnsupportedLookAndFeelException e) {
						e.printStackTrace();
					}
					ActivationFrame af = new ActivationFrame();
					af.setLocationRelativeTo(null);
					af.setVisible(true);
				}
			} else {
				try {

					frame.setLocationRelativeTo(null);
					frame.setResizable(false);
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}

		} else {
			
		
		
		System.out.println("hi this is main methord ");
		Main_Frame frame = new Main_Frame(true, 4);
		frame.setLocationRelativeTo(null);
		frame.temppath = frame.textField_1.getText();
		frame.cal = Calendar.getInstance();
		frame.calendertime = frame.getRidOfIllegalFileNameCharacters(frame.cal.getTime().toString());
		
		
	 multi=new main_multiplefile( frame,  frame.demo,  frame.messageboxtitle);
		multi.setLocationRelativeTo(null);
		multi.setVisible(true);
	}
	
	}
	
	
	
	

	void destinationPath() throws Exception {
		jFileChooser = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");

		jFileChooser.setMultiSelectionEnabled(true);

		jFileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

		jFileChooser.showOpenDialog(Main_Frame.this);

		File file = jFileChooser.getSelectedFile();

		String destination = file.getAbsolutePath();

		tf_Destination_Location.setText(destination);

	}

	void filter_file() throws Exception {

		jFileChooser = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");

		jFileChooser.setBackground(Color.WHITE);

		jFileChooser.setAcceptAllFileFilterUsed(false);
		FileNameExtensionFilter filter = null;
		if (fileoption.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {

			filter = new FileNameExtensionFilter(".ost", "ost");

		} else if (fileoption.equalsIgnoreCase("DBX")) {

			filter = new FileNameExtensionFilter(".dbx", "dbx");

		} else if (fileoption.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {

			filter = new FileNameExtensionFilter(".pst", "pst");

		} else if (fileoption.equalsIgnoreCase("EML File (.eml)")) {
			filter = new FileNameExtensionFilter(".eml", "eml");

		} else if (fileoption.equalsIgnoreCase("EMLX File (.emlx)")) {
			filter = new FileNameExtensionFilter(".emlx", "emlx");

		} else if (fileoption.equalsIgnoreCase("OFT File (.oft)")) {
			filter = new FileNameExtensionFilter(".oft", "oft");

		} else if (fileoption.equalsIgnoreCase("Message File (.msg)")) {
			filter = new FileNameExtensionFilter(".msg", "msg");

		} else if (fileoption.equalsIgnoreCase("Zimbra files (.tgz)")) {
			filter = new FileNameExtensionFilter(".tgz", "tgz");

		} else if (fileoption.equalsIgnoreCase("Maildir")) {

			jFileChooser.setAcceptAllFileFilterUsed(true);
		} else if (fileoption.equalsIgnoreCase("MBOX")) {

			jFileChooser.setFileFilter(new FileNameExtensionFilter(".mbox", "mbx", "mbox"));
			jFileChooser.setAcceptAllFileFilterUsed(true);
		} else if (fileoption.equalsIgnoreCase("OLM File (.olm)")) {
			filter = new FileNameExtensionFilter(".olm", "olm");

		}
		jFileChooser.addChoosableFileFilter(filter);

		if (jFileChooser.showOpenDialog(Main_Frame.this) == JFileChooser.APPROVE_OPTION) {
			// jFileChooser.showOpenDialog(null);

			file = jFileChooser.getSelectedFile();

			if (!(file == null)) {
				if (!(filepath == null)) {
					if (filepath.equalsIgnoreCase(file.getAbsolutePath())) {
						JOptionPane.showMessageDialog(frame, "Same file added ", messageboxtitle,
								JOptionPane.INFORMATION_MESSAGE,
								new ImageIcon(Main_Frame.class.getResource("/information.png")));
					}
				}

				filepath = file.getAbsolutePath();

				if (fileoption.equalsIgnoreCase("MBOX")) {
					String fnam = file.getName();
					fname = fnam.replace(".mbx", "").replace(".mbox", "");

				} else if (fileoption.equalsIgnoreCase("DBX")) {
					String fnam = file.getName();
					fname = fnam.replace(".pst", "");

				} else if (fileoption.equalsIgnoreCase("MICROSOFT OUTLOOK (.pst)")) {
					String fnam = file.getName();
					fname = fnam.replace(".pst", "");

				} else if (fileoption.equalsIgnoreCase("Exchange Offline Storage (.ost)")) {
					String fnam = file.getName();
					fname = fnam.replace(".ost", "");
				} else if (fileoption.equalsIgnoreCase("EML File (.eml)")) {
					String fnam = file.getName();
					fname = fnam.replace(".eml", "");
				} else if (fileoption.equalsIgnoreCase("EMLX File (.emlx)")) {
					String fnam = file.getName();
					fname = fnam.replace(".emlx", "");
				} else if (fileoption.equalsIgnoreCase("OFT File (.oft)")) {
					String fnam = file.getName();
					fname = fnam.replace(".oft", "");
				} else if (fileoption.equalsIgnoreCase("Message File (.msg)")) {
					String fnam = file.getName();
					fname = fnam.replace(".msg", "");
				} else if (fileoption.equalsIgnoreCase("OLM File (.olm)")) {
					String fnam = file.getName();
					fname = fnam.replace(".olm", "");
				} else if (fileoption.equalsIgnoreCase("Maildir")) {
					String fnam = file.getName();
					fname = fnam;
				}
			}
		} else {
			comboBox_FiletypeChooser.setEnabled(true);
		}

	}

	public void readAnOST_PstFile() {

		try {
			pst = PersonalStorage.fromFile(filepath);
		} catch (Exception e) {

			if (e.getMessage().contains("File not found")) {
				JOptionPane.showMessageDialog(frame,
						"File is in use in other application please close that application", messageboxtitle,
						JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));

			}

			else {
				JOptionPane.showMessageDialog(frame, "File is Currupted  Please Choose another file  ", messageboxtitle,
						JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));

			}
		}

		model = (DefaultTreeModel) tree.getModel();

		File fileup = new File(filepath);

		String sname = fileup.getName();
		sname = filepath.replace(",", "");

		root = new DefaultMutableTreeNode("<html><b>" + sname);

		model.setRoot(root);

		folderInfo = pst.getRootFolder();
		String rootname = folderInfo.getDisplayName().replaceAll("[\\[\\]],", "");
		if (rootname.equalsIgnoreCase("")) {
			rootname = "Root Folder";
		}
		DefaultMutableTreeNode node1 = new DefaultMutableTreeNode("<html><b>" + rootname);
		root.add(node1);

		FolderInfoCollection folderInfoCollection = null;
		try {
			folderInfoCollection = pst.getRootFolder().getSubFolders();
		} catch (Exception e) {
			e.printStackTrace();
			JOptionPane.showMessageDialog(frame, "File is Currupted  Please Choose another file  ", messageboxtitle,
					JOptionPane.INFORMATION_MESSAGE, new ImageIcon(Main_Frame.class.getResource("/information.png")));

		}
		for (int i = 0; i < folderInfoCollection.size(); i++) {

			folderInfo = folderInfoCollection.get_Item(i);
			if (stop_tree) {
				break;
			}

			lblNewLabel_3.setText(folderInfo.getDisplayName());

			String Folder = folderInfo.getDisplayName();
			Folder = Folder.replace(",", "").replace(".", "");
			Folder = getRidOfIllegalFileNameCharacters(Folder);
			Folder = Folder.trim();
			DefaultMutableTreeNode node = new DefaultMutableTreeNode("<html><b>" + Folder);
			node1.add(node);

			if (folderInfo.hasSubFolders()) {

				readOstpstsubfolder(folderInfo, node);

			}

		}

	}

	public void readOstpstsubfolder(FolderInfo f, DefaultMutableTreeNode node) {

		FolderInfoCollection folderCollection = f.getSubFolders();

		for (int i = 0; i < folderCollection.size(); i++) {

			if (stop_tree) {
				break;
			}

			FolderInfo folderInfo = folderCollection.get_Item(i);
			String Folder = folderInfo.getDisplayName();
			Folder = Folder.replace(",", "").replace(".", "");
			Folder = getRidOfIllegalFileNameCharacters(Folder);
			Folder = Folder.trim();
			DefaultMutableTreeNode nod1 = new DefaultMutableTreeNode("<html><b>" + Folder);

			lblNewLabel_3.setText(folderInfo.getDisplayName());

			node.add(nod1);

			if (folderInfo.hasSubFolders()) {
				readOstpstsubfolder(folderInfo, nod1);
			}

		}
	}

	public void readdbxFile() {

		model = (DefaultTreeModel) tree.getModel();

		root = new DefaultMutableTreeNode("<html><b>" + hostName);

		model.setRoot(root);

		String sname = file.getName();
		sname = filepath.replace(",", "");

		DefaultMutableTreeNode node = new DefaultMutableTreeNode("<html><b>" + sname);

		root.add(node);

		DefaultMutableTreeNode child = new DefaultMutableTreeNode("<html><b>" + file.getName());

		node.add(child);
		lblNewLabel_3.setText(file.getName());

	}

	public void readMboxFile() {

		model = (DefaultTreeModel) tree.getModel();

		root = new DefaultMutableTreeNode("<html><b>" + hostName);

		model.setRoot(root);

		String filepath1 = filepath(file);

		String sname = file.getName();
		sname = filepath1.replace(",", "");

		CustomTreeNode node = new CustomTreeNode("<html><b>" + sname);

		node.filepath = file.getAbsolutePath();

		root.add(node);
		String ss = file.getName();
		CustomTreeNode child = new CustomTreeNode("<html><b>" + ss);

		child.filepath = file.getAbsolutePath();

		node.add(child);

		lblNewLabel_3.setText(ss);

	}

	public void readexchange() {
		model = (DefaultTreeModel) tree.getModel();

		root = new DefaultMutableTreeNode("<html><b>" + fileoption);

		model.setRoot(root);

		DefaultMutableTreeNode node = new DefaultMutableTreeNode("<html><b>" + textField_username_p2.getText());

		root.add(node);

		String rootUri = clientforexchange_input.getMailboxInfo().getRootUri();
		listExchangemesingos.clear();
		listFolderinfostring.clear();

		ExchangeFolderInfoCollection folderInfoCollection = clientforexchange_input.listSubFolders(rootUri);
		for (int j = 0; j < folderInfoCollection.size(); j++) {
			try {

				if (stop_tree) {
					break;
				}

				ExchangeFolderInfo folderInfo = folderInfoCollection.get_Item(j);

				String folder = folderInfo.getDisplayName().replaceAll("[\\[\\]]", "");

				if (!(folder.equalsIgnoreCase("PersonMetadata"))) {
					CustomTreeNode t = new CustomTreeNode("<html><b>" + folder);
					// t.filepath = folderInfo.getUri();
					node.add(t);

					listExchangemesingos.add(folderInfo);
					listFolderinfostring.add(folder);

					obTh.ob.MessageLabel.setText(folderInfo.getDisplayName());

					if (folderInfo.getChildFolderCount() > 0) {

						readexchange_subfolder(folderInfo, t, folder);

					}
				}
			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}

		}

	}

	public void readexchange_subfolder(ExchangeFolderInfo folderInfo, DefaultMutableTreeNode node, String s) {
		ExchangeFolderInfoCollection folderInfoCollection = clientforexchange_input.listSubFolders(folderInfo);
		for (int j = 0; j < folderInfoCollection.size(); j++) {
			try {
				if (stop_tree) {
					break;
				}

				ExchangeFolderInfo folderInfo1 = folderInfoCollection.get_Item(j);
				String folder = folderInfo1.getDisplayName().replaceAll("[\\[\\]]", "");
				DefaultMutableTreeNode t = new DefaultMutableTreeNode("<html><b>" + folder);

				node.add(t);
				obTh.ob.MessageLabel.setText(folder);

				s = s + File.separator + folder;
				listExchangemesingos.add(folderInfo1);
				listFolderinfostring.add(s);

				if (folderInfo1.getChildFolderCount() > 0) {

					readexchange_subfolder(folderInfo1, t, s);

				}
				s = removefolder(s);
			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}

		}

	}

	public void readimap() {

		model = (DefaultTreeModel) tree.getModel();

		root = new DefaultMutableTreeNode("<html><b>" + fileoption);

		model.setRoot(root);

		DefaultMutableTreeNode node = new DefaultMutableTreeNode("<html><b>" + textField_username_p2.getText());

		root.add(node);

		ImapFolderInfoCollection folderinfoc = null;
		try {
			folderinfoc = clientforimap_input.listFolders();
		} catch (Exception e1) {

			try {
				clientforimap_input.dispose();

			} finally {

			}
			connectionHandle1();
			e1.printStackTrace();
			folderinfoc = clientforimap_input.listFolders();

		}

		listFolderinfo.clear();
		listFolderinfostring.clear();

		for (ImapFolderInfo folderInfo : folderinfoc) {
			try {

				if (stop_tree) {
					break;
				}

				String folder = folderInfo.getName().replace("INBOX.", "");

				String s[] = folder.split("/");

				folder = s[s.length - 1];
				folder = folder.replaceAll("[\\[\\]]", "");
				DefaultMutableTreeNode t = null;

				if (nextTime == null) {
					obTh.ob.MessageLabel.setText(folder);
				}

				listFolderinfo.add(folderInfo);
				listFolderinfostring.add(folder);

				t = new DefaultMutableTreeNode("<html><b>" + folder);

				node.add(t);
				try {
					if (fileoption.equalsIgnoreCase("Icloud") || fileoption.equalsIgnoreCase("imap")) {
						readimap_subfolder(folderInfo, t, folder);
					} else if (folderInfo.hasChildren()) {
						readimap_subfolder(folderInfo, t, folder);

					}

				} catch (Exception e) {

				}

			} catch (Exception e) {
				e.printStackTrace();
				continue;
			}
		}

	}

	public void readimap_subfolder(ImapFolderInfo folderInfo, DefaultMutableTreeNode node, String pafolder) {
		ImapFolderInfoCollection folderInfoCollection = clientforimap_input.listFolders(folderInfo.getName());
		try {
			String delimString = clientforimap_input.getDelimiter();

			for (ImapFolderInfo folderInfo1 : folderInfoCollection) {
				try {

					if (stop_tree) {
						break;
					}

					String s1 = "";
					if (delimString.equalsIgnoreCase(".")) {
						s1 = getFileExtension(folderInfo1.getName());

					} else {
						String[] s = folderInfo1.getName().split("/");
						s1 = s[s.length - 1];
						s1 = s1.replaceAll("[\\[\\]]", "");

					}

					pafolder = pafolder + File.separator + s1;

					listFolderinfo.add(folderInfo1);
					listFolderinfostring.add(pafolder);
					if (nextTime == null) {
						obTh.ob.MessageLabel.setText(folderInfo1.getName());
					}
					DefaultMutableTreeNode t = new DefaultMutableTreeNode("<html><b>" + s1);
					node.add(t);
					try {
						if (fileoption.equalsIgnoreCase("Icloud") || fileoption.equalsIgnoreCase("imap")) {
							readimap_subfolder(folderInfo1, t, pafolder);
						} else if (folderInfo1.hasChildren()) {

							readimap_subfolder(folderInfo1, t, pafolder);

						}
					} catch (Exception e) {

					}

					pafolder = removefolder(pafolder);
				} catch (Exception e) {
					e.printStackTrace();
					continue;
				}

			}
		} catch (Exception e) {

		}

	}

	void readoperamail(File filearray, DefaultMutableTreeNode node) throws Exception {

		File[] files = filearray.listFiles();

		for (int i = 0; i < files.length; i++) {

			if (files[i].isFile()) {
				String extension = "";
				try {
					extension = getFileExtension(files[i]);
				} catch (Exception e) {
					extension = "";
				}
				if (files[i].length() > 0) {

					DefaultMutableTreeNode t1 = new DefaultMutableTreeNode(foldername);

					node.add(t1);
					if ((extension.equalsIgnoreCase("mbox") || extension.equalsIgnoreCase("")
							|| extension.equalsIgnoreCase("mbs"))) {
						lblNewLabel_3.setText(files[i].getName());
						CustomTreeNode t = new CustomTreeNode("<html><b>" + files[i].getName());
						t.filepath = files[i].getAbsolutePath();
						t1.add(t);
					}
				}
			} else {

				File[] fo = files[i].listFiles();

				if (fo.length > 0) {

					foldername = files[i].getName();

					Boolean cflag = false;
					int k = 0;
					try {
						k = Integer.valueOf(foldername);
						cflag = true;
					} catch (Exception e) {

					}

					if (cflag) {
						if (nummberofDigit(k) == 4) {
							foldername6 = foldername;
						} else {
							foldername6 = foldername6 + "-" + foldername;
						}

						readoperamail(files[i], node);
					} else {
						lblNewLabel_3.setText(foldername);

						DefaultMutableTreeNode t = new DefaultMutableTreeNode(foldername);

						node.add(t);

						readoperamail(files[i], t);
					}
				}
			}

		}

	}

	void readThunderbird(File filearray, DefaultMutableTreeNode node) throws Exception {

		File[] files = filearray.listFiles();

		for (int i = 0; i < files.length; i++) {

			if (files[i].isFile()) {
				String extension = "";
				try {
					extension = getFileExtension(files[i]);
				} catch (Exception e) {
					extension = "";
				}
				if (files[i].length() > 0) {
					if ((extension.equalsIgnoreCase("mbox") || extension.equalsIgnoreCase("")
							|| extension.equalsIgnoreCase("mbs"))) {
						lblNewLabel_3.setText(files[i].getName());
						CustomTreeNode t = new CustomTreeNode("<html><b>" + files[i].getName());
						t.filepath = files[i].getAbsolutePath();
						node.add(t);
					}
				}
			} else {

				File[] fo = files[i].listFiles();

				if (fo.length > 0) {

					foldername = files[i].getName();
					lblNewLabel_3.setText(foldername);

					DefaultMutableTreeNode t = new DefaultMutableTreeNode(foldername);

					node.add(t);

					readThunderbird(files[i], t);
				}
			}

		}

	}

	void readapple_mail(File filearray, DefaultMutableTreeNode node) throws Exception {

		File[] files = filearray.listFiles();

		for (int i = 0; i < files.length; i++) {

			if (files[i].isFile()) {
				String extension = "";
				try {
					extension = getFileExtension(files[i]);
				} catch (Exception e) {
					extension = "";
				}
				if (files[i].length() > 0) {
					if (extension.equalsIgnoreCase("mbox")) {
						lblNewLabel_3.setText(files[i].getName());
						CustomTreeNode t = new CustomTreeNode("<html><b>" + files[i].getName());
						t.filepath = files[i].getAbsolutePath();
						node.add(t);
						kl++;
					}
				}
			} else {

				File[] fo = files[i].listFiles();

				if (fo.length > 0) {
					if (!files[i].getName().equalsIgnoreCase("MailData")) {
						foldername = files[i].getName();
						lblNewLabel_3.setText(foldername);

						DefaultMutableTreeNode t = new DefaultMutableTreeNode(foldername);

						node.add(t);

						readapple_mail(files[i], t);
					}
				}
			}

		}

	}

	private static String getFileExtension(File file) {
		String fileName = file.getName();
		if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
			return fileName.substring(fileName.lastIndexOf(".") + 1);
		else
			return "";
	}

	public void readolmFile() {
		OlmStorage storage = null;

		FileStream stream = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Read);

		try {
			storage = new OlmStorage(stream.toInputStream());

		} catch (Exception e) {
			if (e.getMessage().contains("File not found")) {
				JOptionPane.showMessageDialog(frame,
						"File is in use in other application please close that application", messageboxtitle,
						JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));

			}

			else {
				JOptionPane.showMessageDialog(frame, "File is Currupted  Please Choose another file  ", messageboxtitle,
						JOptionPane.INFORMATION_MESSAGE,
						new ImageIcon(Main_Frame.class.getResource("/information.png")));
			}
		}

		model = (DefaultTreeModel) tree.getModel();

		String filename = filepath.replace(",", "");

		root = new DefaultMutableTreeNode("<html><b>" + filename);

		model.setRoot(root);

		DefaultMutableTreeNode node = new DefaultMutableTreeNode("<html><b>" + fileoption);

		root.add(node);

		try {
			for (OlmFolder folder : storage.getFolderHierarchy()) {

				if (stop_tree) {
					break;
				}

				String foldername = folder.getName().replace(",", "");

				lblNewLabel_3.setText(foldername);

				DefaultMutableTreeNode c = new DefaultMutableTreeNode(
						"<html><b>" + foldername.replaceAll("[\\[\\]]", ""));

				node.add(c);

				if (folder.getSubFolders().size() > 0) {

					getFolder(folder, c);

				}

			}

		} catch (Exception e) {

			e.printStackTrace();

			return;
		} finally {
			storage.dispose();

		}

	}

	private void getFolder(OlmFolder folder, DefaultMutableTreeNode node1) {

		for (OlmFolder subFolder : folder.getSubFolders()) {

			if (stop) {
				break;
			}
			String foldername = subFolder.getName().replace(",", "");
			lblNewLabel_3.setText(foldername);

			DefaultMutableTreeNode nd = new DefaultMutableTreeNode("<html><b>" + foldername.replaceAll("[\\[\\]]", ""));
			node1.add(nd);

			if (subFolder.getSubFolders().size() > 0) {

				getFolder(subFolder, nd);

			}

		}
	}

	public static void visitAllNodes(DefaultMutableTreeNode roe) {

		@SuppressWarnings("unchecked")
		Enumeration<TreeNode> e = roe.depthFirstEnumeration();
		while (e.hasMoreElements()) {
			DefaultMutableTreeNode node = (DefaultMutableTreeNode) e.nextElement();
			lists.add(node);
			listst.add(node.toString().replace("<html><b>", ""));

		}

	}

	public void readmailFile() {
		MailMessage message1 = MailMessage.load(filepath);

		MailConversionOptions option = new MailConversionOptions();
		MapiMessage msg = MapiMessage.fromMailMessage(message1, MapiConversionOptions.getASCIIFormat());
		MailMessage message = msg.toMailMessage(option);

		model = (DefaultTreeModel) tree.getModel();

		root = new DefaultMutableTreeNode("<html><b>" + hostName);

		model.setRoot(root);

		String filepath = filepath(file);

		DefaultMutableTreeNode node = new DefaultMutableTreeNode("<html><b>" + filepath);

		root.add(node);

		DefaultMutableTreeNode child = new DefaultMutableTreeNode("<html><b>" + file.getName());

		node.add(child);

		lblTotalMessageCount.setText("Total Message Count : " + 1);

		String from = "NA";
		try {
			from = message.getFrom().toString();
		} catch (Exception e) {
			from = "NA";
		}

		String Date = null;
		try {
			Date = message1.getDate().toString();
		} catch (Exception e) {
			Date = "NA";
		}

		String Subject = "NA";
		try {
			Subject = message.getSubject();
		} catch (Exception e) {
			Subject = "NA";
		}

		mode = (DefaultTableModel) table_fileinformation.getModel();

		if (msg.getAttachments().size() > 0) {
			ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
			JLabel imagelabl = new JLabel();
			imagelabl.setIcon(icon);
			mode = (DefaultTableModel) table_fileinformation.getModel();

			mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + Date, imagelabl });
		} else {
			mode = (DefaultTableModel) table_fileinformation.getModel();

			mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + Date });
		}

	}

	private void fileinformation_Zimbra() {

		try {
			int i2 = 1;
			Boolean check = false;
			TgzReader reader = new TgzReader(filepath);
			while (reader.readNextMessage()) {
				String folder = reader.getCurrentDirectory();
				System.out.println(folder);
				System.out.println(foldername);
				if (check) {
					if (!(folder.equalsIgnoreCase(foldername))) {
						break;
					}
				}
				if (folder.equalsIgnoreCase(foldername)) {
					check = true;
					System.out.println("in here");
					MailMessage msg1 = reader.getCurrentMessage();
					MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();

					MapiMessage msg = MapiMessage.fromMailMessage(msg1, d);
					lblTotalMessageCount.setText("Total Message Count : " + i2);
					i2++;
					listmail.add(msg1);
					String from = "NA";
					try {
						from = msg.getSenderEmailAddress();
					} catch (Exception e) {

					}

					Date DeliveryTime = null;
					try {
						DeliveryTime = msg.getDeliveryTime();
					} catch (Exception e) {

					}

					String Subject = "NA";
					try {
						Subject = msg.getSubject();
					} catch (Exception e) {

					}

					if (msg.getAttachments().size() > 0) {
						ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
						JLabel imagelabl = new JLabel();
						imagelabl.setIcon(icon);
						mode = (DefaultTableModel) table_fileinformation.getModel();

						mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
								"<html><b>" + DeliveryTime, imagelabl });
					} else {
						mode = (DefaultTableModel) table_fileinformation.getModel();

						mode.addRow(
								new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + DeliveryTime });
					}

				}
			}

		} catch (Exception e) {

		}

	}

	public void fileInhformation_Ost_Pst() throws Exception {
		try {

			FolderInfo f1 = pst.getRootFolder();
			String f1nmae = f1.getDisplayName().replaceAll("[\\[\\]]", "");
			if (f1nmae.equalsIgnoreCase("")) {
				f1nmae = "Root Folder";
			}
			if (foldername.equals(f1nmae)) {
				MessageInfoCollection messageInfoCollection = f1.getContents();
				int i2 = 1;
				for (int j = 0; j < messageInfoCollection.size(); j++)

				{

					try {
						if (Stoppreview) {
							break;
						}

						MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(j);
						MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
						MailConversionOptions de = new MailConversionOptions();
						MapiMessage contact1 = (MapiMessage) pst.extractMessage(messageInfo);
						MailMessage mess = contact1.toMailMessage(de);
						MapiMessage contact = MapiMessage.fromMailMessage(mess, d);

						listmapi.add(contact);
						lblTotalMessageCount.setText("Total Message Count : " + i2);
						i2++;
						String from = "NA";
						String Subject = "NA";
						Date DeliveryTime = null;
						try {
							from = contact.getSenderEmailAddress();
						} catch (Exception a) {
							from = "";
						}
						try {
							Subject = contact.getSubject();
						} catch (Exception a) {
							Subject = "";
						}
						try {
							DeliveryTime = contact.getDeliveryTime();
						} catch (Exception a) {

						}

						if (contact.getAttachments().size() > 0) {
							ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
							JLabel imagelabl = new JLabel();
							imagelabl.setIcon(icon);
							mode = (DefaultTableModel) table_fileinformation.getModel();

							mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
									"<html><b>" + DeliveryTime, imagelabl });
						} else {
							mode = (DefaultTableModel) table_fileinformation.getModel();

							mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
									"<html><b>" + DeliveryTime });
						}

					} catch (Exception e) {
						continue;
					}

				}

				path = "";

			}

			FolderInfoCollection folderInfoCollection = pst.getRootFolder().getSubFolders();

			for (int i = 0; i < folderInfoCollection.size(); i++) {
				try {

					FolderInfo f = folderInfoCollection.get_Item(i);
					String folder = f.getDisplayName();
					folder = folder.replace(",", "").replace(".", "");
					folder = getRidOfIllegalFileNameCharacters(folder);
					folder = folder.replaceAll("[\\[\\]]", "");
					folder = folder.trim();
					path = f1nmae + File.separator + folder;
					int size = f.getContentCount();

					foldername = foldername.replace("[" + size + "]", "");

					if (foldername.equals(path)) {

						folderInfo = folderInfoCollection.get_Item(i);
						MessageInfoCollection messageInfoCollection = f.getContents();
						int i2 = 1;
						for (int j = 0; j < messageInfoCollection.size(); j++)

						{

							try {
								if (Stoppreview) {
									break;
								}

								MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(j);
								MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
								MailConversionOptions de = new MailConversionOptions();
								MapiMessage contact1 = (MapiMessage) pst.extractMessage(messageInfo);
								MailMessage mess = contact1.toMailMessage(de);
								MapiMessage contact = MapiMessage.fromMailMessage(mess, d);

								listmapi.add(contact);
								lblTotalMessageCount.setText("Total Message Count : " + i2);
								i2++;
								String from = "NA";
								String Subject = "NA";
								Date DeliveryTime = null;
								try {
									from = contact.getSenderEmailAddress();
								} catch (Exception a) {
									from = "";
								}
								try {
									Subject = contact.getSubject();
								} catch (Exception a) {
									Subject = "";
								}
								try {
									DeliveryTime = contact.getDeliveryTime();
								} catch (Exception a) {

								}

								if (contact.getAttachments().size() > 0) {
									ImageIcon icon = new ImageIcon(
											Main_Frame.class.getResource("/attachment-icon.png"));
									JLabel imagelabl = new JLabel();
									imagelabl.setIcon(icon);
									mode = (DefaultTableModel) table_fileinformation.getModel();

									mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
											"<html><b>" + DeliveryTime, imagelabl });
								} else {
									mode = (DefaultTableModel) table_fileinformation.getModel();

									mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
											"<html><b>" + DeliveryTime });
								}

							} catch (Exception e) {
								continue;
							}

						}

						path = "";
						break;

					}
					if (f.hasSubFolders()) {
						fileInhformationsubfolder_Ost_Pst(f);

					}
				} catch (Exception e) {
					continue;
				}

			}
		} catch (Exception e) {

		}
	}

	public void fileInhformationsubfolder_Ost_Pst(FolderInfo folder) {
		FolderInfoCollection folderInfoCollection = folder.getSubFolders();

		for (int i = 0; i < folderInfoCollection.size(); i++) {
			try {

				FolderInfo folderInf = folderInfoCollection.get_Item(i);
				int size = folderInf.getContentCount();
				foldername = foldername.replace("[" + size + "]", "");

				String Folder = folderInf.getDisplayName();
				Folder = Folder.replace(",", "").replace(".", "");
				Folder = getRidOfIllegalFileNameCharacters(Folder);
				Folder = Folder.replaceAll("[\\[\\]]", "");
				Folder = Folder.trim();

				path = path + File.separator + Folder;

				if (foldername.equals(path)) {
					folderInfo = folderInfoCollection.get_Item(i);
					MessageInfoCollection messageInfoCollection = folderInf.getContents();
					int i2 = 1;
					for (int j = 0; j < messageInfoCollection.size(); j++)

					{

						try {
							if (Stoppreview) {
								break;
							}

							MessageInfo messageInfo = (MessageInfo) messageInfoCollection.get_Item(j);

							MapiMessage contact1 = (MapiMessage) pst.extractMessage(messageInfo);
							MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
							MailConversionOptions de = new MailConversionOptions();
							MailMessage mess = contact1.toMailMessage(de);
							MapiMessage contact = MapiMessage.fromMailMessage(mess, d);
							listmapi.add(contact);
							String from = "NA";
							String Subject = "NA";
							Date DeliveryTime = null;
							try {
								from = contact.getSenderEmailAddress();
							} catch (Exception a) {
								from = "";
							}
							try {
								Subject = contact.getSubject();
							} catch (Exception a) {
								Subject = "";
							}
							try {
								DeliveryTime = contact.getDeliveryTime();
							} catch (Exception a) {

							}
							lblTotalMessageCount.setText("Total Message Count : " + i2);
							i2++;
							if (contact.getAttachments().size() > 0) {
								ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
								JLabel imagelabl = new JLabel();
								imagelabl.setIcon(icon);
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime, imagelabl });
							} else {
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime });
							}

						} catch (Exception e) {
							continue;
						}

					}

					path = "";
					break;

				}
				if (folderInf.hasSubFolders()) {
					fileInhformationsubfolder_Ost_Pst(folderInf);

				}
				path = removefolder(path);

			} catch (Exception e) {
				continue;
			}
		}

	}

	public void fileInformation_on_mbox() {
		FileStream stream = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Read);

		MboxrdStorageReader mbox = new MboxrdStorageReader(stream.toInputStream(), new MboxLoadOptions());

		MailMessage message1 = mbox.readNextMessage();
		int i2 = 1;
		while (message1 != null) {
			try {
				if (Stoppreview) {
					break;
				}

				MailConversionOptions option = new MailConversionOptions();
				MapiMessage msg = MapiMessage.fromMailMessage(message1, MapiConversionOptions.getASCIIFormat());
				MailMessage message = msg.toMailMessage(option);

				String from = "NA";
				try {
					from = message.getFrom().toString();
				} catch (Exception e1) {

				}

				String Subject = "NA";
				try {
					Subject = message.getSubject();
				} catch (Exception e1) {

				}

				String Date = "NA";
				try {
					Date = message.getDate().toString();
				} catch (Exception e1) {

				}
				lblTotalMessageCount.setText("Total Message Count : " + i2);
				i2++;
				listmail.add(message);
				if (message.getAttachments().size() > 0) {
					ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
					JLabel imagelabl = new JLabel();
					imagelabl.setIcon(icon);
					mode = (DefaultTableModel) table_fileinformation.getModel();

					mode.addRow(
							new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + Date, imagelabl });
				} else {
					mode = (DefaultTableModel) table_fileinformation.getModel();

					mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + Date });
				}

				try {
					message1 = mbox.readNextMessage();

				} catch (Error e) {
					logger.warning("Error : " + e.getMessage() + System.lineSeparator());
				} catch (Exception e) {
					logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
					continue;
				}
			} catch (Exception e) {
				logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
				continue;
			}
		}
		mbox.close();
	}

	public void fileInformation_on_Thunderbird() {

		if (new File(filepath).isFile()) {

			FileStream stream = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Read);

			MboxrdStorageReader mbox = new MboxrdStorageReader(stream.toInputStream(), false);

			MailMessage message1 = mbox.readNextMessage();
			while (message1 != null) {
				try {
					if (Stoppreview) {
						break;
					}

					MailConversionOptions option = new MailConversionOptions();
					MapiMessage msg = MapiMessage.fromMailMessage(message1, MapiConversionOptions.getASCIIFormat());
					MailMessage message = msg.toMailMessage(option);
					listmail.add(message);
					String from = "NA";
					try {
						from = message.getFrom().toString();
					} catch (Exception e1) {

					}

					String Subject = "NA";
					try {
						Subject = message.getSubject();
					} catch (Exception e1) {

					}

					String Date = "NA";
					try {
						Date = message.getDate().toString();
					} catch (Exception e1) {

					}
					if (message.getAttachments().size() > 0) {
						ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
						JLabel imagelabl = new JLabel();
						imagelabl.setIcon(icon);
						mode = (DefaultTableModel) table_fileinformation.getModel();

						mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + Date,
								imagelabl });
					} else {
						mode = (DefaultTableModel) table_fileinformation.getModel();

						mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject, "<html><b>" + Date });
					}

					try {
						message1 = mbox.readNextMessage();

					} catch (Error e) {
						logger.warning("Error : " + e.getMessage() + System.lineSeparator());
					} catch (Exception e) {
						logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
						continue;
					}
				} catch (Exception e) {
					logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
					continue;
				}
			}
			mbox.close();
		} else {
			JOptionPane.showMessageDialog(frame, "Please click on a file", messageboxtitle, JOptionPane.ERROR_MESSAGE,
					new ImageIcon(Main_Frame.class.getResource("/information.png")));
		}

	}

	public void fileInhformation_exchange() throws Exception {

		String rootUri = clientforexchange_input.getMailboxInfo().getRootUri();
		ExchangeFolderInfoCollection folderInfoCollection = clientforexchange_input.listSubFolders(rootUri);

		for (int i = 0; i < folderInfoCollection.size(); i++) {
			try {
				ExchangeFolderInfo f = folderInfoCollection.get_Item(i);
				path = File.separator + f.getDisplayName().replaceAll("[\\[\\]]", "");

				if (foldername.equalsIgnoreCase(path)) {
					ExchangeMessageInfoCollection msgCollection = clientforexchange_input.listMessages(f.getUri());

					for (ExchangeMessageInfo msgInfo : msgCollection) {
						try {
							if (Stoppreview) {
								break;
							}

							String strMessageURI = msgInfo.getUniqueUri();
							MailMessage msg1 = clientforexchange_input.fetchMessage(strMessageURI);
							listmail.add(msg1);

							String from = msg1.getFrom().getAddress();

							Date DeliveryTime = msg1.getDate();

							String Subject = msg1.getSubject();
							try {
								byte[] arr = from.getBytes("UTF-16");
								byte[] abrr = Subject.getBytes("UTF-16");

								from = new String(arr);
								Subject = new String(abrr);
							} catch (Exception e) {

							}

							if (msg1.getAttachments().size() > 0) {
								ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
								JLabel imagelabl = new JLabel();
								imagelabl.setIcon(icon);
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime, imagelabl });
							} else {
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime });
							}

						}

						catch (Error e) {
							logger.warning("Error : " + e.getMessage() + System.lineSeparator());
						} catch (Exception e) {
							logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
							e.printStackTrace();
							continue;
						}
					}
					exchangeFolderInfo = folderInfoCollection.get_Item(i);
					path = "";
					break;

				}
				if (f.getChildFolderCount() > 0) {
					fileInhformationsubfolder_exchange(f);

				}
			} catch (Exception e) {

				continue;
			}

		}

	}

	public void fileInhformationsubfolder_exchange(ExchangeFolderInfo folder) {
		ExchangeFolderInfoCollection folderInfoCollection = clientforexchange_input.listSubFolders(folder);

		for (int i = 0; i < folderInfoCollection.size(); i++) {
			try {
				ExchangeFolderInfo f1 = folderInfoCollection.get_Item(i);
				path = path + File.separator + f1.getDisplayName().replaceAll("[\\[\\]]", "");

				if (foldername.equals(path)) {
					ExchangeMessageInfoCollection msgCollection = clientforexchange_input.listMessages(f1.getUri());

					for (ExchangeMessageInfo msgInfo : msgCollection) {
						try {
							if (Stoppreview) {
								break;
							}

							String strMessageURI = msgInfo.getUniqueUri();
							MailMessage msg1 = clientforexchange_input.fetchMessage(strMessageURI);
							listmail.add(msg1);

							String from = msg1.getFrom().getAddress();

							Date DeliveryTime = msg1.getDate();

							String Subject = msg1.getSubject();

							if (msg1.getAttachments().size() > 0) {
								ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
								JLabel imagelabl = new JLabel();
								imagelabl.setIcon(icon);
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime, imagelabl });
							} else {
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime });
							}

						} catch (Error e) {
							logger.warning("Error : " + e.getMessage() + System.lineSeparator());
						} catch (Exception e) {
							logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
							continue;
						}
					}
					exchangeFolderInfo = folderInfoCollection.get_Item(i);
					path = "";
					break;

				}
				if (f1.getChildFolderCount() > 0) {
					fileInhformationsubfolder_exchange(f1);

				}
				path = removefolder(path);
			} catch (Exception e) {

				continue;
			}

		}

	}

	public void fileInhformation_imap() throws Exception {

		ImapFolderInfoCollection folderinfoc = clientforimap_input.listFolders(iconnforimap_input);

		for (int i = 0; i < folderinfoc.size(); i++) {
			try {
				ImapFolderInfo f = folderinfoc.get_Item(i);
				path = File.separator + f.getName();

				if (foldername.equalsIgnoreCase(path)) {
					imapFolderInfo = folderinfoc.get_Item(i);
					ImapMessageInfoCollection msgCollection = clientforimap_input.listMessages(f.getName());
					clientforimap_input.selectFolder(f.getName());

					for (int j = 0; j < msgCollection.size(); j++) {
						try {
							ImapMessageInfo mess = msgCollection.get_Item(j);

							MailMessage msg1 = clientforimap_input.fetchMessage(mess.getUniqueId());
							if (Stoppreview) {
								break;
							}

							listmail.add(msg1);

							String from = mess.getSender().toString();

							Date DeliveryTime = msg1.getDate();

							String Subject = msg1.getSubject();

							if (msg1.getAttachments().size() > 0) {
								ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
								JLabel imagelabl = new JLabel();
								imagelabl.setIcon(icon);
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime, imagelabl });
							} else {
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime });
							}

						} catch (Error e) {
							logger.warning("Error : " + e.getMessage() + System.lineSeparator());
						} catch (Exception e) {
							logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
							e.printStackTrace();
							continue;
						}
					}

					path = "";
					break;

				}
				if (f.hasChildren()) {
					fileInhformationsubfolder_imap(f);

				}
			} catch (Exception e) {

				continue;
			}
		}

	}

	public void fileInhformationsubfolder_imap(ImapFolderInfo folder) {
		ImapFolderInfoCollection folderinfoc = clientforimap_input.listFolders(folder.getName());

		for (int i = 0; i < folderinfoc.size(); i++) {
			try {
				ImapFolderInfo f1 = folderinfoc.get_Item(i);

				path = File.separator + f1.getName().replace("/", File.separator);

				if (foldername.equals(path)) {
					ImapMessageInfoCollection msgCollection = clientforimap_input.listMessages(f1.getName());

					imapFolderInfo = null;
					imapFolderInfo = folderinfoc.get_Item(i);
					clientforimap_input.selectFolder(f1.getName());
					for (int j = 0; j < msgCollection.size(); j++) {

						try {
							ImapMessageInfo msgInfo = msgCollection.get_Item(j);
							MailMessage msg1 = clientforimap_input.fetchMessage(msgInfo.getUniqueId());
							if (Stoppreview) {
								break;
							}

							listmail.add(msg1);

							String from = msgInfo.getSender().toString();

							Date DeliveryTime = msg1.getDate();

							String Subject = msgInfo.getSubject();

							if (msg1.getAttachments().size() > 0) {
								ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
								JLabel imagelabl = new JLabel();
								imagelabl.setIcon(icon);
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime, imagelabl });
							} else {
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime });
							}

						} catch (Error e) {
							logger.warning("Error : " + e.getMessage() + System.lineSeparator());
						} catch (Exception e) {
							logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
							e.printStackTrace();
							continue;
						}
					}

					path = "";
					break;

				}
				if (f1.hasChildren()) {
					fileInhformationsubfolder_imap(f1);

				}
			} catch (Exception e) {
				logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
				continue;
			}
		}

	}

	public void fileinformation_olm() {
		OlmStorage storage = new OlmStorage(filepath);

		for (OlmFolder folder : storage.getFolderHierarchy()) {

			String pa1 = folder.getName();

			int i2 = 1;

			if (pa1.equalsIgnoreCase(foldername.trim())) {

				if (folder.hasMessages()) {

					long foldercount = folder.getMessageCount();
					Iterator<MapiMessage> it = storage.enumerateMessages(folder).iterator();
					for (int i11 = 0; i11 < foldercount; i11++) {
						try {
							if (!it.hasNext() || (it.next() == null)) {

								continue;
							}
						} catch (Exception e) {

						}

						if (Stoppreview) {
							break;
						}

						MapiMessage msg1 = it.next();

						MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
						MailConversionOptions de = new MailConversionOptions();
						MailMessage mess = msg1.toMailMessage(de);
						MapiMessage msg = MapiMessage.fromMailMessage(mess, d);

						listmapi.add(msg);
						lblTotalMessageCount.setText("Total Message Count : " + i2);
						i2++;
						String from = "";
						try {
							from = msg.getSenderEmailAddress();
						} catch (Exception e) {
							from = "";
						}

						Date DeliveryTime = null;
						try {
							DeliveryTime = msg.getDeliveryTime();
						} catch (Exception e) {

						}

						String Subject = "";
						try {
							Subject = msg.getSubject();
						} catch (Exception e) {
							Subject = "";
						}

						if (msg.getAttachments().size() > 0) {
							ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
							JLabel imagelabl = new JLabel();
							imagelabl.setIcon(icon);
							mode = (DefaultTableModel) table_fileinformation.getModel();

							mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
									"<html><b>" + DeliveryTime, imagelabl });
						} else {
							mode = (DefaultTableModel) table_fileinformation.getModel();

							mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
									"<html><b>" + DeliveryTime });
						}

					}
				}
			}

			if (folder.getSubFolders().size() > 0) {

				getFolderolminfo(folder, pa1, storage);

			}

		}

	}

	private void getFolderolminfo(OlmFolder folder, String rootFolder, OlmStorage storage) {

		for (OlmFolder subFolder : folder.getSubFolders()) {

			String curpath = rootFolder + File.separator + subFolder.getName().replaceAll("[\\[\\]]", "");

			if (curpath.equals(foldername)) {
				int i2 = 1;
				if (subFolder.hasMessages()) {
					long foldercount = subFolder.getMessageCount();
					Iterator<MapiMessage> it = storage.enumerateMessages(subFolder).iterator();

					for (int i11 = 0; i11 < foldercount; i11++) {
						try {
							try {
								if (!it.hasNext() || (it.next() == null)) {

									continue;
								}
							} catch (Exception e) {

							}

							if (Stoppreview) {
								break;
							}

							MapiMessage msg1 = it.next();
							MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
							MailConversionOptions de = new MailConversionOptions();
							MailMessage mess = msg1.toMailMessage(de);
							MapiMessage msg = MapiMessage.fromMailMessage(mess, d);

							listmapi.add(msg);
							lblTotalMessageCount.setText("Total Message Count : " + i2);
							i2++;
							String from = "";
							try {
								from = msg.getSenderEmailAddress();
							} catch (Exception e) {
								from = "";
							}

							Date DeliveryTime = null;
							try {
								DeliveryTime = msg.getDeliveryTime();
							} catch (Exception e) {

							}

							String Subject = "";
							try {
								Subject = msg.getSubject();
							} catch (Exception e) {
								Subject = "";
							}

							if (msg.getAttachments().size() > 0) {
								ImageIcon icon = new ImageIcon(Main_Frame.class.getResource("/attachment-icon.png"));
								JLabel imagelabl = new JLabel();
								imagelabl.setIcon(icon);
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime, imagelabl });
							} else {
								mode = (DefaultTableModel) table_fileinformation.getModel();

								mode.addRow(new Object[] { "<html><b>" + from, "<html><b>" + Subject,
										"<html><b>" + DeliveryTime });
							}
						} catch (Exception e) {
							logger.warning("Exception : " + e.getMessage() + System.lineSeparator());
							continue;
						}
					}
				}

				break;

			}

			if (subFolder.getSubFolders().size() > 0) {

				getFolderolminfo(subFolder, curpath, storage);

			}
			curpath = removefolder(curpath);

		}

	}

	boolean checkdate(MapiMessage message1) {
		for (int k = 0; k < fromList.size(); k++) {
			fromdate = fromList.get(k);
			todate = toList.get(k);
			System.out.println(" Date : " + message1.getDeliveryTime());
			if (message1.getDeliveryTime() != null) {
				if (message1.getDeliveryTime().after(fromdate) && message1.getDeliveryTime().before(todate)) {
					datevalidflag = true;
					break;

				} else {
					datevalidflag = false;
				}
			}
		}

		return datevalidflag;
	}

	void nextsign() {

		// obTh.ob.setVisible(false);
		if (fileoption.equalsIgnoreCase("OFFICE 365") || fileoption.equalsIgnoreCase("Live Exchange")
				|| fileoption.equalsIgnoreCase("Hotmail")) {

			readexchange();

		} else {

			readimap();

		}

		Icon open = new ImageIcon(Main_Frame.class.getResource("/Open-folder-accept-icon.png"));
		Icon close = new ImageIcon(Main_Frame.class.getResource("/closed-folder-add-icon.png"));
		Icon Ram = new ImageIcon(Main_Frame.class.getResource("/leaf-icon.png"));
		DefaultCheckboxTreeCellRenderer render = (DefaultCheckboxTreeCellRenderer) tree.getCellRenderer();

		render.setClosedIcon(close);
		render.setOpenIcon(open);
		render.setLeafIcon(Ram);

		tree.expandRow(0);

		btn_next_pane2.doClick();

		tree.expandAll();

	}

	void sign() {

		SwingWorker sw1 = new SwingWorker() {
			@Override
			protected Object doInBackground() {
				obTh = new loadingThreadclassformailbox(frame);

				obTh.start();
				// obTh.ob.setVisible(false);
				if (fileoption.equalsIgnoreCase("OFFICE 365") || fileoption.equalsIgnoreCase("Live Exchange")
						|| fileoption.equalsIgnoreCase("Hotmail")) {

					CardLayout card = (CardLayout) Cardlayout.getLayout();
					card.show(Cardlayout, "panel_2");
					readexchange();

				} else {

					CardLayout card = (CardLayout) Cardlayout.getLayout();
					card.show(Cardlayout, "panel_2");

					readimap();

				}

				return null;
			}

			@Override
			protected void done() {

				Icon open = new ImageIcon(Main_Frame.class.getResource("/Open-folder-accept-icon.png"));
				Icon close = new ImageIcon(Main_Frame.class.getResource("/closed-folder-add-icon.png"));
				Icon Ram = new ImageIcon(Main_Frame.class.getResource("/leaf-icon.png"));
				DefaultCheckboxTreeCellRenderer render = (DefaultCheckboxTreeCellRenderer) tree.getCellRenderer();

				render.setClosedIcon(close);
				render.setOpenIcon(open);
				render.setLeafIcon(Ram);

				tree.expandRow(0);

				tree.expandAll();
				obTh.close();

			}
		};

		sw1.execute();
	}

	public boolean isValid(String email) {
		String emailRegex = "^[a-zA-Z0-9_+&*-]+(?:\\." + "[a-zA-Z0-9_+&*-]+)*@" + "(?:[a-zA-Z0-9-]+\\.)+[a-z"
				+ "A-Z]{2,7}$";

		Pattern pat = Pattern.compile(emailRegex);
		if (email == null)
			return false;
		return pat.matcher(email).matches();
	}

	public void Mapimess_CSV(MapiMessage message, CSVWriter writer) {

		String subname = getRidOfIllegalFileNameCharacters(namingconventionmapi(message));
		if (message.getMessageClass().equals("IPM.Contact")) {
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
			String prefix = null;
			String fileunder = null;
			String fileunderid = null;
			String generation = null;
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

			MapiContact con = (MapiContact) message.toMapiMessageItem();
			MapiContactProfessionalPropertySet ProfPropSet = null;
			try {
				ProfPropSet = con.getProfessionalInfo();
			} catch (Exception e) {

			}
			MapiContactElectronicAddress email1 = null;
			try {
				email1 = con.getElectronicAddresses().getEmail1();
			} catch (Exception e) {

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
			try {
				NamePropSet = con.getNameInfo();
			} catch (Exception e) {

			}
			try {

				title = ProfPropSet.getTitle();
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

				bussinessfaxaddresstype = bussinessfax.getAddressType();
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

				prefix = NamePropSet.getDisplayNamePrefix();
			} catch (Exception ep) {
				prefix = "";
			}
			try {
				if (prefix.equalsIgnoreCase("null") || prefix.contains("meta") || prefix.contains("aspose")) {
					prefix = "NA";
				}
			} catch (Exception e1) {
				prefix = "NA";
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

				generation = String.valueOf(NamePropSet.getGeneration());
			} catch (Exception ep) {
				generation = "";
			}
			try {
				if (generation.equalsIgnoreCase("null") || generation.contains("meta")
						|| generation.contains("aspose")) {
					generation = "NA";
				}
			} catch (Exception e1) {
				generation = "NA";
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

				homeCity = String.valueOf(contacthomephys.getCity());
			} catch (Exception ep) {
				homeCity = "";
			}
			try {
				if (homeCity.equalsIgnoreCase("null") || homeCity.contains("meta") || homeCity.contains("aspose")) {
					homeCity = "NA";
				}
			} catch (Exception e1) {
				homeCity = "NA";
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

			try {
				String[] data1 = { firstName, middleName, lastname, prefix, Email1addresstype, Email1displayname,
						Email1address, Email1fax, Email2addresstype, Email2displayname, Email2fax, Email2address,
						Email3address, Email3addresstype, Email3displayname, Email3fax, homefaxaddress,
						homefaxaddresstype, homefaxdisplayname, homefaxno, primaryfaxaddress, primaryfaxaddresstype,
						primaryfaxdisplayname, primaryfaxno, bussinessfaxaddress, bussinessfaxaddresstype,
						bussinessfaxdisplayname, bussinessfaxno, WeddingAnniversary, birthday, fileunder, fileunderid,
						generation, title, account, BusinessHomePage, ComputerNetworkName, CustomerId, FreeBusyLocation,
						FtpSite, Gender, GovernmentIdNumber, Hobbies, Html, InstantMessagingAddress, Language, Location,
						Notes, OrganizationalIdNumber, PersonalHomePage, ReferredByName, SpouseName, homeAddress,
						homegetStreet, homeCity, homeCountry, homeCountryCode, homePostalCode, homegetPostOfficeBox,
						homeStateOrProvince, otherAddress, othergetStreet, otherCity, otherCountry, otherCountryCode,
						otherPostalCode, othergetPostOfficeBox, otherStateOrProvince, workAddress, workgetStreet,
						workCity, workCountry, workCountryCode, workPostalCode, workgetPostOfficeBox,
						workStateOrProvince, Assistant, CompanyName, DepartmentName, ManagerName, OfficeLocation,
						Profession, getTitle, AssistantTelephoneNumber, Business2TelephoneNumber,
						BusinessTelephoneNumber, CallbackTelephoneNumber, CarTelephoneNumber,
						CompanyMainTelephoneNumber, Home2TelephoneNumber, HomeTelephoneNumber, IsdnNumber,
						MobileTelephoneNumber, OtherTelephoneNumber, PagerTelephoneNumber, PrimaryTelephoneNumber,
						RadioTelephoneNumber, TelexNumber, TtyTddPhoneNumber };

				writer.writeNext(data1);

			} catch (Error e) {
				logger.warning("ERROR : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

			catch (Exception e) {
				logger.warning("Exception : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

		} else if (message.getMessageClass().equals("IPM.Appointment")
				|| message.getMessageClass().contains("IPM.Schedule.Meeting")) {

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
				String[] data1 = { subject, startdate, enddate, alldayevent, reminder, requiredattend, remindertime,
						requiredattend, categories, location, mileage };

				writer.writeNext(data1);

			} catch (Error e) {
				logger.warning("ERROR : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

			catch (Exception e) {
				logger.warning("Exception : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

		} else if (message.getMessageClass().equals("IPM.StickyNote") || message.getMessageClass().equals("IPM.Task")) {

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

			try {
				String date = null;
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
				String subject = null;
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
				String getBody = null;
				try {
					getBody = message.getBody();
				} catch (Exception e) {
					try {
						getBody = message.getBodyHtml();
					} catch (Exception e2) {
						getBody = "NA";
					}

				}

				try {
					if (getBody.equalsIgnoreCase("null") || getBody.contains("meta") || getBody.contains("aspose")) {
						getBody = "NA";
					}
					if (getBody.length() >= 150) {
						getBody = getBody.substring(0, 150);
					}

				} catch (Exception e1) {
					getBody = "NA";
				}

				String getSenderEmailAddress = null;
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
				String getReplyTo = null;
				try {
					for (int i = 0; i < message.getRecipients().size(); i++) {
						String toid = null;
						try {
							toid = message.getRecipients().get_Item(i).getEmailAddress();
						} catch (Exception e) {

						}
						if (i == 0) {
							getReplyTo = toid;
						} else {
							getReplyTo = getReplyTo + "," + toid;

						}

					}
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
							+ File.separator + subname);

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
				logger.warning("ERROR : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

			catch (Exception e) {
				logger.warning("Exception : " + e.getMessage() + "Message" + " " + message.getDeliveryTime()
						+ System.lineSeparator());
			}

		}

		try {
			count_destination++;

			for (int j = 0; j < message.getAttachments().size(); j++) {

				MapiAttachment att = message.getAttachments().get_Item(j);

				String s = getFileExtension(att.getLongFileName());
				String attFileName = getRidOfIllegalFileNameCharacters(att.getLongFileName().replace("." + s, ""));

				att.save(destination_path + File.separator + path + File.separator + "Attachment" + File.separator
						+ subname + File.separator + attFileName + "." + s);

			}
		} catch (Exception e) {

		}

	}

	public void Mailmess_CSV(MailMessage message, CSVWriter writer) {

		String subname = getRidOfIllegalFileNameCharacters(namingconventionmail(message));

		try {
			String date = null;
			try {
				date = message.getDate().toString();
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
			String subject = null;
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
			String getBody = null;
			try {
				getBody = message.getBody();
			} catch (Exception e) {

				getBody = "NA";

			}

			try {
				if (getBody.equalsIgnoreCase("null") || getBody.contains("meta") || getBody.contains("aspose")) {
					getBody = "NA";
				}
				if (getBody.length() >= 150) {
					getBody = getBody.substring(0, 150);
				}

			} catch (Exception e1) {
				getBody = "NA";
			}

			String getSenderEmailAddress = null;

			try {

				getSenderEmailAddress = message.getSender().getAddress();
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

			String getDisplayCc = null;
			try {

				MailAddressCollection mac = message.getCC();

				for (int i = 0; i < mac.size(); i++) {

					if (i == 0) {
						getDisplayCc = mac.get_Item(i).getAddress();
					} else {

						getDisplayCc = getDisplayCc + "," + mac.get_Item(i).getAddress();

					}
				}
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

				MailAddressCollection mac = message.getBcc();

				for (int i = 0; i < mac.size(); i++) {

					if (i == 0) {
						getDisplayBcc = mac.get_Item(i).getAddress();
					} else {

						getDisplayBcc = getDisplayBcc + "," + mac.get_Item(i).getAddress();

					}
				}

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

			File fd = new File(destination_path + File.separator + path + File.separator + "Attachment" + File.separator
					+ subname);

			fd.mkdirs();

			String[] data1 = { date, subject, getBody, getSenderEmailAddress, getDisplayCc, getDisplayBcc,
					fd.getAbsolutePath() };

			writer.writeNext(data1);

			if (message.getAttachments().size() > 0) {
				new File(destination_path + File.separator + path + File.separator + "Attachment" + File.separator
						+ subname).mkdirs();

			}

			for (int j = 0; j < message.getAttachments().size(); j++) {
				Attachment att = (Attachment) message.getAttachments().get_Item(j);

				String s = getFileExtension(att.getName());
				String attFileName = getRidOfIllegalFileNameCharacters(att.getName().replace("." + s, ""));

				att.save(destination_path + File.separator + path + File.separator + "Attachment" + File.separator
						+ subname + File.separator + attFileName + "." + s);

			}

			count_destination++;

		} catch (Error e) {
			logger.warning("ERROR : " + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());
		}

		catch (Exception e) {
			logger.warning(
					"Exception : " + e.getMessage() + "Message" + " " + message.getDate() + System.lineSeparator());

		}

	}

	static String getFileExtension(String fileName) {

		if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0)
			return fileName.substring(fileName.lastIndexOf(".") + 1);
		else
			return "";
	}

	public ImapClient connectiontogmail_input() throws Exception {
		clientforimap_input = new ImapClient("imap.gmail.com", 993, username_p2, password_p2);

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_input.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_input = clientforimap_input.createConnection();
		return clientforimap_input;
	}

	public ImapClient connectiontonYandex_input() throws Exception {
		clientforimap_input = new ImapClient("imap.yandex.com", 993, username_p2, password_p2);

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_input.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_input = clientforimap_input.createConnection();
		return clientforimap_input;
	}

	public ImapClient connectiontoyahoo_input() throws Exception {
		clientforimap_input = new ImapClient("imap.mail.yahoo.com", 993, username_p2, password_p2);

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_input.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_input = clientforimap_input.createConnection();

		return clientforimap_input;
	}

	public ImapClient connectiontoaol_input() throws Exception {
		clientforimap_input = new ImapClient("imap.aol.com", 993, username_p2, password_p2);

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_input.setTimeout(5 * 60 * 1000);
		clientforimap_input.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_input = clientforimap_input.createConnection();
		return clientforimap_input;
	}

	public ImapClient connectiontozoho_input() throws Exception {
		try {
			clientforimap_input = new ImapClient("imap.zoho.com", 993, username_p2, password_p2);

			clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

			EmailClient.setSocketsLayerVersion2(true);
			EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
			clientforimap_input.setTimeout(2 * 60 * 1000);
			iconnforimap_input = clientforimap_input.createConnection();
		} catch (Exception e) {
			clientforimap_input = new ImapClient("imappro.zoho.in", 993, username_p2, password_p2);

			clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

			EmailClient.setSocketsLayerVersion2(true);
			EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
			clientforimap_input.setTimeout(2 * 60 * 1000);
			iconnforimap_input = clientforimap_input.createConnection();
		}
		return clientforimap_input;
	}

	public ImapClient connectiontoimap_input() throws Exception {
		clientforimap_input = new ImapClient(domain_p2, portmo, username_p2, password_p2);

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_input.setTimeout(5 * 60 * 1000);
		iconnforimap_input = clientforimap_input.createConnection();
		return clientforimap_input;
	}

	public ImapClient connectiontoGoDaddy_input() throws Exception {
		clientforimap_input = new ImapClient("imap.secureserver.net", 993, username_p2, password_p2);

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_input.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_input = clientforimap_input.createConnection();
		return clientforimap_input;
	}

	public ImapClient connectiontoHostgator_input() throws Exception {
		clientforimap_input = new ImapClient(domain_p2, portmo, username_p2, password_p2);

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_input.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_input = clientforimap_input.createConnection();
		return clientforimap_input;
	}

	public ImapClient connectiontoicloud_input() throws Exception {
		clientforimap_input = new ImapClient("imap.mail.me.com", 993, username_p2, password_p2);

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		clientforimap_input.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_input.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_input = clientforimap_input.createConnection();
		return clientforimap_input;
	}

	public ImapClient connectiontoinaws_input() throws Exception {
		clientforimap_input = new ImapClient(domain_p2, portmo, username_p2, password_p2);

		clientforimap_input.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		clientforimap_input.setTimeout(5 * 60 * 1000);
		iconnforimap_input = clientforimap_input.createConnection();
		return clientforimap_input;
	}

	public ImapClient connectiontoinaws_output() throws Exception {
		clientforimap_output = new ImapClient(domain_p3, portnofiletype, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		clientforimap_output.setTimeout(5 * 60 * 1000);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontoimap_output() throws Exception {
		clientforimap_output = new ImapClient(domain_p3, portnofiletype, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		clientforimap_output.setTimeout(5 * 60 * 1000);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public IEWSClient conntiontooffice365_input() throws Exception {
		clientforexchange_input = EWSClient.getEWSClient(mailboxUri, username_p2, password_p2);
		clientforexchange_input.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		return clientforexchange_input;
	}

	public IEWSClient conntiontohotmail_input() throws Exception {
		clientforexchange_input = EWSClient.getEWSClient("https://outlook.live.com/EWS/Exchange.asmx", username_p2,
				password_p2);
		clientforexchange_input.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		return clientforexchange_input;
	}

	public IEWSClient connectionwithexchangeserver_input() throws Exception {

		clientforexchange_input = EWSClient.getEWSClient("https://" + domain_p2 + "/ews/Exchange.asmx", username_p2,
				password_p2);

		clientforexchange_input.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforexchange_input;

	}

	public IEWSClient connectionwithexchangeserver_output() throws Exception {

		clientforexchange_output = EWSClient.getEWSClient("https://" + domain_p3 + "/ews/Exchange.asmx", username_p3,
				password_p3);
		clientforexchange_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforexchange_output;

	}

	public ImapClient connectiontogmail_output() throws Exception {
		clientforimap_output = new ImapClient("imap.gmail.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		clientforimap_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);

		iconnforimap_output = clientforimap_output.createConnection();

		return clientforimap_output;
	}

	public ImapClient connectiontoGoDaddy_output() throws Exception {
		clientforimap_output = new ImapClient("imap.secureserver.net", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		clientforimap_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontoHostgator_output() throws Exception {
		clientforimap_output = new ImapClient(domain_p3, portnofiletype, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		clientforimap_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontoicloud_output() throws Exception {
		clientforimap_output = new ImapClient("imap.mail.me.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		clientforimap_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontoYandex_output() throws Exception {
		clientforimap_output = new ImapClient("imap.yandex.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		clientforimap_output.setTimeout(5 * 60 * 1000);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontozoho_output() throws Exception {
		try {
			clientforimap_output = new ImapClient("imap.zoho.com", 993, username_p3, password_p3);

			clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

			clientforimap_output.setTimeout(2 * 60 * 1000);

			EmailClient.setSocketsLayerVersion2(true);
			clientforimap_output.setConnectionCheckupPeriod(50000);
			EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
			iconnforimap_output = clientforimap_output.createConnection();
		} catch (Exception e) {
			clientforimap_output = new ImapClient("imappro.zoho.in", 993, username_p3, password_p3);

			clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

			clientforimap_output.setTimeout(2 * 60 * 1000);

			EmailClient.setSocketsLayerVersion2(true);
			clientforimap_output.setConnectionCheckupPeriod(50000);
			EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
			iconnforimap_output = clientforimap_output.createConnection();
		}
		return clientforimap_output;
	}

	public ImapClient connectiontoaol_output() throws Exception {
		clientforimap_output = new ImapClient("imap.aol.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);

		clientforimap_output.setTimeout(5 * 60 * 1000);
		clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		iconnforimap_output = clientforimap_output.createConnection();
		return clientforimap_output;
	}

	public ImapClient connectiontoyahoo_output() throws Exception {
		clientforimap_output = new ImapClient("imap.mail.yahoo.com", 993, username_p3, password_p3);

		clientforimap_output.setSecurityOptions(SecurityOptions.Auto);

		EmailClient.setSocketsLayerVersion2(true);
		clientforimap_output.setTimeout(5 * 60 * 1000);

		// clientforimap_output.setConnectionCheckupPeriod(50000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);
		System.out.println("trying to re-connect");
		iconnforimap_output = clientforimap_output.createConnection();
		System.out.println("connection done");
		return clientforimap_output;
	}

	public IEWSClient conntiontooffice365_output() throws Exception {
		clientforexchange_output = EWSClient.getEWSClient(mailboxUri, username_p3, password_p3);
		EmailClient.setSocketsLayerVersion2(true);

		clientforexchange_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforexchange_output;
	}

	public IEWSClient conntiontohotmail_output() throws Exception {
		clientforexchange_output = EWSClient.getEWSClient("https://outlook.live.com/EWS/Exchange.asmx", username_p3,
				password_p3);

		EmailClient.setSocketsLayerVersion2(true);

		clientforexchange_output.setTimeout(5 * 60 * 1000);
		EmailClient.setSocketsLayerVersion2DisableSSLCertificateValidation(true);

		return clientforexchange_output;
	}

	public static String getRidOfIllegalFileNameCharacters(String strName) {

		String strLegalName = strName.replace(":", " ").replace("\\", "").replace("?", "").replace("/", "")
				.replace("|", "").replace("*", "").replace("<", "").replace(">", "").replace("\t", "")
				.replace("//s", "").replace("\"", "");
		if (strLegalName.length() >= 100) {
			strLegalName = strLegalName.substring(0, 100);
		}

		return strLegalName;

//		String strLegalName = strName.replace(":", " ").replace("\\", "").replace("?", "").replace("/", "")
//				.replace("|", "").replace("*", "").replace("<", "").replace(">", "").replace("\t", "")
//				.replace("//s", "").replace("\"", "").replace(".", "").replace("", "_");
//
//		if (strLegalName.length() >= 100) {
//			strLegalName = strLegalName.substring(0, 100);
//		}
//
//		return strLegalName;
	}

	void openBrowser(String url) {
		if (Desktop.isDesktopSupported()) {
			Desktop desktop = Desktop.getDesktop();
			try {
				desktop.browse(new URI(url));
			} catch (IOException | URISyntaxException e) {
				logger.warning("Warning : " + e.getMessage());
			}
		} else {
			Runtime runtime = Runtime.getRuntime();
			try {
				runtime.exec("xdg-open " + url);
			} catch (IOException e) {
				logger.warning("Warning : " + e.getMessage());
			}
		}
	}

	private String filepath(File file) {
		String fileName = file.getAbsolutePath();
		String filepath = fileName.replace(file.getName(), "");
		return filepath;
	}

	public Logger logFile() {
		Logger logger = Logger.getLogger(messageboxtitle + ".log");
		FileHandler fh;
		try {
			Calendar cal = Calendar.getInstance();
			String d = getRidOfIllegalFileNameCharacters(cal.getTime().toString());
			fh = new FileHandler(textField_hi.getText() + File.separator + messageboxtitle + d + ".log");
			logger.addHandler(fh);
			SimpleFormatter formatter = new SimpleFormatter();
			fh.setFormatter(formatter);
		} catch (SecurityException e) {
			logger.severe(e.getMessage());
		} catch (IOException e) {
			logger.severe(e.getMessage());
		}
		return logger;
	}

	private StringBuilder Duration(long startTime) {
		long elapsedTime = System.currentTimeMillis() - startTime;
		long elapsedSeconds = elapsedTime / 1000;
		long secondsDisplay = elapsedSeconds % 60;
		long elapsedMinutes = elapsedSeconds / 60;
		StringBuilder br = new StringBuilder();
		br.append(elapsedMinutes);
		br.append(':');
		br.append(secondsDisplay);
		return br;
	}

	String namingconventionmapi(MapiMessage msg, Date d) {
		String filename = null;
		String frm;
		try {
			frm = msg.getSenderEmailAddress();
		} catch (Exception ep) {
			frm = "";
		}
		if (frm != null) {

		} else {
			frm = "";
		}

		if (frm.length() > 20) {
			frm = frm.substring(0, 20);
		}

		String sub;
		try {
			sub = msg.getSubject();
		} catch (Exception ep) {
			sub = "";
		}

		if (sub != null) {

		} else {
			sub = "";
		}

		if (sub.length() > 40) {
			sub = sub.substring(0, 40);
		}

		String dstr = "";
		String combox_selected = comboBox.getSelectedItem().toString();
		try {

			Calendar cal = Calendar.getInstance();
			cal.setTime(d);
			DecimalFormat formatter = new DecimalFormat("00");

			int date = cal.get(Calendar.DATE);
			String dateformate = formatter.format(date);

			int month = cal.get(Calendar.MONTH);
			month++;
			String monthformate = formatter.format(month);

			int year = cal.get(Calendar.YEAR);
			if (combox_selected.contains("DD-MM-YYYY")) {

				dstr = dateformate + "-" + monthformate + "-" + year;
			} else if (combox_selected.contains("MM-DD-YYYY")) {

				dstr = monthformate + "-" + dateformate + "-" + year;
			} else if (combox_selected.contains("YYYY-MM-DD")) {

				dstr = year + "-" + monthformate + "-" + dateformate;
			} else if (combox_selected.contains("YYYY-DD-MM")) {

				dstr = year + "-" + dateformate + "-" + monthformate;
			}

		} catch (Exception ep) {
			dstr = "";
		}

		if (combox_selected.equalsIgnoreCase("Subject")) {
			filename = sub;
		} else if (combox_selected.contains("Subject_Date")) {
			filename = sub + "_" + dstr;
		} else if (combox_selected.contains("Date_Subject")) {
			filename = dstr + "_" + sub;
		} else if (combox_selected.contains("From_Subject_Date")) {
			filename = frm + "_" + sub + "_" + dstr;
		} else if (combox_selected.contains("Date_From_Subject")) {
			filename = dstr + "_" + frm + "_" + sub;
		}
		filename = getRidOfIllegalFileNameCharacters(filename);
		return filename;
	}

	String namingconventionforcal(MapiMessage msg, MapiCalendar cal) {
		String filename = null;
		String frm;
		try {
			frm = msg.getSenderEmailAddress();
		} catch (Exception ep) {
			frm = "";
		}
		if (frm != null) {

		} else {
			frm = "";
		}

		if (frm.length() > 20) {
			frm = frm.substring(0, 20);
		}

		String sub;
		try {
			sub = msg.getSubject();
		} catch (Exception ep) {
			sub = "";
		}

		if (sub != null) {

		} else {
			sub = "";
		}

		if (sub.length() > 40) {
			sub = sub.substring(0, 40);
		}
		String dstr = "";
		Date d;
		String combox_selected = comboBox.getSelectedItem().toString();
		try {
			d = cal.getStartDate();
			Calendar cal1 = Calendar.getInstance();
			cal1.setTime(d);
			DecimalFormat formatter = new DecimalFormat("00");

			int date = cal1.get(Calendar.DATE);
			String dateformate = formatter.format(date);

			int month = cal1.get(Calendar.MONTH);
			month++;
			String monthformate = formatter.format(month);

			int year = cal1.get(Calendar.YEAR);
			if (combox_selected.contains("DD-MM-YYYY")) {

				dstr = dateformate + "-" + monthformate + "-" + year;
			} else if (combox_selected.contains("MM-DD-YYYY")) {

				dstr = monthformate + "-" + dateformate + "-" + year;
			} else if (combox_selected.contains("YYYY-MM-DD")) {

				dstr = year + "-" + monthformate + "-" + dateformate;
			} else if (combox_selected.contains("YYYY-DD-MM")) {

				dstr = year + "-" + dateformate + "-" + monthformate;
			}

		} catch (Exception ep) {
			dstr = "";
		}

		if (combox_selected.equalsIgnoreCase("Subject")) {
			filename = sub;
		} else if (combox_selected.contains("Subject_Date")) {
			filename = sub + "_" + dstr;
		} else if (combox_selected.contains("Date_Subject")) {
			filename = dstr + "_" + sub;
		} else if (combox_selected.contains("From_Subject_Date")) {
			filename = frm + "_" + sub + "_" + dstr;
		} else if (combox_selected.contains("Date_From_Subject")) {
			filename = dstr + "_" + frm + "_" + sub;
		}
		filename = getRidOfIllegalFileNameCharacters(filename);
		return filename;
	}

	String namingconventionmapi(MapiMessage msg) {
		String filename = null;
		String frm;
		try {
			frm = msg.getSenderEmailAddress();
		} catch (Exception ep) {
			frm = "Na";
		}
		if (frm != null) {

		} else {
			frm = "Na";
		}

		if (frm.length() > 20) {
			frm = frm.substring(0, 20);
		}

		String sub;
		try {
			sub = msg.getSubject();
		} catch (Exception ep) {
			sub = "Na";
		}

		if (sub != null) {

		} else {
			sub = "Na";
		}

		if (sub.length() > 40) {
			sub = sub.substring(0, 40);
		}
		String dstr = "";
		Date d;
		String combox_selected = comboBox.getSelectedItem().toString();
		try {
			d = msg.getDeliveryTime();
			Calendar cal = Calendar.getInstance();
			cal.setTime(d);

			DecimalFormat formatter = new DecimalFormat("00");

			int date = cal.get(Calendar.DATE);
			String dateformate = formatter.format(date);

			int month = cal.get(Calendar.MONTH);
			month++;
			String monthformate = formatter.format(month);

			int year = cal.get(Calendar.YEAR);
			if (combox_selected.contains("DD-MM-YYYY")) {

				dstr = dateformate + "-" + monthformate + "-" + year;
			} else if (combox_selected.contains("MM-DD-YYYY")) {

				dstr = monthformate + "-" + dateformate + "-" + year;
			} else if (combox_selected.contains("YYYY-MM-DD")) {

				dstr = year + "-" + monthformate + "-" + dateformate;
			} else if (combox_selected.contains("YYYY-DD-MM")) {

				dstr = year + "-" + dateformate + "-" + monthformate;
			}

		} catch (Exception ep) {
			dstr = "Na";
		}

		if (combox_selected.equalsIgnoreCase("Subject")) {
			filename = sub;
		} else if (combox_selected.contains("From_Subject_Date")) {
			filename = frm + "_" + sub + "_" + dstr;
		} else if (combox_selected.contains("Subject_Date")) {
			filename = sub + "_" + dstr;
		} else if (combox_selected.contains("Date_Subject")) {
			filename = dstr + "_" + sub;
		} else if (combox_selected.contains("Date_From_Subject")) {
			filename = dstr + "_" + frm + "_" + sub;
		}
		filename = getRidOfIllegalFileNameCharacters(filename);
		return filename;
	}

	String namingconventionmapi(MapiMessage msg, File file) {
		String filename = "";
		String frm;
		try {
			frm = msg.getSenderEmailAddress();
		} catch (Exception ep) {
			frm = "";
		}
		if (frm != null) {

		} else {
			frm = "";
		}

		if (frm.length() > 20) {
			frm = frm.substring(0, 20);
		}

		String sub;
		try {
			sub = msg.getSubject();
		} catch (Exception ep) {
			sub = "";
		}

		if (sub != null) {

		} else {
			sub = "";
		}

		if (sub.length() > 80) {
			sub = sub.substring(0, 80);
		}

		String dstr = "";
		Date d;
		String combox_selected = comboBox.getSelectedItem().toString();
		try {
			d = msg.getDeliveryTime();
			Calendar cal = Calendar.getInstance();
			cal.setTime(d);
			DecimalFormat formatter = new DecimalFormat("00");

			int date = cal.get(Calendar.DATE);
			String dateformate = formatter.format(date);

			int month = cal.get(Calendar.MONTH);
			month++;
			String monthformate = formatter.format(month);

			int year = cal.get(Calendar.YEAR);
			if (combox_selected.contains("DD-MM-YYYY")) {

				dstr = dateformate + "-" + monthformate + "-" + year;
			} else if (combox_selected.contains("MM-DD-YYYY")) {

				dstr = monthformate + "-" + dateformate + "-" + year;
			} else if (combox_selected.contains("YYYY-MM-DD")) {

				dstr = year + "-" + monthformate + "-" + dateformate;
			} else if (combox_selected.contains("YYYY-DD-MM")) {

				dstr = year + "-" + dateformate + "-" + monthformate;
			}

		} catch (Exception ep) {
			dstr = "";
		}

		if (combox_selected.contains("From_Subject_Date")) {
			filename = frm + "_" + sub + "_" + dstr;
		} else if (combox_selected.equalsIgnoreCase("Subject")) {
			filename = sub;
		} else if (combox_selected.contains("Subject_Date")) {
			filename = sub + "_" + dstr;
		} else if (combox_selected.contains("Date_Subject")) {
			filename = dstr + "_" + sub;
		} else if (combox_selected.contains("Date_From_Subject")) {
			filename = dstr + "_" + frm + "_" + sub;
		} else if (combox_selected.equalsIgnoreCase("Original File Name")) {
			filename = file.getName().replace(".msg", "").replace(".eml", "").replace(".emlx", "");
		}

		filename = getRidOfIllegalFileNameCharacters(filename);
		return filename;
	}

	public void recursiveDelete(File file) {

		if (!file.exists())
			return;

		if (file.isDirectory()) {
			for (File f : file.listFiles()) {

				recursiveDelete(f);
			}
		}

		if (file.getName().contains(hashcode("userdetails"))) {
			System.out.println(file.getAbsolutePath());
			file.delete();

		}

	}

	public void connectionHandle() {
		lbl_progressreport.setText("INTERNET Connection  LOST ");
		label_12.setIcon(new ImageIcon(Main_Frame.class.getResource("/images.jpg")));

		while (true) {
			System.out.println("Connection not established ");
			try {
				lbl_progressreport.setText("Connecting to Server Please Wait");
				if (filetype.equalsIgnoreCase("OFFICE 365")) {
					conntiontooffice365_output();
				}

				else if (filetype.equalsIgnoreCase("GMAIL")) {

					connectiontogmail_output();

				} else if (filetype.equalsIgnoreCase("Yandex Mail")) {
					connectiontoYandex_output();

				} else if (filetype.equalsIgnoreCase("Hostgator email")) {
					connectiontoHostgator_output();

				} else if (filetype.equalsIgnoreCase("Icloud")) {
					connectiontoicloud_output();

				} else if (filetype.equalsIgnoreCase("GoDaddy email")) {
					connectiontoGoDaddy_output();

				} else if (filetype.equalsIgnoreCase("Live Exchange")) {
					connectionwithexchangeserver_output();

				} else if (filetype.equalsIgnoreCase("IMAP")) {

					connectiontoimap_output();

				} else if (filetype.equalsIgnoreCase("Hotmail")) {
					conntiontohotmail_output();

				} else if (filetype.equalsIgnoreCase("Zoho MAIL")) {

					connectiontozoho_output();

				} else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {

					connectiontoyahoo_output();

				} else if (filetype.equalsIgnoreCase("AOL")) {

					connectiontoaol_output();

				} else if (filetype.equalsIgnoreCase("Amazon WorkMail")) {

					connectiontoinaws_output();

				}

				if (fileoption.equalsIgnoreCase("Yahoo Mail")) {

					connectiontoyahoo_input();

				} else if (fileoption.equalsIgnoreCase("Yandex Mail")) {
					connectiontonYandex_input();

				} else if (fileoption.equalsIgnoreCase("Zoho MAIL")) {

					connectiontozoho_input();

				} else if (fileoption.equalsIgnoreCase("Icloud")) {
					connectiontoicloud_input();

				} else if (fileoption.equalsIgnoreCase("GoDaddy email")) {
					connectiontoGoDaddy_input();

				} else if (fileoption.equalsIgnoreCase("Hostgator email")) {
					connectiontoHostgator_input();

				} else if (fileoption.equalsIgnoreCase("YAHOO MAIL")) {

					connectiontozoho_output();

				} else if (fileoption.equalsIgnoreCase("Gmail")) {

					connectiontogmail_input();

				} else if (fileoption.equalsIgnoreCase("Aol")) {

					connectiontoaol_input();

				} else if (fileoption.equalsIgnoreCase("Amazon WorkMail")) {

					connectiontoinaws_input();

				} else if (fileoption.equalsIgnoreCase("IMAP")) {

					connectiontoimap_input();

				} else if (fileoption.equalsIgnoreCase("Hotmail")) {
					conntiontohotmail_input();

				}

				else if (fileoption.equalsIgnoreCase("OFFICE 365")) {
					conntiontooffice365_input();

				} else if (fileoption.equalsIgnoreCase("Live Exchange")) {
					connectionwithexchangeserver_input();

				}
				label_12.setIcon(new ImageIcon(Main_Frame.class.getResource("/download.png")));
				lbl_progressreport.setText("Connection established Retriving Messasge");
				break;
			} catch (Exception e) {
				lbl_progressreport.setText("INTERNET Connection  LOST ");

			}

		}

		Progressbar.setVisible(true);

	}

	public void connectionHandle1() {

		while (true) {
			try {

				if (filetype.equalsIgnoreCase("OFFICE 365")) {
					conntiontooffice365_output();
				} else if (filetype.equalsIgnoreCase("Amazon WorkMail")) {

					connectiontoinaws_output();

				} else if (filetype.equalsIgnoreCase("GMAIL")) {

					connectiontogmail_output();

				} else if (filetype.equalsIgnoreCase("Hostgator email")) {
					connectiontoHostgator_output();

				} else if (filetype.equalsIgnoreCase("Icloud")) {
					connectiontoicloud_output();

				} else if (filetype.equalsIgnoreCase("GoDaddy email")) {
					connectiontoGoDaddy_output();

				} else if (filetype.equalsIgnoreCase("Live Exchange")) {
					connectionwithexchangeserver_output();

				} else if (filetype.equalsIgnoreCase("IMAP")) {

					connectiontoimap_output();

				} else if (filetype.equalsIgnoreCase("Hotmail")) {
					conntiontohotmail_output();

				} else if (filetype.equalsIgnoreCase("Zoho MAIL")) {

					connectiontozoho_output();

				} else if (filetype.equalsIgnoreCase("YAHOO MAIL")) {

					connectiontoyahoo_output();

				} else if (filetype.equalsIgnoreCase("AOL")) {

					connectiontoaol_output();

				}
				if (fileoption.equalsIgnoreCase("Yahoo Mail")) {

					connectiontoyahoo_input();

				} else if (fileoption.equalsIgnoreCase("Zoho MAIL")) {

					connectiontozoho_input();

				} else if (fileoption.equalsIgnoreCase("Amazon WorkMail")) {

					connectiontoinaws_input();

				} else if (fileoption.equalsIgnoreCase("Icloud")) {
					connectiontoicloud_input();

				} else if (fileoption.equalsIgnoreCase("GoDaddy email")) {
					connectiontoGoDaddy_input();

				} else if (fileoption.equalsIgnoreCase("Hostgator email")) {
					connectiontoHostgator_input();

				} else if (fileoption.equalsIgnoreCase("YAHOO MAIL")) {

					connectiontozoho_output();

				} else if (fileoption.equalsIgnoreCase("Gmail")) {

					connectiontogmail_input();

				} else if (fileoption.equalsIgnoreCase("Aol")) {

					connectiontoaol_input();

				} else if (fileoption.equalsIgnoreCase("IMAP")) {

					connectiontoimap_input();

				} else if (fileoption.equalsIgnoreCase("Hotmail")) {
					conntiontohotmail_input();

				}

				else if (fileoption.equalsIgnoreCase("OFFICE 365")) {
					conntiontooffice365_input();

				} else if (fileoption.equalsIgnoreCase("Live Exchange")) {
					connectionwithexchangeserver_input();

				}

				break;

			} catch (Exception e) {
				e.printStackTrace();
			}
		}

	}

	String calendarname(Appointment ap) {
		String s = "";
		try {
			s = ap.getOrganizer().getDisplayName().replaceAll("[\\[\\]]", "");
		} catch (Exception e) {
			s = "Calendar";
		}
		s = getRidOfIllegalFileNameCharacters(s);
		return s;

	}

	String contactname(Contact ap) {
		String s = "";
		try {
			s = ap.getDisplayName().replaceAll("[\\[\\]]", "");
		} catch (Exception e) {
			s = "Contact";
		}
		s = getRidOfIllegalFileNameCharacters(s);
		return s;

	}

	String contactname(MapiContact ap) {
		String s = "";
		try {
			s = getRidOfIllegalFileNameCharacters(ap.getSubject());
		} catch (Exception e) {
			s = "Contact";
		}
		s = getRidOfIllegalFileNameCharacters(s);
		return s;

	}

	String namingconventionmail(MailMessage msg, File file) {
		String filename = null;
		String frm;
		try {
			frm = msg.getFrom().toString();
		} catch (Exception ep) {
			frm = "";
		}
		if (frm != null) {

		} else {
			frm = "";
		}

		if (frm.length() > 20) {
			frm = frm.substring(0, 20);
		}

		String sub;
		try {
			sub = msg.getSubject();
		} catch (Exception ep) {
			sub = "";
		}

		if (sub != null) {

		} else {
			sub = "";
		}

		if (sub.length() > 40) {
			sub = sub.substring(0, 40);
		}

		String dstr = "";
		Date d;

		String combox_selected = comboBox.getSelectedItem().toString();
		try {
			d = msg.getDate();
			Calendar cal = Calendar.getInstance();
			cal.setTime(d);
			DecimalFormat formatter = new DecimalFormat("00");

			int date = cal.get(Calendar.DATE);
			String dateformate = formatter.format(date);

			int month = cal.get(Calendar.MONTH);
			month++;
			String monthformate = formatter.format(month);

			int year = cal.get(Calendar.YEAR);
			if (combox_selected.contains("DD-MM-YYYY")) {

				dstr = dateformate + "-" + monthformate + "-" + year;
			} else if (combox_selected.contains("MM-DD-YYYY")) {

				dstr = monthformate + "-" + dateformate + "-" + year;
			} else if (combox_selected.contains("YYYY-MM-DD")) {

				dstr = year + "-" + monthformate + "-" + dateformate;
			} else if (combox_selected.contains("YYYY-DD-MM")) {

				dstr = year + "-" + dateformate + "-" + monthformate;
			}
		} catch (Exception ep) {
			dstr = "";
		}

		if (combox_selected.equalsIgnoreCase("Subject")) {
			filename = sub;
		} else if (combox_selected.contains("Subject_Date")) {
			filename = sub + "_" + dstr;
		} else if (combox_selected.contains("Date_Subject")) {
			filename = dstr + "_" + sub;
		} else if (combox_selected.contains("From_Subject_Date")) {
			filename = frm + "_" + sub + "_" + dstr;
		} else if (combox_selected.contains("Date_From_Subject")) {
			filename = dstr + "_" + frm + "_" + sub;
		} else if (combox_selected.equalsIgnoreCase("Original File Name")) {
			filename = file.getName().replace(".msg", "").replace(".eml", "").replace(".emlx", "");
		}

		filename = getRidOfIllegalFileNameCharacters(filename);
		return filename;
	}

	String namingconventionmail(MailMessage msg) {
		String filename = null;
		String frm;
		try {
			frm = getRidOfIllegalFileNameCharacters(msg.getFrom().toString());
		} catch (Exception ep) {
			frm = "";
		}
		if (frm != null) {

		} else {
			frm = "";
		}

		if (frm.length() > 20) {
			frm = frm.substring(0, 20);
		}

		String sub;
		try {
			sub = getRidOfIllegalFileNameCharacters(msg.getSubject());
		} catch (Exception ep) {
			sub = "";
		}

		if (sub != null) {

		} else {
			sub = "";
		}

		if (sub.length() > 40) {
			sub = sub.substring(0, 40);
		}

		String dstr = "";
		Date d;
		String combox_selected = comboBox.getSelectedItem().toString();

		try {
			d = msg.getDate();
			Calendar cal = Calendar.getInstance();
			cal.setTime(d);
			DecimalFormat formatter = new DecimalFormat("00");

			int date = cal.get(Calendar.DATE);
			String dateformate = formatter.format(date);

			int month = cal.get(Calendar.MONTH);
			month++;
			String monthformate = formatter.format(month);

			int year = cal.get(Calendar.YEAR);
			if (combox_selected.contains("DD-MM-YYYY")) {

				dstr = dateformate + "-" + monthformate + "-" + year;
			} else if (combox_selected.contains("MM-DD-YYYY")) {

				dstr = monthformate + "-" + dateformate + "-" + year;
			} else if (combox_selected.contains("YYYY-MM-DD")) {

				dstr = year + "-" + monthformate + "-" + dateformate;
			} else if (combox_selected.contains("YYYY-DD-MM")) {

				dstr = year + "-" + dateformate + "-" + monthformate;
			}

		} catch (Exception ep) {
			dstr = "";
		}

		if (combox_selected.contains("From_Subject_Date")) {
			filename = frm + "_" + sub + "_" + dstr;
		} else if (combox_selected.equalsIgnoreCase("Subject")) {
			filename = sub;
		} else if (combox_selected.contains("Subject_Date")) {
			filename = sub + "_" + dstr;
		} else if (combox_selected.contains("Date_Subject")) {
			filename = dstr + "_" + sub;
		} else if (combox_selected.contains("Date_From_Subject")) {
			filename = dstr + "_" + frm + "_" + sub;
		}

		filename = getRidOfIllegalFileNameCharacters(filename);
		return filename;
	}

	String duplicacymail(MailMessage msg) {
		String frm;
		try {
			frm = getRidOfIllegalFileNameCharacters(msg.getFrom().toString());
		} catch (Exception ep) {
			frm = "";
		}
		if (frm != null) {

		} else {
			frm = "";
		}
		String to;
		try {
			to = msg.getTo().get_Item(0).getAddress();
		} catch (Exception e) {
			to = "";
		}
		if (to != null) {

		} else {
			to = "";
		}
		String body;
		try {
			body = msg.getBody();
		} catch (Exception ep) {
			body = "";
		}

		if (body != null) {
			try {
				body = body.substring(0, 40);
			} catch (Exception e) {

			}

		} else {
			body = "";
		}

		String bcc;
		try {
			bcc = msg.getBcc().get_Item(0).getAddress();
		} catch (Exception ep) {
			bcc = "";
		}

		if (bcc != null) {

		} else {
			bcc = "";
		}

		String sub;
		try {
			sub = getRidOfIllegalFileNameCharacters(msg.getSubject());
		} catch (Exception ep) {
			sub = "";
		}

		if (sub != null) {

		} else {
			sub = "";
		}

		SimpleDateFormat formatter = new SimpleDateFormat("dd-MM-YYYY");
		String dstr;
		Date d;
		try {
			d = msg.getDate();
			dstr = formatter.format(d);
		} catch (Exception ep) {
			dstr = "";
		}

		String input = sub + frm + dstr + to + body + bcc;

		MessageDigest md = null;
		try {
			md = MessageDigest.getInstance("MD5");
		} catch (NoSuchAlgorithmException e) {

			e.printStackTrace();
		}
		byte[] messageDigest = md.digest(input.getBytes());

		BigInteger no = new BigInteger(1, messageDigest);
		String hashtext = no.toString(16);
		while (hashtext.length() < 32) {
			hashtext = "0" + hashtext;
		}
		System.out.println(hashtext);
		return hashtext;

	}

	void mapicsv(MapiMessage message, Date Receiveddate, CSVWriter writer) {

		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = duplicacymapi(message);
			input = input.replaceAll("\\s", "");
			input = input.trim();

			if (!listduplicacy.contains(input)) {
				listduplicacy.add(input);
				if (DateFilter.isSelected()) {
					if (datevalidflag) {
						Mapimess_CSV(message, writer);
						foldermessagecount++;
					}
				} else {
					Mapimess_CSV(message, writer);
					foldermessagecount++;
				}
			}
		} else {
			if (DateFilter.isSelected()) {
				if (datevalidflag) {
					Mapimess_CSV(message, writer);
					foldermessagecount++;
				}
			} else {
				Mapimess_CSV(message, writer);
				foldermessagecount++;
			}
		}

	}

	String mapiexchange(MapiMessage message, Date Receiveddate, String Folderuri) {
		String Messageid = "";
		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = duplicacymapi(message);
			input = input.replaceAll("\\s", "");
			input = input.trim();

			if (!listduplicacy.contains(input)) {
				// System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (chckbx_Mail_Filter.isSelected()) {
					if (Receiveddate.after(mailfilterstartdate) && Receiveddate.before(mailfilterenddate)) {
						Messageid = clientforexchange_output.appendMessage(Folderuri, message, true);
						foldermessagecount++;
						count_destination++;

					} else if (Receiveddate.equals(mailfilterstartdate) || Receiveddate.equals(mailfilterenddate)) {
						Messageid = clientforexchange_output.appendMessage(Folderuri, message, true);
						foldermessagecount++;
						count_destination++;

					}
				} else {

					Messageid = clientforexchange_output.appendMessage(Folderuri, message, true);
					foldermessagecount++;
					count_destination++;
				}
			} else {
				// System.out.println(" duplicate message");
				// System.out.println(input);

			}
		} else {
			if (chckbx_Mail_Filter.isSelected()) {
				if (Receiveddate.after(mailfilterstartdate) && Receiveddate.before(mailfilterenddate)) {
					Messageid = clientforexchange_output.appendMessage(Folderuri, message, true);
					foldermessagecount++;
					count_destination++;

				} else if (Receiveddate.equals(mailfilterstartdate) || Receiveddate.equals(mailfilterenddate)) {

					Messageid = clientforexchange_output.appendMessage(Folderuri, message, true);
					foldermessagecount++;
					count_destination++;
				}
			} else {
				Messageid = clientforexchange_output.appendMessage(Folderuri, message, true);
				foldermessagecount++;
				count_destination++;

			}
		}
		return Messageid;
	}

	void mailmbox(MailMessage message, Date Receiveddate, MboxrdStorageWriter wr) {

		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = duplicacymail(message);
			input = input.replaceAll("\\s", "");
			input = input.trim();

			if (!listduplicacy.contains(input)) {
				listduplicacy.add(input);

				if (DateFilter.isSelected()) {
					if (datevalidflag) {
						wr.writeMessage(message);
						count_destination++;
						foldermessagecount++;
					}
				} else {
					wr.writeMessage(message);
					count_destination++;
					foldermessagecount++;
				}
			}
		} else {
			if (DateFilter.isSelected()) {
				if (datevalidflag) {
					wr.writeMessage(message);
					count_destination++;
					foldermessagecount++;
				}
			} else {
				wr.writeMessage(message);
				count_destination++;
				foldermessagecount++;
			}
		}

	}

	public String removefoldergmail(String path) {
		String[] str1 = path.split("/");

		for (int j = 0; j < str1.length; j++) {

			if (j == 0) {
				path = str1[j];
			} else if (j == str1.length - 1) {
			} else {
				path = path + "/" + str1[j];
			}
		}
		return path;
	}

	void openbrowserturntwostepoff(String openBrowserfor) {
		if (openBrowserfor.equalsIgnoreCase("OFFICE 365")) {
			openBrowser(multifatcorauthicationfor365);
		} else if (openBrowserfor.equalsIgnoreCase("Yandex Mail")) {
			openBrowser(turnofftwostepverificationYandexMail);
		} else if (openBrowserfor.equalsIgnoreCase("Zoho Mail")) {
			openBrowser(turnofftwostepverificationZohoMail);
		} else if (openBrowserfor.equalsIgnoreCase("GMAIL")) {
			openBrowser(turnofftwostepverificationgmail);

		} else if (openBrowserfor.equalsIgnoreCase("Yahoo Mail")) {
			openBrowser(generatethirdpartypassyahoo);
		} else if (openBrowserfor.equalsIgnoreCase("AOL")) {
			openBrowser(createapppasswordforaol);
		} else if (openBrowserfor.equalsIgnoreCase("HOTMAIL")) {
			openBrowser(createnewpasswordforhotmail);
		} else if (openBrowserfor.equalsIgnoreCase("icloud")) {
			openBrowser(All_Data.thirdpartyicloud);
		}
	}

	void openbrowserenableimap(String openBrowserfor) {
		if (openBrowserfor.equalsIgnoreCase("OFFICE 365")) {

		} else if (openBrowserfor.equalsIgnoreCase("IMAP")) {

		} else if (openBrowserfor.equalsIgnoreCase("GMAIL")) {

			openBrowser(enableimapgmail);
		} else if (openBrowserfor.equalsIgnoreCase("Zoho Mail")) {
			openBrowser(turnofftwostepverificationZohoMail);
		} else if (openBrowserfor.equalsIgnoreCase("AOL")) {

		} else if (openBrowserfor.equalsIgnoreCase("HOTMAIL")) {

		} else if (openBrowserfor.equalsIgnoreCase("Live Exchange")) {

		}

	}

	void mailcsv(MailMessage message, Date Receiveddate, CSVWriter writer) {

		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = duplicacymail(message);
			input = input.replaceAll("\\s", "");
			input = input.trim();

			if (!listduplicacy.contains(input)) {

				listduplicacy.add(input);

				if (chckbx_Mail_Filter.isSelected()) {
					if (Receiveddate.after(mailfilterstartdate) && Receiveddate.before(mailfilterenddate)) {
						Mailmess_CSV(message, writer);

					} else if (Receiveddate.equals(mailfilterstartdate) || Receiveddate.equals(mailfilterenddate)) {
						Mailmess_CSV(message, writer);

					}
				} else {
					Mailmess_CSV(message, writer);

				}
			}
		} else {
			if (chckbx_Mail_Filter.isSelected()) {
				if (Receiveddate.after(mailfilterstartdate) && Receiveddate.before(mailfilterenddate)) {
					Mailmess_CSV(message, writer);

				} else if (Receiveddate.equals(mailfilterstartdate) || Receiveddate.equals(mailfilterenddate)) {
					Mailmess_CSV(message, writer);

				}
			} else {
				Mailmess_CSV(message, writer);

			}
		}

	}

	String mailexchange(MailMessage message, Date Receiveddate, String Folderuri) throws Exception {
		String Messageid = "";
		if (chckbxRemoveDuplicacy.isSelected()) {

			String input = duplicacymail(message);
			input = input.replaceAll("\\s", "");
			input = input.trim();

			if (!listduplicacy.contains(input)) {
				// System.out.println("Not a duplicate message");
				listduplicacy.add(input);

				if (DateFilter.isSelected()) {
					if (datevalidflag) {
						Messageid = clientforexchange_output.appendMessage(Folderuri, message);

						count_destination++;

					}
				} else {

					Messageid = clientforexchange_output.appendMessage(Folderuri, message);

					count_destination++;
				}
			}
		} else {
			if (DateFilter.isSelected()) {
				if (datevalidflag) {
					Messageid = clientforexchange_output.appendMessage(Folderuri, message);

					count_destination++;

				}
			} else {
				Messageid = clientforexchange_output.appendMessage(Folderuri, message);

				count_destination++;

			}
		}
		return Messageid;
	}

	String mailimap(MailMessage message, Date Receiveddate, String path) throws Exception {
		String Messageid = "";
		try {
			if (chckbxRemoveDuplicacy.isSelected()) {
				String input = duplicacymail(message);
				input = input.replaceAll("\\s", "");
				input = input.trim();

				if (!listduplicacy.contains(input)) {

					listduplicacy.add(input);

					if (DateFilter.isSelected()) {
						if (datevalidflag) {
							Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
							foldermessagecount++;
							count_destination++;
						}
					} else {
						Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
						foldermessagecount++;
						count_destination++;
					}
				}
			} else {
				if (DateFilter.isSelected()) {
					if (datevalidflag) {
						Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
						foldermessagecount++;
						count_destination++;
					}
				} else {
					System.out.println("---->>start");
					Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
					System.out.println("-->>end");
					count_destination++;
				}
			}
		} catch (Exception e) {

			e.printStackTrace();

			StringWriter sw = new StringWriter();
			e.printStackTrace(new PrintWriter(sw));
			String exceptionAsString = sw.toString();
			if (exceptionAsString.contains("Message too large")) {
				File f = new File(
						(System.getProperty("user.home") + File.separator + "Desktop") + File.separator + calendertime
								+ File.separator + "Attachment" + File.separator + namingconventionmail(message));
				f.mkdirs();
				logger.info("Message size was greater than allowed size so attachment has been deleted and saved in "
						+ f.getAbsolutePath());

				MapiConversionOptions d = MapiConversionOptions.getASCIIFormat();
				MapiMessage message1 = MapiMessage.fromMailMessage(message, d);

				for (MapiAttachment attachment : message1.getAttachments()) {

					attachment.save(f.getAbsolutePath() + File.separator
							+ getRidOfIllegalFileNameCharacters(attachment.getLongFileName()));

				}

				message1.getAttachments().clear();
				message1.getAttachments().removeAll(message1.getAttachments());

				MailConversionOptions d1 = new MailConversionOptions();
				message = message1.toMailMessage(d1);

				Messageid = clientforimap_output.appendMessage(iconnforimap_output, path, message);
				count_destination++;
			}

		}
		return Messageid;
	}

	String duplicacymapiCal(MapiCalendar calender) {

		String checkBoxCalLoc = null;
		String checkBoxCompanycal = null;

		String checkBoxCalSub = null;

		String chckbxStartdate = null;
		String chckbxEnddate = null;
		String checkBoxCalCat = null;
		MessageDigest md = null;

		try {
			checkBoxCalSub = calender.getSubject();
		} catch (Exception ep) {
			checkBoxCalSub = "";
		}

		if (checkBoxCalSub != null) {

		} else {
			checkBoxCalSub = "";
		}

		try {

			chckbxStartdate = calender.getStartDate().toString();
		} catch (Exception ep) {
			chckbxStartdate = "";
		}

		if (chckbxStartdate != null) {

		} else {
			chckbxStartdate = "";
		}

		try {
			chckbxEnddate = calender.getEndDate().toString();
		} catch (Exception ep) {
			chckbxStartdate = "";
		}
		if (chckbxEnddate != null) {

		} else {
			chckbxEnddate = "";
		}

		String input = checkBoxCalSub + chckbxStartdate + chckbxEnddate + checkBoxCalLoc + checkBoxCompanycal
				+ checkBoxCalCat;
		try {
			md = MessageDigest.getInstance("MD5");
		} catch (NoSuchAlgorithmException e) {

			e.printStackTrace();
		}
		byte[] messageDigest = md.digest(input.getBytes());

		// Convert byte array into signum representation
		BigInteger no = new BigInteger(1, messageDigest);
		String hashtext = no.toString(16);
		while (hashtext.length() < 32) {
			hashtext = "0" + hashtext;
		}
		return hashtext;
	}

	int nummberofDigit(int k) {
		int count = 0;
		while (k != 0) {

			k /= 10;
			++count;
		}

		return count;

	}

	String duplicacymapiTask(MapiTask task) {

		String chckbxtaskenddate = null;
		String checkBoxtaskibilling = null;
		String chckbxtaskCategories_1 = null;
		String chckbxtaskCompanies = null;
		String chckbxtaskStatus = null;
		String checkBoxtasksub = null;
		String checkBox_taskbody = null;
		String checkBox_taskstartdate = null;

		MessageDigest md = null;

		try {

			checkBoxtaskibilling = task.getBilling();
		} catch (Exception ep) {
			checkBoxtaskibilling = "";
		}
		if (checkBoxtaskibilling != null) {

		} else {
			checkBoxtaskibilling = "";
		}

		try {

			int status = task.getStatus();

			chckbxtaskStatus = String.valueOf(status);

		} catch (Exception ep) {
			chckbxtaskStatus = "";
		}

		if (chckbxtaskStatus != null) {

		} else {
			chckbxtaskStatus = "";
		}

		try {
			checkBoxtasksub = task.getSubject();
		} catch (Exception ep) {
			checkBoxtasksub = "";
		}

		if (checkBoxtasksub != null) {

		} else {
			checkBoxtasksub = "";
		}

		try {

			checkBox_taskstartdate = task.getStartDate().toString();
		} catch (Exception ep) {
			checkBox_taskstartdate = "";
		}

		if (checkBox_taskstartdate != null) {

		} else {
			checkBox_taskstartdate = "";
		}

		try {
			chckbxtaskenddate = task.getDueDate().toString();
		} catch (Exception ep) {
			chckbxtaskenddate = "";
		}
		if (chckbxtaskenddate != null) {

		} else {
			chckbxtaskenddate = "";
		}

		try {

			String[] st = task.getCategories();
			chckbxtaskCategories_1 = st[0].toString();
		} catch (Exception ep) {
			chckbxtaskCategories_1 = "";
		}
		if (chckbxtaskCategories_1 != null) {

		} else {
			chckbxtaskCategories_1 = "";
		}

		try {

			checkBox_taskbody = task.getBody();
		} catch (Exception ep) {
			checkBox_taskbody = "";
		}
		if (checkBox_taskbody != null) {

		} else {
			checkBox_taskbody = "";
		}

		try {

			String[] st = task.getCategories();
			chckbxtaskCompanies = st[0].toString();
		} catch (Exception ep) {
			chckbxtaskCompanies = "";
		}
		if (chckbxtaskCompanies != null) {

		} else {
			chckbxtaskCompanies = "";
		}

		String input = chckbxtaskCompanies + checkBox_taskbody + chckbxtaskCategories_1 + checkBox_taskstartdate
				+ chckbxtaskenddate + checkBoxtasksub + chckbxtaskStatus + checkBoxtaskibilling;
		try {
			md = MessageDigest.getInstance("MD5");
		} catch (NoSuchAlgorithmException e) {

			e.printStackTrace();
		}
		byte[] messageDigest = md.digest(input.getBytes());

		BigInteger no = new BigInteger(1, messageDigest);
		String hashtext = no.toString(16);
		while (hashtext.length() < 32) {
			hashtext = "0" + hashtext;
		}

		return hashtext;
	}

	public static String hashcode(String input) {
		input = getRidOfIllegalFileNameCharacters(input);
		MessageDigest md = null;
		try {
			md = MessageDigest.getInstance("MD5");
		} catch (NoSuchAlgorithmException e) {

			e.printStackTrace();
		}
		byte[] messageDigest = md.digest(input.getBytes());

		BigInteger no = new BigInteger(1, messageDigest);
		String hashtext = no.toString(16);
		while (hashtext.length() < 32) {
			hashtext = "0" + hashtext;
		}

		return hashtext;
	}

	String duplicacymapiContact(MapiContact Contact) {
		String checkBoxsub = null;
		String chckbxFullName = null;
		String chckbxEmailAddress = null;
		String chckbxMobilenumber = null;
		String chckbxJobtitle = null;
		String chckbxLocation = null;
		String chckbxCompany = null;
		String chckbxCategories = null;
		String chckbxBirthday = null;
		MessageDigest md = null;

		try {
			new MapiContactNamePropertySet();
			chckbxFullName = Contact.getNameInfo().getDisplayName().replaceAll("[\\[\\]]", "");
		} catch (Exception ep) {
			chckbxFullName = "";
		}
		if (chckbxFullName != null) {

		} else {
			chckbxFullName = "";
		}

		try {

			Contact.getElectronicAddresses().getDefaultEmailAddress();
			chckbxEmailAddress = MapiContactElectronicAddress.to_MapiContactElectronicAddress(
					Contact.getElectronicAddresses().getDefaultEmailAddress().toString()).toString();
		} catch (Exception ep) {
			chckbxEmailAddress = "";
		}

		if (chckbxEmailAddress != null) {

		} else {
			chckbxEmailAddress = "";
		}

		try {
			checkBoxsub = Contact.getSubject();
		} catch (Exception ep) {
			checkBoxsub = "";
		}

		if (checkBoxsub != null) {

		} else {
			checkBoxsub = "";
		}

		try {
			new MapiContactTelephonePropertySet();
			chckbxMobilenumber = Contact.getTelephones().getMobileTelephoneNumber().toString();
		} catch (Exception e) {
			chckbxMobilenumber = "";
		}
		if (chckbxMobilenumber != null) {

		} else {
			chckbxMobilenumber = "";
		}

		try {
			new MapiContactProfessionalPropertySet();
			chckbxJobtitle = Contact.getProfessionalInfo().getTitle();
		} catch (Exception ep) {
			chckbxJobtitle = "";
		}

		if (chckbxJobtitle != null) {

		} else {
			chckbxJobtitle = "";
		}

		try {
			chckbxLocation = Contact.getPersonalInfo().getLocation();
		} catch (Exception ep) {
			chckbxLocation = "";
		}
		if (chckbxLocation != null) {

		} else {
			chckbxLocation = "";
		}

		try {

			chckbxJobtitle = Contact.getProfessionalInfo().getCompanyName();
		} catch (Exception ep) {
			chckbxCompany = "";
		}
		if (chckbxCompany != null) {

		} else {
			chckbxCompany = "";
		}

		try {
			chckbxCategories = Contact.getCategories().toString();
		} catch (Exception ep) {
			chckbxCategories = "";
		}
		if (chckbxCategories != null) {

		} else {
			chckbxCategories = "";
		}

		try {
			chckbxBirthday = Contact.getEvents().getBirthday().toString();
		} catch (Exception ep) {
			chckbxBirthday = "";
		}
		if (chckbxBirthday != null) {

		} else {
			chckbxBirthday = "";
		}

		String input = chckbxBirthday + chckbxCategories + chckbxCompany + chckbxLocation + chckbxJobtitle
				+ chckbxJobtitle + chckbxEmailAddress + chckbxFullName + checkBoxsub;
		System.out.println("iNPUT : " + input);
		try {
			md = MessageDigest.getInstance("MD5");
		} catch (NoSuchAlgorithmException e) {

			e.printStackTrace();
		}
		byte[] messageDigest = md.digest(input.getBytes());

		BigInteger no = new BigInteger(1, messageDigest);
		String hashtext = no.toString(16);
		while (hashtext.length() < 32) {
			hashtext = "0" + hashtext;
		}
		return hashtext;
	}

	public String removefolder(String path) {

		String[] str1;
		if (OS.contains("windows")) {
			str1 = path.split("\\\\");
		} else {
			str1 = path.split(File.separator);
		}

		for (int j = 0; j < str1.length; j++) {

			if (j == 0) {
				path = str1[j];
			} else if (!(j == str1.length - 1)) {
				path = path + File.separator + str1[j];
			}
		}
		return path;
	}

	int convertMillisintohour(int Millis) {

		int i = Millis / (1000 * 60 * 60);
		return i;
	}

	int convertMillisintomin(int Millis) {
		int i = (Millis % (1000 * 60 * 60)) / (1000 * 60);
		return i;
	}

	String duplicacymapi(MapiMessage msg) {
		String frm;
		try {
			frm = msg.getSenderEmailAddress();
		} catch (Exception ep) {
			frm = "";
		}
		if (frm != null) {

		} else {
			frm = "";
		}

		String to;
		try {
			to = msg.getDisplayTo();
		} catch (Exception e) {
			to = "";
		}
		if (to != null) {

		} else {
			to = "";
		}
		String sub;
		try {
			sub = msg.getSubject();
		} catch (Exception ep) {
			sub = "";
		}

		if (sub != null) {

		} else {
			sub = "";
		}
		String body;
		try {
			body = msg.getBody();
		} catch (Exception ep) {
			body = "";
		}

		if (body != null) {
			try {
				body = body.substring(0, 40);
			} catch (Exception e) {

			}

		} else {
			body = "";
		}

		String bcc;
		try {
			bcc = msg.getDisplayBcc();
		} catch (Exception ep) {
			bcc = "";
		}

		if (bcc != null) {

		} else {
			bcc = "";
		}

		Date d;
		String dstr;
		try {
			d = msg.getDeliveryTime();
			dstr = d.toString();
		} catch (Exception ep) {
			dstr = "";
		}

//		String input = sub + frm + body + dstr + to + bcc;
		String input = sub.replace(" ", "") + frm.replace(" ", "") + body.replace(" ", "") + to.replace(" ", "")
				+ bcc.replace(" ", "");
		System.out.println("input : " + input);
		MessageDigest md = null;
		try {
			md = MessageDigest.getInstance("MD5");
		} catch (NoSuchAlgorithmException e) {

			e.printStackTrace();
		}
		byte[] messageDigest = md.digest(input.getBytes());

		BigInteger no = new BigInteger(1, messageDigest);
		String hashtext = no.toString(16);
		while (hashtext.length() < 32) {
			hashtext = "0" + hashtext;
		}
		return hashtext;

	}

	private Map<String, ImageIcon> createImageMap_input(DefaultComboBoxModel<String> l1) {
		Map<String, ImageIcon> map = new HashMap<>();
		try {

			for (int i = 0; i < ad.sop.length; i++) {

				map.put(ad.sop[i], new ImageIcon(Main_Frame.class.getResource(ad.sop_img[i])));

			}

		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return map;
	}

	Map<String, ImageIcon> createImageMap_output(DefaultComboBoxModel<String> l1) {
		Map<String, ImageIcon> map = new HashMap<>();
		try {

			for (int i = 0; i < file_sfd.length; i++) {

				map.put(file_sfd[i], new ImageIcon(Main_Frame.class.getResource(filesfd_img[i])));

			}
			for (int i = 0; i < email_sfd.length; i++) {

				map.put(email_sfd[i], new ImageIcon(Main_Frame.class.getResource(emailsfd_img[i])));

			}
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return map;
	}

	public void runNextTimeComponents(Boolean next) {

		label_14.setVisible(next);
		dateChooserNextSchedular.setVisible(next);
		spinner.setVisible(next);
		rdbtnOnce.setVisible(next);
		rdbtnEveryWeek.setVisible(next);
		rdbtnEveryday.setVisible(next);
		rdbtnOnWeekDay.setVisible(next);
		rdbtnOnmonthDay.setVisible(next);
		rdbtnEveryMonth.setVisible(next);

	}

	public JMenuItem MeniItemFileFormat(JMenuItem pstb, String toolname, String toolrefurl) {
		pstb.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent ev) {

				SwingUtilities.invokeLater(new Runnable() {

					public void run() {

						if (toolname != null) {
							All_Data.input_default = toolname;
							comboBox_FiletypeChooser.setSelectedItem(All_Data.input_default);
						} else {
							openBrowser(toolrefurl);
						}

					}
				});

			}
		});
		return pstb;

	}

	class ListRenderer_output extends DefaultListCellRenderer {
		/**
		 *
		 */
		private static final long serialVersionUID = 1L;
		Font font = new Font("helvitica", Font.PLAIN, 15);

		@Override
		public Component getListCellRendererComponent(JList list, Object value, int index, boolean isSelected,
				boolean cellHasFocus) {
			JLabel label = (JLabel) super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
			label.setIcon(imageMap_output.get((String) value));
			label.setHorizontalTextPosition(JLabel.RIGHT);
			label.setFont(font);
			return label;
		}
	}

	class ListRenderer extends DefaultListCellRenderer {

		private static final long serialVersionUID = 1L;
		Font font = new Font("helvitica", Font.PLAIN, 15);

		@Override
		public Component getListCellRendererComponent(JList list, Object value, int index, boolean isSelected,
				boolean cellHasFocus) {
			JLabel label = (JLabel) super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
			label.setIcon(imageMap.get((String) value));
			label.setHorizontalTextPosition(JLabel.RIGHT);
			label.setFont(font);
			return label;
		}
	}
}
