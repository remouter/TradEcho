//javac -cp .:/home/exp.exactpro.com/oleg.legkov/Downloads/TE/poi-3.9-20121203.jar hw1.java
//java -cp .:/home/exp.exactpro.com/oleg.legkov/Corrector_3.0/lib/poi-3.9-20121203.jar Corrector
//javac -cp .:/home/exp.exactpro.com/oleg.legkov/Corrector_3.0/lib/poi-3.9-20121203.jar Corrector.java



//javac -cp .;C:\tmp\Corrector_3.0\lib\poi-3.9-20121203.jar Corrector.java
//java -cp .;C:\tmp\Corrector_3.0\lib\poi-3.9-20121203.jar Corrector

import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.util.*;

import javax.swing.*;
import javax.swing.event.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.xml.parsers.*;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.w3c.dom.*;
import java.time.*;
import java.time.format.*;

public class Corrector {
	//private static final String HOME = "/home/exp.exactpro.com/oleg.legkov/Corrector_3.0/";
	private static final String HOME = "C:\\tmp\\Corrector_3.0\\";
	private static final String INFORMATION = "Information_3.0.xls";	
	
	
	private BitSet bitset;
	private String fontName;
	private short fontSize;
	private static JTextArea textArea = new JTextArea();
	private static JCheckBox checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6, checkBox7, checkBox8, checkBox9, checkBox10,
		checkBox11, checkBox12, checkBox13, checkBox14, checkBox15, checkBox16, checkBox17, checkBox18, checkBox19, checkBox20, checkBox21,
		checkBox22, checkBox23, checkBox24, checkBox25, checkBox26, checkBox27, checkBox28;
	private static ArrayList<String> fileNames = new ArrayList<String>();
	private static JButton startButton;
	private static JComboBox<String> combo, fontNameCombo;
	private static JComboBox<Short> fontSizeCombo;
	private static TreeMap<String, String> matrixUsers, matrixUsersOleg, matrixUsersArtemKh;
	private static final LocalDateTime time = LocalDateTime.now();
	private static final DateTimeFormatter format = DateTimeFormatter.ofPattern("yyyy/MM/dd hh:mm:ss");
	private static String title = "APA Matrices Corrector v1.0 " +  time.format(format) + " [TradeEcho v.3.0.0]";
	private static boolean showFinalDialog = true;


	public Corrector(){
		bitset = new BitSet(40);
		fontName = "Arial";
		fontSize = 10;
		matrixUsersOleg = new TreeMap<String, String>();
		matrixUsersOleg.put("Executer", "QA_III_OL");
		matrixUsersOleg.put("ExecuterFIX", "III_F_O");
		matrixUsersOleg.put("ExecuterLEI", "QA_III_OL_LEI");
		matrixUsersOleg.put("ConterParty", "QA_MMM_OL");
		matrixUsersOleg.put("ContraFIX", "MMM_F_O");
		matrixUsersOleg.put("ConterPartyLEI", "QA_MMM_OL_LEI");

		matrixUsersArtemKh = new TreeMap<String, String>();
		matrixUsersArtemKh.put("Executer", "QA_AKH_1");
		matrixUsersArtemKh.put("ExecuterFIX", "AKH_11FD");
		matrixUsersArtemKh.put("ExecuterLEI", "AKH_LEI_1");
		matrixUsersArtemKh.put("ConterParty", "QA_AKH_2");
		matrixUsersArtemKh.put("ContraFIX", "AKH_21FD");
		matrixUsersArtemKh.put("ConterPartyLEI", "AKH_LEI_2");

		Image icon = Toolkit.getDefaultToolkit().getImage("c:\\Users\\user\\workspace\\TE_11032018\\src\\icon.jpg");

		JFrame frame = new JFrame();
		frame.setTitle(title);
		frame.setSize(800, 600);
		//frame.setResizable(false);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		//frame.setLayout(new GridLayout(1, 2));
		frame.setLayout(new BorderLayout());
		frame.setIconImage(icon);

		JPanel left = new JPanel();
		//left.setLayout(new GridLayout(2, 1));
		left.setLayout(new BorderLayout());
		//left.setPreferredSize(new Dimension(400, 500));

		//textArea.setPreferredSize(new Dimension(100, 900));
		textArea.getDocument().addDocumentListener(new DocumentListener(){
			public void changedUpdate(DocumentEvent arg0){ startButton.setEnabled(false); }
			public void insertUpdate(DocumentEvent arg0){ startButton.setEnabled(false); }
			public void removeUpdate(DocumentEvent arg0){ startButton.setEnabled(false); }
		});

		JScrollPane scrollPane = new JScrollPane(textArea);
		scrollPane.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
		scrollPane.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_AS_NEEDED);
		//scrollPane.setPreferredSize(new Dimension(100, 900));

		left.add(scrollPane, BorderLayout.CENTER);

		JPanel leftButtonsPanel = new JPanel();
		leftButtonsPanel.setLayout(new FlowLayout());

		JButton button = new JButton("Load...");

		button.addActionListener(new ActionListener(){
			public void actionPerformed(ActionEvent ae){
				JFileChooser dialog = new JFileChooser();
				dialog.setCurrentDirectory(new File("."));
				dialog.setMultiSelectionEnabled(true);
				FileNameExtensionFilter filter = new FileNameExtensionFilter("xls", "xls");
				dialog.setFileFilter(filter);
				int flag = dialog.showOpenDialog(new JFrame());
				if(flag == JFileChooser.APPROVE_OPTION){
					File[] files = dialog.getSelectedFiles();
					fileNames = new ArrayList<String>();
					//System.out.println(files.length);
					String lst = "";
					for(File f : files){
						lst += f.getName() + "\n";
						fileNames.add(f.getAbsolutePath());
					}
					textArea.setText(lst);
					startButton.setEnabled(true);
				}
			}
		});

		leftButtonsPanel.add(button);

		startButton = new JButton("Start");
		startButton.setEnabled(false);
		leftButtonsPanel.add(startButton);

		left.add(leftButtonsPanel, BorderLayout.SOUTH);
		frame.add(left, SwingUtilities.CENTER);


		//New Right Panel
		JPanel controlPanel = new JPanel();
		controlPanel.setPreferredSize(new Dimension(400, 800));
		controlPanel.setLayout(new GridLayout(30, 2));

		controlPanel.add(new JLabel("<html><font color=red>If first time correction choose this!"));

		checkBox1 = new JCheckBox("Clear martix");
		checkBox1.setToolTipText("Íîâàÿ âåðñèÿ äëÿ ïðåä î÷èñêè, íå àôôåêòèò öâåòîâóþ ñõåìó");
		controlPanel.add(checkBox1);


		checkBox2 = new JCheckBox("Fix Buy & Sell IDs");
		checkBox2.setToolTipText("Óäàëåíèå Buy&Sell TradeID èç ìàòðèöû");
		checkBox3 = new JCheckBox("Remove Extra FIX Headers");
		checkBox3.setToolTipText("Óäàëåíèå íå íóæíûõ FIX headers èç áëîêà count");
		checkBox4 = new JCheckBox("Remove Unused Variables");
		checkBox4.setToolTipText("Óäàëåíèå íå èñïîëüçóåìûõ ïåðåìåííûõ");
		checkBox5 = new JCheckBox("Remove Empty Rows");
		checkBox5.setToolTipText("Óäàëåíèå ïóñòûõ ñòðîê");
		checkBox26 = new JCheckBox("Remove Headers in Cancellations");
		checkBox26.setToolTipText("Óäàëåíèå Headers äëÿ êàíñåëîâ");

		controlPanel.add(new JLabel("<html><font color=blue>Func with removes"));
		controlPanel.add(checkBox2);
		controlPanel.add(checkBox3);
		controlPanel.add(checkBox4);
		controlPanel.add(checkBox5);
		controlPanel.add(checkBox26);

		checkBox6 = new JCheckBox("Add empty line to the end of case");
		checkBox6.setToolTipText("Äîáàâëåíèå ïóñòûõ ñðîê â êîíåö òåñò êåéçà åñëè íåîáõîäèìî");
		checkBox19 = new JCheckBox("Correct FIX Counts (new)");
		checkBox19.setToolTipText("Êîððåêöèÿ áëîêà count äëÿ FIX ñîîáùåíèé (Íîâàÿ âåðñèÿ)");

		controlPanel.add(new JLabel("<html><font color=blue>Func with addition"));
		controlPanel.add(checkBox19);
		controlPanel.add(checkBox6);

		checkBox22 = new JCheckBox("Correct test numbers");
		checkBox22.setToolTipText("Íóìåðàöèÿ òåñò êåéçîâ");
		checkBox7 = new JCheckBox("Add counts filters");
		checkBox7.setToolTipText("Äîáàâëåíèå ôèëüòðîâ â áëîê count FIX & RTF");

		JPanel fontPanel = new JPanel();
		fontPanel.setPreferredSize(new Dimension(100, 50));

		checkBox8 = new JCheckBox("Correct fonts");
		checkBox8.setToolTipText("Óñòàíîâêà øðèôòà");
		checkBox8.addActionListener(new ActionListener(){
			public void actionPerformed(ActionEvent e){
				fontNameCombo.setEnabled(checkBox8.isSelected());
				fontSizeCombo.setEnabled(checkBox8.isSelected());
			}
		});

		String[] fontNames = { "Arial", "Arial Narrow", "Consolas", "New Times Romans" };
		fontNameCombo = new JComboBox<String>(fontNames);
		fontNameCombo.setSelectedItem("Arial Narrow");
		fontNameCombo.setEnabled(false);

		Short[] fontSizes = { 8, 10, 12, 14, 16, 18, 20, 22 };
		fontSizeCombo = new JComboBox<Short>(fontSizes);
		fontSizeCombo.setSelectedIndex(3);
		fontSizeCombo.setEnabled(false);
		fontPanel.add(checkBox8);
		fontPanel.add(fontNameCombo);
		fontPanel.add(fontSizeCombo);

		checkBox20 = new JCheckBox("Correct Height & width");
		checkBox20.setToolTipText("Âûðàâíèâàíèå ÿ÷ååê ïî âûñîòå è øèðèíå");
		checkBox9 = new JCheckBox("Correct Headers");
		checkBox9.setToolTipText("Êîððåêöèÿ headers â áëîêå count");
		checkBox10 = new JCheckBox("Fix Diff");
		checkBox10.setToolTipText("Êîððåêöèÿ ôóíêöèè CheckMessage");
		checkBox11 = new JCheckBox("Fix Persistense names");
		checkBox11.setToolTipText("Êîððåêöèÿ íàçâàíèé äëÿ ôàéëîâ Persistense");
		checkBox12 = new JCheckBox("Set Line Numbers");
		checkBox12.setToolTipText("Íóìåðàöèÿ ñòðîê");
		checkBox18 = new JCheckBox("Correct SaveMessages");
		checkBox18.setToolTipText("Êîððåêöèÿ áëîêà ñîõðàíèÿ ñîîáùåíé ïîä íîâûå headers");
		checkBox21 = new JCheckBox("Correct Flag Names");
		checkBox21.setToolTipText("Ïðèñâîåíèå ôëàãàì ïîíÿòíûõ íàâçàíèé");
		checkBox23 = new JCheckBox("Correct LEI Names");
		checkBox23.setToolTipText("Êîððåêöèÿ Member LEI");
		checkBox24 = new JCheckBox("Correct noParty Names");
		checkBox24.setToolTipText("Êîððåêöèÿ NoParty");
		checkBox25 = new JCheckBox("Correct Price Conditions");
		checkBox25.setToolTipText("Êîððåêöèÿ No Price Conditions");


		controlPanel.add(new JLabel("<html><font color=blue>General Corrections"));
		controlPanel.add(checkBox22);
		controlPanel.add(checkBox7);

		controlPanel.add(checkBox20);
		controlPanel.add(checkBox20);
		controlPanel.add(checkBox9);
		controlPanel.add(checkBox18);
		controlPanel.add(checkBox10);
		controlPanel.add(checkBox11);
		controlPanel.add(checkBox12);
		controlPanel.add(checkBox21);
		controlPanel.add(checkBox23);
		controlPanel.add(checkBox24);
		controlPanel.add(checkBox25);

		checkBox13 = new JCheckBox("Remove dashes");
		checkBox13.setToolTipText("Óäàëåíèå _[1...x]");
		checkBox14 = new JCheckBox("Correct KnownBug");
		checkBox14.setToolTipText("Çàìåíà KnownBug íà Expected");
		checkBox15 = new JCheckBox("Replace References names");
		checkBox15.setToolTipText("Ïåðåèìåíîâàíèå ññûëîê ñîîáùåíèé");
		checkBox16 = new JCheckBox("Correct Messages Descriptions");
		checkBox16.setToolTipText("Ïåðåèìåíîâàíèå îïèñàíèÿ ñîîáùåíèé");

		controlPanel.add(new JLabel("<html><font color=blue>FIX & RTF Corrections"));
		controlPanel.add(checkBox13);
		controlPanel.add(checkBox14);
		controlPanel.add(checkBox15);
		controlPanel.add(checkBox16);

		checkBox27 = new JCheckBox("Set New Headers from XML");
		checkBox27.setToolTipText("Óñòàíîâêà õåäåðîâ èç XML");
		checkBox28 = new JCheckBox("Set All Values");
		checkBox28.setToolTipText("Óñòàíîâêà âñåõ çíà÷åíèé èç xml");

		checkBox17 = new JCheckBox("Choose All");
		checkBox17.setToolTipText("Âûáðàòü Âñå");

		controlPanel.add(new JLabel("<html><font color=blue>NEW_____________________________________________"));
		controlPanel.add(checkBox27);
		controlPanel.add(checkBox28);
		controlPanel.add(checkBox17);

		checkBox17.addActionListener(new ActionListener(){
			public void actionPerformed(ActionEvent e){
				checkBox1.setSelected(checkBox17.isSelected());
				checkBox2.setSelected(checkBox17.isSelected());
				checkBox3.setSelected(checkBox17.isSelected());
				checkBox4.setSelected(checkBox17.isSelected());
				checkBox5.setSelected(checkBox17.isSelected());
				checkBox6.setSelected(checkBox17.isSelected());
				checkBox7.setSelected(checkBox17.isSelected());
				checkBox8.setSelected(checkBox17.isSelected());
				checkBox9.setSelected(checkBox17.isSelected());
				checkBox10.setSelected(checkBox17.isSelected());
				checkBox11.setSelected(checkBox17.isSelected());
				checkBox12.setSelected(checkBox17.isSelected());
				checkBox13.setSelected(checkBox17.isSelected());
				checkBox14.setSelected(checkBox17.isSelected());
				checkBox15.setSelected(checkBox17.isSelected());
				checkBox16.setSelected(checkBox17.isSelected());
				checkBox18.setSelected(checkBox17.isSelected());
				checkBox19.setSelected(checkBox17.isSelected());
				checkBox20.setSelected(checkBox17.isSelected());
				checkBox21.setSelected(checkBox17.isSelected());
				checkBox22.setSelected(checkBox17.isSelected());
				fontNameCombo.setEnabled(checkBox17.isSelected());
				fontSizeCombo.setEnabled(checkBox17.isSelected());
				checkBox23.setSelected(checkBox17.isSelected());
				checkBox24.setSelected(checkBox17.isSelected());
				checkBox25.setSelected(checkBox17.isSelected());
				checkBox26.setSelected(checkBox17.isSelected());
				checkBox27.setSelected(checkBox17.isSelected());
				checkBox28.setSelected(checkBox17.isSelected());
			}
		});

		String[] users = {"None", "Oleg", "ArtemKh", "ArtemK", "AlekseyG", "AlekseyS", "Andrey", "Aleksandr", "Luydmila", "Natalia", "Leonid"};
		combo = new JComboBox<String>(users);
		combo.setToolTipText("Çàìåíà þçåðîâ â ìàòðèöå");
		controlPanel.add(combo);
		controlPanel.add(fontPanel);

		combo.addActionListener(new ActionListener(){
			public void actionPerformed(ActionEvent e){
				if(combo.getSelectedItem() == "AlekseyS") startButton.setEnabled(false);
				else startButton.setEnabled(true);
			}
		});


		frame.add(controlPanel, BorderLayout.EAST);
		frame.setVisible(true);

		startButton.addActionListener(new ActionListener(){
			public void actionPerformed(ActionEvent e){
				//Îáùàÿ ïðåä î÷èñòêà
				bitset.clear();
				if(checkBox1.isSelected()) bitset.set(1, true); //te.newClearFunc();

				//Óäàëåíèå è äîáàâëåíèå ñòðîê
				//03042018 3, 4, 17, 5, 15, 34, 6, 7, 27, 8 ================> 3, 4, 5, 15, 17, 34, 6, 7, 27, 8

				if(checkBox2.isSelected()) bitset.set(2, true); //te.fixBuySellId2();
				if(checkBox3.isSelected()) bitset.set(3, true); //te.removeFixHeaders();

				if(checkBox4.isSelected()) bitset.set(4, true); //te.fixUnusedVariables();
				if(checkBox13.isSelected()) bitset.set(13, true); //te.removeDashes();
				if(checkBox15.isSelected()) bitset.set(15, true); //te.replaceReferenceNames(); // need remove dashes first
				if(checkBox26.isSelected()) bitset.set(26, true); //te.mergeTCRforUnPublishedandCancel();
				if(checkBox5.isSelected()) bitset.set(5, true); //te.removeEmptyRows();
				if(checkBox19.isSelected()) bitset.set(19, true); //te.correctFixCounts(); // new
				if(checkBox6.isSelected()) bitset.set(6, true); //te.addEmptyLineToTheEnd();

				//Êîððåêöèÿ äàííûõ
				String user = (String)combo.getSelectedItem();
				if(user != "None") bitset.set(8, true); //te.usersRempacement(matrixUsers);
				matrixUsers = matrixUsersOleg;
				if(user.matches("Oleg")) matrixUsers = matrixUsersOleg;
				if(user.matches("ArtemKh")) matrixUsers = matrixUsersArtemKh;

				if(checkBox21.isSelected()) bitset.set(21, true); //te.correctFlagNames();
				if(checkBox23.isSelected()) bitset.set(23, true); //te.leiNamesCorrection();
				if(checkBox24.isSelected()) bitset.set(24, true); //te.noPartyNamesCorrection();
				if(checkBox25.isSelected()) bitset.set(25, true); //te.priceConditionsCorrection();
				if(checkBox14.isSelected()) bitset.set(14, true); //te.replaceKnownBug();
				if(checkBox10.isSelected()) bitset.set(10, true); //te.fixDiff2();

				//Êîððåêöèÿ õåäåðîâ è êàêíòîâ
				if(checkBox9.isSelected()) bitset.set(9, true); //te.correctHeaders();
				if(checkBox7.isSelected()) bitset.set(7, true); //te.addCountFilters08032018();
				if(checkBox18.isSelected()) bitset.set(18, true); //te.fixSaveMessagesPossition();

				//Êðàñîòà
				if(checkBox16.isSelected()) bitset.set(16, true); //te.correctMessageDescription();
				if(checkBox22.isSelected()) bitset.set(22, true); //te.correctTestNumbers();
				if(checkBox11.isSelected()) bitset.set(11, true); //te.fixPersistenceNums();
				if(checkBox12.isSelected()) bitset.set(12, true); //te.lineNumbers();

				if(checkBox8.isSelected()) {
					fontName = (String)fontNameCombo.getSelectedItem();
					fontSize = (Short)fontSizeCombo.getSelectedItem();
					bitset.set(8, true); //te.correctFonts( fontName, fontSize );
				}

				if(checkBox27.isSelected()) bitset.set(27, true); //te.newFIXHeaders();
				if(checkBox28.isSelected()) bitset.set(28, true); //te.setAllValues();

				try{ start(); }catch(Exception ex){}







				/*		try{
							BitSet bitset = new BitSet(40);
							//bitset.set(0, true);
							System.out.println("BITSET: " + bitset);


							//Îáùàÿ ïðåä î÷èñòêà
							if(checkBox1.isSelected()) bitset.set(1, true); //te.newClearFunc();


							//Óäàëåíèå è äîáàâëåíèå ñòðîê
							//03042018 3, 4, 17, 5, 15, 34, 6, 7, 27, 8 ================> 3, 4, 5, 15, 17, 34, 6, 7, 27, 8
							//

							if(checkBox2.isSelected()) bitset.set(2, true); //te.fixBuySellId2();
							if(checkBox3.isSelected()) bitset.set(3, true); //te.removeFixHeaders();


							if(checkBox4.isSelected()) bitset.set(4, true); //te.fixUnusedVariables();
							if(checkBox13.isSelected()) bitset.set(13, true); //te.removeDashes();
							if(checkBox15.isSelected()) bitset.set(15, true); //te.replaceReferenceNames(); // need remove dashes first
							if(checkBox26.isSelected()) bitset.set(26, true); //te.mergeTCRforUnPublishedandCancel();
							if(checkBox5.isSelected()) bitset.set(5, true); //te.removeEmptyRows();
							if(checkBox19.isSelected()) bitset.set(19, true); //te.correctFixCounts(); // new
							if(checkBox6.isSelected()) bitset.set(6, true); //te.addEmptyLineToTheEnd();

							//Êîððåêöèÿ äàííûõ
							if(user != "None") bitset.set(8, true); //te.usersRempacement(matrixUsers);

							if(checkBox21.isSelected()) bitset.set(21, true); //te.correctFlagNames();
							if(checkBox23.isSelected()) bitset.set(23, true); //te.leiNamesCorrection();
							if(checkBox24.isSelected()) bitset.set(24, true); //te.noPartyNamesCorrection();
							if(checkBox25.isSelected()) bitset.set(25, true); //te.priceConditionsCorrection();
							if(checkBox14.isSelected()) bitset.set(14, true); //te.replaceKnownBug();
							if(checkBox10.isSelected()) bitset.set(10, true); //te.fixDiff2();


							//Êîððåêöèÿ õåäåðîâ è êàêíòîâ
							if(checkBox9.isSelected()) bitset.set(9, true); //te.correctHeaders();
							if(checkBox7.isSelected()) bitset.set(7, true); //te.addCountFilters08032018();
							if(checkBox18.isSelected()) bitset.set(18, true); //te.fixSaveMessagesPossition();


							//Êðàñîòà
							if(checkBox16.isSelected()) bitset.set(16, true); //te.correctMessageDescription();
							if(checkBox22.isSelected()) bitset.set(22, true); //te.correctTestNumbers();
							if(checkBox11.isSelected()) bitset.set(11, true); //te.fixPersistenceNums();
							if(checkBox12.isSelected()) bitset.set(12, true); //te.lineNumbers();

							String fontName = "Arial";
							Short fontSize = 10;
							if(checkBox8.isSelected()) {
								fontName = (String)fontNameCombo.getSelectedItem();
								fontSize = (Short)fontSizeCombo.getSelectedItem();
								bitset.set(8, true); //te.correctFonts( fontName, fontSize );
							}


							if(checkBox27.isSelected()) bitset.set(27, true); //te.newFIXHeaders();
							if(checkBox28.isSelected()) bitset.set(28, true); //te.setAllValues();

							TradeEcho te = new TradeEcho(file);
							te.start(bitset, matrixUsers, fontName, fontSize);
							System.out.println("After Start");

							//te.setString(file);

							//Îáùàÿ ïðåä î÷èñòêà
//							if(checkBox1.isSelected()) te.newClearFunc();
//							te.addRFTMissingColors();
//
//							//Óäàëåíèå è äîáàâëåíèå ñòðîê
//							//03042018 3, 4, 17, 5, 15, 34, 6, 7, 27, 8 ================> 3, 4, 5, 15, 17, 34, 6, 7, 27, 8
//							//
//
//							if(checkBox2.isSelected()) te.fixBuySellId2();
//							if(checkBox3.isSelected()) te.removeFixHeaders();
//
//
//							if(checkBox4.isSelected()) te.fixUnusedVariables();
//							if(checkBox13.isSelected()) te.removeDashes();
//							if(checkBox15.isSelected()) te.replaceReferenceNames(); // need remove dashes first
//							if(checkBox26.isSelected()) te.mergeTCRforUnPublishedandCancel();
//							if(checkBox5.isSelected()) te.removeEmptyRows();
//							if(checkBox19.isSelected()) te.correctFixCounts(); // new
//							if(checkBox6.isSelected()) te.addEmptyLineToTheEnd();
//
//							//Êîððåêöèÿ äàííûõ
//							if(user != "None") te.usersRempacement(matrixUsers);
//
//							if(checkBox21.isSelected()) te.correctFlagNames();
//							if(checkBox23.isSelected()) te.leiNamesCorrection();
//							if(checkBox24.isSelected()) te.noPartyNamesCorrection();
//							if(checkBox25.isSelected()) te.priceConditionsCorrection();
//							if(checkBox14.isSelected()) te.replaceKnownBug();
//							if(checkBox10.isSelected()) te.fixDiff2();
//
//
//							//Êîððåêöèÿ õåäåðîâ è êàêíòîâ
//							if(checkBox9.isSelected()) te.correctHeaders();
//							if(checkBox7.isSelected()) te.addCountFilters08032018();
//							if(checkBox18.isSelected()) te.fixSaveMessagesPossition();
//
//
//							//Êðàñîòà
//							if(checkBox16.isSelected()) te.correctMessageDescription();
//							if(checkBox22.isSelected()) te.correctTestNumbers();
//							if(checkBox11.isSelected()) te.fixPersistenceNums();
//							if(checkBox12.isSelected()) te.lineNumbers();
//							if(checkBox8.isSelected()) {
//								String fontName = (String)fontNameCombo.getSelectedItem();
//								Short fontSize = (Short)fontSizeCombo.getSelectedItem();
//								te.correctFonts( fontName, fontSize );
//							}
//
//
//							if(checkBox27.isSelected()) te.newFIXHeaders();
//							if(checkBox28.isSelected()) te.setAllValues();
//
//							te.closeAll();

						}catch(FileNotFoundException fnfe){
							JOptionPane.showMessageDialog(new JFrame(), "<html>Output file is opened in another program. "
									+ "<br>Please close it to continue", "Warning", JOptionPane.WARNING_MESSAGE);
							showFinalDialog = false;
						}
						catch(Exception ex){
							StringWriter sw = new StringWriter();
							PrintWriter pw = new PrintWriter(sw);
							ex.printStackTrace(pw);
							JOptionPane.showMessageDialog(new JFrame(), sw.toString(), "Error", JOptionPane.ERROR_MESSAGE);
						}*/

					System.out.println("Done!");
					if(showFinalDialog) JOptionPane.showMessageDialog(new JFrame(), "Complete");
				}
			});
		}

	public void start(){
		showFinalDialog = true;
		try{
			for(String file : fileNames){
				if(file == null){ System.out.println("Nothing to execute"); return; }
				new TradeEcho(file).start(bitset, matrixUsers, fontName, fontSize);
				//System.out.println("After Start");
			}
		}catch(FileNotFoundException fnfe){
			JOptionPane.showMessageDialog(new JFrame(), fnfe.toString(), "Warning", JOptionPane.WARNING_MESSAGE);
			showFinalDialog = false;
		}
		catch(Exception ex){
			StringWriter sw = new StringWriter();
			PrintWriter pw = new PrintWriter(sw);
			ex.printStackTrace(pw);
			JOptionPane.showMessageDialog(new JFrame(), sw.toString(), "Error", JOptionPane.ERROR_MESSAGE);
		}
	}

	public static void main(String[] args) throws Exception{
		SwingUtilities.invokeLater(new Runnable(){
			public void run(){
				new Corrector();
			}
		});
	}
}


class TradeEcho implements Runnable{
	//private static final String HOME = "/home/exp.exactpro.com/oleg.legkov/Corrector_3.0/";
	private static final String HOME = "C:\\tmp\\Corrector_3.0\\";
	private static final String INFORMATION = "Information_3.0.xls";
	private static final String INSTRUMENTS = "Instruments_3.0.xls";
	private static final String XMLDB = "MessagesDB_3.0.xml";
	
	
	
	private static final int rtfTagsNumber = 125;
	private File infile;
	private File outFile;
	private FileInputStream in;
	private FileOutputStream out;
	private FileWriter log;
	private String inDocName, outDocName;
	private HSSFWorkbook doc;
	private Set<Integer> caseStartEnd;
	private Map<String, String> mVariableNames;
	private Set<Integer> rowsToDelete;
	private Thread thrd;

	public TradeEcho(String inDocName) throws Exception{
		thrd = new Thread(inDocName);
		thrd.start();
		caseStartEnd = new TreeSet<Integer>();
		mVariableNames = new TreeMap<String, String>();
		rowsToDelete = new TreeSet<Integer>();

		this.inDocName = inDocName;
		outDocName = inDocName.substring(0, inDocName.length() - 4) + "_output.xls";
		infile = new File(this.inDocName);
		outFile = new File(outDocName);
		in = new FileInputStream(infile);
		out = new FileOutputStream(outFile);
		log = new FileWriter("logs.txt");
		doc = new HSSFWorkbook(in);
	}

	public void start(BitSet b, TreeMap<String, String> users, String fontName, Short fontSize) throws Exception{
		//System.out.println("Thread: " + thrd.getName());

		//usersRempacement(users);
		addRFTMissingColors();

		if(b.get(8)) correctFonts( fontName, fontSize );

		if(b.get(1)) newClearFunc();

		if(b.get(2)) fixBuySellId2();
		if(b.get(3)) removeFixHeaders();

		if(b.get(4)) fixUnusedVariables();
		if(b.get(13)) removeDashes();
		if(b.get(15)) replaceReferenceNames();
		if(b.get(26)) mergeTCRforUnPublishedandCancel();
		if(b.get(5)) removeEmptyRows();
		if(b.get(19)) correctFixCounts();
		if(b.get(6)) addEmptyLineToTheEnd();

		if(b.get(21)) correctFlagNames();
		if(b.get(23)) leiNamesCorrection();
		if(b.get(24)) noPartyNamesCorrection();
		if(b.get(25)) priceConditionsCorrection();
		if(b.get(14)) replaceKnownBug();
		if(b.get(10)) fixDiff2();

		if(b.get(9)) correctHeaders();
		if(b.get(7)) addCountFilters08032018();
		if(b.get(18)) fixSaveMessagesPossition();

		if(b.get(16)) correctMessageDescription();
		if(b.get(22)) correctTestNumbers();
		if(b.get(11)) fixPersistenceNums();
		if(b.get(12)) lineNumbers();

		if(b.get(27)) newFIXHeaders();
		if(b.get(28)) setAllValues();

		closeAll();
	}

	public void run(){	}

	public void addRFTMissingColors() throws IOException{
		//System.out.println("addRFTMissingColors");
		log.write("\naddRFTMissingColors Func\n");
		CharSequence seq1 = "dss", seq2 = "gtp";

		HSSFSheet sheet = doc.getSheetAt(0);

		int lastRow = sheet.getLastRowNum();
		for(int i = 0; i < lastRow; i++){
			Row row = sheet.getRow(i);
			if(row == null) continue;

			Cell cell6 = row.getCell(6);
			if(cell6 == null || cell6.getCellType() != Cell.CELL_TYPE_STRING ||
					!(cell6.getStringCellValue().contains(seq1) || cell6.getStringCellValue().contains(seq2))) continue;

			Cell cell8 = row.getCell(8);
			if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING || !cell8.getStringCellValue().matches("receive")) continue;

			CellStyle style = row.getCell(0).getCellStyle();
			for(int j = 1; j < row.getLastCellNum(); j++){
				Cell tempCell = row.getCell(j);
				if(tempCell == null) row.createCell(j, Cell.CELL_TYPE_BLANK);
				row.getCell(j).setCellStyle(style);
			}
		}
	}

	public void setAllValues() throws Exception{
		//System.out.println("setAllValues");
		log.write("\nsetAllValues Func\n");

		//LOAD HEADERS FOR COUNT AND NAMES
		//FileInputStream information = new FileInputStream(HOME + INFORMATION); // WINDOWS
		FileInputStream information = new FileInputStream(HOME + INFORMATION); // LINUX
		HSSFWorkbook book = new HSSFWorkbook(information);
		HSSFSheet infoSheet = book.getSheet("messages2");
		int infoSheetLastRow = infoSheet.getLastRowNum();

		LinkedHashMap<Integer, String> tcrcHeader = new LinkedHashMap<Integer, String>();
		LinkedHashMap<Integer, String> tcraHeader = new LinkedHashMap<Integer, String>();
		LinkedHashMap<Integer, String> tcrsHeader = new LinkedHashMap<Integer, String>();


		for(int i = 0; i < infoSheetLastRow; i++){
			Row row = infoSheet.getRow(i);
			switch(row.getCell(0).getStringCellValue()){
				case("TCR-C"):
					int tcrcCount = 14;
					for(int j = 1; j < row.getLastCellNum(); j++)
						if(row.getCell(j).getCellType() != Cell.CELL_TYPE_BLANK){
							String value = row.getCell(j).getStringCellValue();
							if(value.matches("SecurityIDSource")) value = "Instrument";
							if(value.matches("SecurityID")) continue;
							if(value.matches("CountryOfIssue")) continue;
							tcrcHeader.put(tcrcCount++, value);
						}
					break;
				case("TCR-Ack"):
					int tcraCount = 14;
					for(int j = 1; j < row.getLastCellNum(); j++)
						if(row.getCell(j).getCellType() != Cell.CELL_TYPE_BLANK){
							String value = row.getCell(j).getStringCellValue();
							if(value.matches("SecurityIDSource")) value = "Instrument";
							if(value.matches("SecurityID")) continue;
							if(value.matches("CountryOfIssue")) continue;
							tcraHeader.put(tcraCount++, value);
						}
					break;
				case("TCR-S"):
					int tcsCount = 14;
					for(int j = 1; j < row.getLastCellNum(); j++)
						if(row.getCell(j).getCellType() != Cell.CELL_TYPE_BLANK){
							String value = row.getCell(j).getStringCellValue();
							if(value.matches("SecurityIDSource")) value = "Instrument";
							if(value.matches("SecurityID")) continue;
							if(value.matches("CountryOfIssue")) continue;
							tcrsHeader.put(tcsCount++, value);
						}
					break;
				default: break;
			}
		}
		//System.out.println("tcrcHeader: " + tcrcHeader.toString());
		//System.out.println("tcraHeader: " + tcraHeader.toString());
		//System.out.println("tcrsHeader: " + tcrsHeader.toString());


		//LOAD INSTRUMENTS
		LinkedHashMap<Integer, LinkedHashMap<String, String>> instMap = new LinkedHashMap<Integer, LinkedHashMap<String, String>>();
		//FileInputStream instrumentsStream = new FileInputStream(HOME + INSTRUMENTS); //WINDOWS
		//System.out.println("Ater");
		FileInputStream instrumentsStream = new FileInputStream(HOME + INSTRUMENTS); //LINUX
		HSSFWorkbook instDoc = new HSSFWorkbook(instrumentsStream);
		HSSFSheet instSheet = instDoc.getSheetAt(0);
		int instSheetLastRow = instSheet.getLastRowNum();

		for(int i = 1; i <= instSheetLastRow; i++){
			LinkedHashMap<String, String> tempMap = new LinkedHashMap<String, String>();
			Row headerRow = instSheet.getRow(0);
			Row instRow = instSheet.getRow(i);

			int headerRowLastCell = headerRow.getLastCellNum();

			for(int j = 0; j < headerRowLastCell; j++){
				Cell headerCell = headerRow.getCell(j);
				Cell instCell = instRow.getCell(j);

				if(instCell == null){ tempMap.put(headerCell.getStringCellValue(), ""); continue; }

				String headerCellvalue = headerCell.getStringCellValue();
				String instCellvalue = "";

				switch(instCell.getCellType()){
					case(Cell.CELL_TYPE_BLANK): break;
					case(Cell.CELL_TYPE_NUMERIC):
						double d = instCell.getNumericCellValue();
						if(d % 1 == 0){
							Integer tmp = (int)instCell.getNumericCellValue();
							instCellvalue = Integer.toString(tmp);
						} else instCellvalue = Double.toString(instCell.getNumericCellValue());
						break;
					case(Cell.CELL_TYPE_STRING): instCellvalue = instCell.getStringCellValue(); break;
				}
				tempMap.put(headerCellvalue, instCellvalue);
			}
			//System.out.println("Here: " + instRow.getCell(1));
			instMap.put((int)instRow.getCell(1).getNumericCellValue(), tempMap);
		}

		HSSFSheet sheet = doc.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();
		boolean isReject = false, isPersistence = false;
		boolean isBuy = false, isNewFirmTradeID = false;
		String persistenceName = "";

		String[] groupsArr = { "NoParty1", "Side1" };

		TreeSet<String> groups = new TreeSet<String>(Arrays.asList(groupsArr));

		String[] fixTCRCArr = { "si", "su", "ssd", "smd", "ai", "au", "asdb", "amdb", "asda", "amda" };

		TreeSet<String> fixTCRC = new TreeSet<String>(Arrays.asList(fixTCRCArr));

		String[] amendTCRArr = { "ai", "au", "asdb", "amdb", "asda", "amda" };

		TreeSet<String> amendTCRC = new TreeSet<String>(Arrays.asList(amendTCRArr));
		LinkedHashMap<String, String> originalTCR = new LinkedHashMap<String, String>();
		LinkedHashMap<String, String> amendTCR = null;


		String[] fixTCRAArr = { "si_a", "su_a", "ssd_a", "smd_a", "ci_a", "cu_a", "csdb_a", "cmdb_a", "csda_a", "cmda_a", "pri_a",
		                      "pru_a", "prsda_a", "prmda_a", "prsd_a", "prmd_a", "ai_a", "au_a", "asdb_a", "amdb_a", "asda_a", "amda_a" };

		TreeSet<String> fixTCRA = new TreeSet<String>(Arrays.asList(fixTCRAArr));

		String[] fixTCRSArr = { "si_e", "si_c", "su_e", "su_c", "ssd_e", "ssd_c", "smd_e", "smd_c", "psd_e", "psd_c", "pmd_e", "pmd_c",
				"ci_e", "ci_c", "cu_e", "cu_c", "csdb_e", "csdb_c", "cmdb_e", "cmdb_c", "csda_e", "csda_c", "cmda_e", "cmda_c", "prsd_e",
				"prsd_c", "prmd_e", "prmd_c", "ai_e", "ai_c", "au_e", "au_c", "asdb_e", "asdb_c", "amdb_e", "amdb_c", "asda_e", "asda_c",
				"amda_e", "amda_c" };

		TreeSet<String> fixTCRS = new TreeSet<String>(Arrays.asList(fixTCRSArr));

		String[] RTFTTCArr = { "si_dss_a", "si_dss_e", "si_dss_c", "si_gtp", "su_dss_a", "su_dss_e", "su_dss_c", "ssd_dss_a", "ssd_dss_e",
				"ssd_dss_c", "psd_dss_e", "psd_dss_c", "psd_gtp", "smd_dss_a", "smd_dss_e", "smd_dss_c", "pmd_dss_e", "pmd_dss_c", "pmd_gtp",
				"ci_dss_a", "ci_dss_e", "ci_dss_c", "ci_gtp", "cu_dss_a", "cu_dss_e", "cu_dss_c", "csdb_dss_a", "csdb_dss_e", "csdb_dss_c",
				"cmdb_dss_a", "cmdb_dss_e", "cmdb_dss_c", "csda_dss_a", "csda_dss_e", "csda_dss_c", "csda_gtp", "cmda_dss_a", "cmda_dss_e",
				"cmda_dss_c", "cmda_gtp", "pri_dss_a", "pru_dss_a", "prsda_dss_a", "prmda_dss_a", "prsd_dss_a", "prsd_dss_e", "prsd_dss_c",
				"prsd_gtp", "prmd_dss_a", "prmd_dss_e", "prmd_dss_c", "prmd_gtp", "ai_dss_a", "ai_dss_e", "ai_dss_c", "ai_gtp", "au_dss_a",
				"au_dss_e", "au_dss_c", "asdb_dss_a", "asdb_dss_e", "asdb_dss_c", "amdb_dss_a", "amdb_dss_e", "amdb_dss_c", "asda_dss_a",
				"asda_dss_e", "asda_dss_c", "asda_gtp", "amda_dss_a", "amda_dss_e", "amda_dss_c", "amda_gtp" };

		TreeSet<String> RTFTTC = new TreeSet<String>(Arrays.asList(RTFTTCArr));

		LinkedHashMap<String, Integer> instToUse = new LinkedHashMap<String, Integer>();

		//File inputFile = new File(HOME + XMLDB)); // WINDOWS
		File inputFile = new File(HOME + XMLDB); // LINUX
		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
		Document doc = dBuilder.parse(inputFile);
		doc.getDocumentElement().normalize();

		//Replace
//		String[] skipArr = { "RejectText", "NoTradePriceConditions", "NoTrdRegPublications", "NoSides" };
//		TreeSet<String> skipValues = new TreeSet<String>(Arrays.asList(skipArr));

		LinkedHashMap<String, String> currectInst = new LinkedHashMap<String, String>();

		for(int i = 0; i < lastRow; i++){
			Row row = sheet.getRow(i);
			if(row == null) continue;

			Cell cell6 = row.getCell(6);
			if(cell6 == null || cell6.getCellType() != Cell.CELL_TYPE_STRING) continue;

			//INSTRUMENT
			if(cell6.getStringCellValue().contains("InstrumentISIN")) continue;
			if(cell6.getStringCellValue().contains("InstrumentSource")) continue;
			if(cell6.getStringCellValue().contains("InstrumentCurrency")) continue;
			if(cell6.getStringCellValue().contains("InstrumentUniverse")) continue;
			if(cell6.getStringCellValue().contains("Instrument")) instToUse.put(cell6.getStringCellValue(), (int)row.getCell(13).getNumericCellValue());

			//IF InsAPA
			if(cell6.getStringCellValue().matches("InsAPA")){
				//System.out.println("securityID:" + row.getCell(14).getStringCellValue());
				String securityID = row.getCell(14).getStringCellValue().substring(2, row.getCell(14).getStringCellValue().length() - 1);
				//System.out.println("securityID:" + securityID);
				if(securityID.contains("InstrumentISIN")) originalTCR.put("instPrefix", securityID.replaceAll("InstrumentISIN", ""));
				else originalTCR.put("instPrefix", securityID.replaceAll("Instrument", ""));
				//System.out.println("instPrefix:" + originalTCR.get("instPrefix"));

				if(securityID.contains("ISIN")) securityID = securityID.replace("ISIN", "");
				Integer temp = instToUse.get(securityID);

				if(instMap.containsKey(temp)) currectInst = instMap.get(temp);
				else JOptionPane.showMessageDialog(new JFrame(), "Instrument " + temp + " info not found");
			}

			//IF PERSISTENCE
			if(!isPersistence && cell6.getStringCellValue().contains("test") && sheet.getRow(i).getCell(8).getStringCellValue().matches("LoadMessage")){
				persistenceName = cell6.getStringCellValue();
				isPersistence = true;
			}

			if(isPersistence && cell6.getStringCellValue().matches("si")) isPersistence = false; // IF NEW CASE RESET FLAG
			if(isPersistence && cell6.getStringCellValue().matches("su")) isPersistence = false; // IF NEW CASE RESET FLAG
			if(isPersistence && cell6.getStringCellValue().matches("ssd")) isPersistence = false; // IF NEW CASE RESET FLAG
			if(isPersistence && cell6.getStringCellValue().matches("smd")) isPersistence = false; // IF NEW CASE RESET FLAG


			//GROUP DATA
			if(groups.contains(cell6.getStringCellValue())){
				switch(cell6.getStringCellValue()){
					case("NoParty1"):
						Cell noParty1 = row.getCell(16);
						String noParty1Value = "";
						switch(noParty1.getCellType()){
							case(Cell.CELL_TYPE_NUMERIC): noParty1Value = String.valueOf((int)noParty1.getNumericCellValue()); break;
							case(Cell.CELL_TYPE_STRING): noParty1Value = noParty1.getStringCellValue(); break;
						}
						originalTCR.put("NoParty1", noParty1Value);
						break;

					case("Side1"):
						Cell side1 = row.getCell(14);
						String side1Value = "";
						switch(side1.getCellType()){
							case(Cell.CELL_TYPE_NUMERIC): side1Value = String.valueOf((int)side1.getNumericCellValue()); break;
							case(Cell.CELL_TYPE_STRING): side1Value = side1.getStringCellValue(); break;
						}
						originalTCR.put("Side1", side1Value);

						Cell TradingSessionSubId = row.getCell(19);
						if(TradingSessionSubId == null){
							originalTCR.put("TradingSessionSubId", "");
							continue;
						}

						String TradingSessionSubIdValue = "";
						switch(TradingSessionSubId.getCellType()){
							case(Cell.CELL_TYPE_NUMERIC): TradingSessionSubIdValue = String.valueOf((int)TradingSessionSubId.getNumericCellValue()); break;
							case(Cell.CELL_TYPE_STRING): TradingSessionSubIdValue = TradingSessionSubId.getStringCellValue(); break;
						}
						originalTCR.put("TradingSessionSubId", TradingSessionSubIdValue);
						break;
				}
			}

			//IF TCRC
			if(fixTCRC.contains(cell6.getStringCellValue())){
				amendTCR = new LinkedHashMap<String, String>();
				isNewFirmTradeID = false;
				Cell reference = row.getCell(6);
				String referenceValue = reference.getStringCellValue();

				if(amendTCRC.contains(cell6.getStringCellValue())) amendTCR.put("Reference", referenceValue);
				else originalTCR.put("Reference", referenceValue);

				String cellValue = "";
				for(int j = 14; j < row.getLastCellNum(); j++){
					Cell tempCell = row.getCell(j);
					if(tempCell == null) continue;

					switch(tempCell.getCellType()){
						case(Cell.CELL_TYPE_BLANK): cellValue = ""; break;
						case(Cell.CELL_TYPE_BOOLEAN): cellValue = Boolean.toString(tempCell.getBooleanCellValue()); break;
						case(Cell.CELL_TYPE_FORMULA): cellValue = Boolean.toString(tempCell.getBooleanCellValue()); break;
						case(Cell.CELL_TYPE_NUMERIC):
						if(tempCell.getNumericCellValue() % 1 == 0){
							Integer tmp = (int)tempCell.getNumericCellValue();
							cellValue = Integer.toString(tmp);
						} else cellValue = Double.toString(tempCell.getNumericCellValue());
						break;
						case(Cell.CELL_TYPE_STRING): cellValue = tempCell.getStringCellValue();
					}

					Cell headerCell = sheet.getRow( i - 1 ).getCell(j);
					if(headerCell == null || headerCell.getCellType() != Cell.CELL_TYPE_STRING) continue;

					if(amendTCRC.contains(cell6.getStringCellValue())) amendTCR.put(headerCell.getStringCellValue(), cellValue);
					else originalTCR.put(headerCell.getStringCellValue(), cellValue);
				}
				continue;
			}

			//IF Cancel || Pre-Release
			if(cell6.getStringCellValue().matches("ci")){
				Cell firmTradeIdCell = row.getCell(15);
				String firmTradeIdCellValue = firmTradeIdCell.getStringCellValue();
				if(firmTradeIdCellValue.contains("ClOrdID")) isNewFirmTradeID = true;
			}

			//IF TCR-ACK
			if(fixTCRA.contains(cell6.getStringCellValue())){
				System.out.println("TCR-ACK");
				int lastCell = sheet.getRow(i - 1).getLastCellNum();
				CellStyle style = row.getCell(0).getCellStyle();

				for(int j = 14; j < lastCell; j++){
					Cell tempHeaderCell = sheet.getRow(i - 1).getCell(j);
					if(tempHeaderCell == null) continue;
					String tagName = "", tagValue = "";
					int tagNumber;
					Element eElement1 = null;
					boolean isBug = false;

					switch(tempHeaderCell.getCellType()){
						case(Cell.CELL_TYPE_STRING): tagName = tempHeaderCell.getStringCellValue(); break;
						case(Cell.CELL_TYPE_NUMERIC): tagName = String.valueOf((int)tempHeaderCell.getNumericCellValue()); break;
					}

					switch(tagName){
						case("TradeID"): tagNumber = 0; break;
						case("FirmTradeID"): tagNumber = 1; break;
						case("Instrument"): tagNumber = 2; break;
						case("Currency"): tagNumber = 3; break;
						case("TradeReportRejectReason"): tagNumber = 4; break;
						case("RejectText"): tagNumber = 5; break;
						case("TradeReportTransType"): tagNumber = 6; break;
						case("TrdRptStatus"): tagNumber = 7; break;
						case(" "): tagNumber = -1; break;
						default: tagNumber = -1; break;
					}

					if(tagNumber == -1){
						continue;
					}
					
					
					/*if(tagNumber == -2){
						JOptionPane.showMessageDialog(null, "Unknown tag: " + tagName);
						continue;
					}*/

					NodeList nList1 = doc.getElementsByTagName("message");

			        switch(cell6.getStringCellValue()){
		        		case("si_a"): eElement1 = (Element) nList1.item(0); break;
		        		case("su_a"): eElement1 = (Element) nList1.item(1); break;
		        		case("ssd_a"): eElement1 = (Element) nList1.item(2); break;
		        		case("smd_a"): eElement1 = (Element) nList1.item(3); break;
		        		case("ci_a"): eElement1 = (Element) nList1.item(4); break;
		        		case("cu_a"): eElement1 = (Element) nList1.item(5); break;
		        		case("csdb_a"): eElement1 = (Element) nList1.item(6); break;
		        		case("cmdb_a"): eElement1 = (Element) nList1.item(7); break;
		        		case("csda_a"): eElement1 = (Element) nList1.item(8); break;
		        		case("cmda_a"): eElement1 = (Element) nList1.item(9); break;
		        		case("pri_a"): eElement1 = (Element) nList1.item(10); break;
		        		case("pru_a"): eElement1 = (Element) nList1.item(11); break;
		        		case("prsda_a"): eElement1 = (Element) nList1.item(12); break;
		        		case("prmda_a"): eElement1 = (Element) nList1.item(13); break;
		        		case("prsd_a"): eElement1 = (Element) nList1.item(14); break;
		        		case("prmd_a"): eElement1 = (Element) nList1.item(15); break;
		        		case("ai_a"): eElement1 = (Element) nList1.item(16); break;
		        		case("au_a"): eElement1 = (Element) nList1.item(17); break;
		        		case("asdb_a"): eElement1 = (Element) nList1.item(18); break;
		        		case("amdb_a"): eElement1 = (Element) nList1.item(19); break;
		        		case("asda_a"): eElement1 = (Element) nList1.item(20); break;
		        		case("amda_a"): eElement1 = (Element) nList1.item(21); break;
			        }

			        NodeList nList2 = eElement1.getElementsByTagName("tag");
			        Element eElement2 = (Element) nList2.item(tagNumber);

			        if(eElement2.getElementsByTagName("bug").item(0).getTextContent().matches("")){
			        	tagValue = eElement2.getElementsByTagName("original").item(0).getTextContent();
			        } else {
			        	tagValue = eElement2.getElementsByTagName("bug").item(0).getTextContent();
			        	isBug = true;
			        }

			        //НЕ МЕНЯЕМ
			        if(tagValue.equals("DON'T CHANGE")){						
			        	switch(row.getCell(j).getCellType()){
		        			case(Cell.CELL_TYPE_STRING): tagValue = row.getCell(j).getStringCellValue(); break;
		        			case(Cell.CELL_TYPE_NUMERIC): tagValue = String.valueOf((int)row.getCell(j).getNumericCellValue()); break;
			        	}
						//System.out.println("НЕ МЕНЯЕМ: " + tagValue);
			        }

			        //VALIDATIONS CONDITIONAL
			        if(!isBug && eElement2.getElementsByTagName("original").item(0).getTextContent().contains("CONDITIONAL")){
			        	switch(tagName){
			        		case("FirmTradeID"):
			        			String firmTradeID = eElement2.getElementsByTagName("original").item(0).getTextContent();
			        			String[] firmTradeIDArr = firmTradeID.split(";");

			        			if(!isNewFirmTradeID) tagValue = firmTradeIDArr[1];
			        			else tagValue = firmTradeIDArr[2];
			        			break;
			        	}
			        }

			        //CHECK FOR PERSISTENCE
			        if(isPersistence){
			        	if(tagValue.contains("si.")) tagValue = tagValue.replace("si.", persistenceName + ".");
			        	if(tagValue.contains("si_a.")) tagValue = tagValue.replace("si_a.", persistenceName + "Ack.");
			        	if(tagValue.contains("si_e.")) tagValue = tagValue.replace("si_e.", persistenceName + ".");

			        	if(tagValue.contains("su.")) tagValue = tagValue.replace("su.", persistenceName + ".");
			        	if(tagValue.contains("su_a.")) tagValue = tagValue.replace("su_a.", persistenceName + "Ack.");
			        	if(tagValue.contains("su_e.")) tagValue = tagValue.replace("su_e.", persistenceName + ".");

			        	if(tagValue.contains("ssd.")) tagValue = tagValue.replace("ssd.", persistenceName + ".");
			        	if(tagValue.contains("ssd_a.")) tagValue = tagValue.replace("ssd_a.", persistenceName + "Ack.");
			        	if(tagValue.contains("ssd_e.")) tagValue = tagValue.replace("ssd_e.", persistenceName + ".");

			        	if(tagValue.contains("smd.")) tagValue = tagValue.replace("smd.", persistenceName + ".");
			        	if(tagValue.contains("smd_a.")) tagValue = tagValue.replace("smd_a.", persistenceName + "Ack.");
			        	if(tagValue.contains("smd_e.")) tagValue = tagValue.replace("smd_e.", persistenceName + ".");
			        }

			        Cell tempMessageCell = row.createCell(j, Cell.CELL_TYPE_STRING);
					tempMessageCell.setCellValue(tagValue);
					tempMessageCell.setCellStyle(style);
				}
				//System.out.println("Row: " + i);
				if(row.getCell(19) != null && !row.getCell(19).getStringCellValue().matches("#")) isReject = true; else isReject = false;
			}

			//IF TCR-S
			if(fixTCRS.contains(cell6.getStringCellValue())){
				int lastCell = sheet.getRow(i - 1).getLastCellNum();
				CellStyle style = row.getCell(0).getCellStyle();

				//for(int j = 14; j < lastCell; j++){
				for(int j = 14; j < lastCell; j++){
					String tagName = "", tagValue = "";
					Element eElement1 = null;
					boolean isBug = false;

					tagName = tcrsHeader.get(j);

					//System.out.println("tagName: " + tagName);
					NodeList nList1 = doc.getElementsByTagName("message");

					//System.out.println("Message: " + cell6.getStringCellValue() + ", ROW: " + i);
					switch(cell6.getStringCellValue()){
						case("si_e"): eElement1 = (Element) nList1.item(22); break;
						case("si_c"): eElement1 = (Element) nList1.item(23); break;
						case("su_e"): eElement1 = (Element) nList1.item(24); break;
						case("su_c"): eElement1 = (Element) nList1.item(25); break;
						case("ssd_e"): eElement1 = (Element) nList1.item(26); break;
						case("ssd_c"): eElement1 = (Element) nList1.item(27); break;
						case("smd_e"): eElement1 = (Element) nList1.item(28); break;
						case("smd_c"): eElement1 = (Element) nList1.item(29); break;
						case("psd_e"): eElement1 = (Element) nList1.item(30); break;
						case("psd_c"): eElement1 = (Element) nList1.item(31); break;
						case("pmd_e"): eElement1 = (Element) nList1.item(32); break;
						case("pmd_c"): eElement1 = (Element) nList1.item(33); break;
						case("ci_e"): eElement1 = (Element) nList1.item(34); break;
						case("ci_c"): eElement1 = (Element) nList1.item(35); break;
						case("cu_e"): eElement1 = (Element) nList1.item(36); break;
						case("cu_c"): eElement1 = (Element) nList1.item(37); break;
						case("csdb_e"): eElement1 = (Element) nList1.item(38); break;
						case("csdb_c"): eElement1 = (Element) nList1.item(39); break;
						case("cmdb_e"): eElement1 = (Element) nList1.item(40); break;
						case("cmdb_c"): eElement1 = (Element) nList1.item(41); break;
						case("csda_e"): eElement1 = (Element) nList1.item(42); break;
						case("csda_c"): eElement1 = (Element) nList1.item(43); break;
						case("cmda_e"): eElement1 = (Element) nList1.item(44); break;
						case("cmda_c"): eElement1 = (Element) nList1.item(45); break;
						case("prsd_e"): eElement1 = (Element) nList1.item(46); break;
						case("prsd_c"): eElement1 = (Element) nList1.item(47); break;
						case("prmd_e"): eElement1 = (Element) nList1.item(48); break;
						case("prmd_c"): eElement1 = (Element) nList1.item(49); break;
						case("ai_e"): eElement1 = (Element) nList1.item(50); break;
						case("ai_c"): eElement1 = (Element) nList1.item(51); break;
						case("au_e"): eElement1 = (Element) nList1.item(52); break;
						case("au_c"): eElement1 = (Element) nList1.item(53); break;
						case("asdb_e"): eElement1 = (Element) nList1.item(54); break;
						case("asdb_c"): eElement1 = (Element) nList1.item(55); break;
						case("amdb_e"): eElement1 = (Element) nList1.item(56); break;
						case("amdb_c"): eElement1 = (Element) nList1.item(57); break;
						case("asda_e"): eElement1 = (Element) nList1.item(58); break;
						case("asda_c"): eElement1 = (Element) nList1.item(59); break;
						case("amda_e"): eElement1 = (Element) nList1.item(60); break;
						case("amda_c"): eElement1 = (Element) nList1.item(61); break;
					}

					if(eElement1 == null) System.out.println("NULL");
					NodeList nList2 = eElement1.getElementsByTagName("tag");

					Element eElement2 = (Element) nList2.item(j - 14);
					if(eElement2 == null) continue;

					// IF BUG is SET UP
					//System.out.println("Row: " + i + ", cell: " + j);
					if(!eElement2.getElementsByTagName("bug").item(0).getTextContent().matches("")){
						//System.out.println("originalTCR: " + originalTCR);
						isBuy = originalTCR.get("Side1").matches("BUY");
						switch(tagName){
							case("TradeNumber"):
							case("TotNumTradeReports"):
								String instTotNumTradeReports = currectInst.get("Sub Asset Class");
								String venueTotNumTradeReports = originalTCR.get("VenueType");
								String matchTotNumTradeReports = originalTCR.get("MatchType");
								if(!instTotNumTradeReports.matches("Shares") && (venueTotNumTradeReports.matches("O")
										&& (matchTotNumTradeReports.matches("9") || matchTotNumTradeReports.matches("1")) || venueTotNumTradeReports.matches("D"))){
									tagValue = eElement2.getElementsByTagName("bug").item(0).getTextContent();
								} //else tagValue = eElement2.getElementsByTagName("original").item(0).getTextContent();
								break;


							case("TrdSubType"):
								if(originalTCR.get("TrdSubType").matches("")){
									tagValue = eElement2.getElementsByTagName("original").item(0).getTextContent();
									//System.out.println("original");
									break;
								}

								String venueTypeTrdType = originalTCR.get("VenueType");
								String matchTypeTrdType = originalTCR.get("MatchType");

								if(venueTypeTrdType.matches("O") && (matchTypeTrdType.matches("9") || matchTypeTrdType.matches("1"))){
									tagValue = eElement2.getElementsByTagName("bug").item(0).getTextContent();
								} else tagValue = eElement2.getElementsByTagName("original").item(0).getTextContent();
								isBug = true;
								break;

							case("QtyType"):
								tagValue = eElement2.getElementsByTagName("bug").item(0).getTextContent();
								isBug = true;
								break;

							case("Issuer"):
								tagValue = eElement2.getElementsByTagName("bug").item(0).getTextContent();
								isBug = true;
								break;

							case("AggPublicationID"):
								tagValue = eElement2.getElementsByTagName("bug").item(0).getTextContent();
								isBug = true;
								break;

							case("TradeReportID"):
								if(isBuy) tagValue = eElement2.getElementsByTagName("bug").item(0).getTextContent();
								else {
									String tradeReportID = eElement2.getElementsByTagName("original").item(0).getTextContent();
									String[] tradeReportIDArr = tradeReportID.split(";");
									tagValue = tradeReportIDArr[2];
								}
								isBug = true;
								break;

							default:
								System.out.println("WARNING ADD BUG!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
								break;
						}
					}

					// IF DEFAULT + TCRC
					else if(!eElement2.getElementsByTagName("default").item(0).getTextContent().matches("") && !originalTCR.get(tagName).matches("")){
						tagValue = originalTCR.get(tagName);
					}

					// IF DEFAULT + !TCRC
					else if(!eElement2.getElementsByTagName("default").item(0).getTextContent().matches("") && originalTCR.get(tagName).matches("")){
						tagValue = eElement2.getElementsByTagName("default").item(0).getTextContent();
						//System.out.println("IN DEFAULT: " + tagValue);
					}

					else { tagValue = eElement2.getElementsByTagName("original").item(0).getTextContent();
						//System.out.println("original tagName: " + tagName);
					}

					//System.out.println("tagName: " + tagName + ", tagValue: " + tagValue);

					//НЕ МЕНЯЕМ
					if(tagValue.equals("DON'T CHANGE")){ tagValue = row.getCell(j).getStringCellValue(); }

					//VALIDATIONS CONDITIONAL
					if(!isBug && eElement2.getElementsByTagName("original").item(0).getTextContent().contains("CONDITIONAL")){
						isBuy = originalTCR.get("Side1").matches("BUY");

						switch(tagName){
							case("PreviouslyReported"):
								// OTC & SI validation
								String instPreviouslyReported = currectInst.get("Sub Asset Class");
								String venuePreviouslyReported = originalTCR.get("VenueType");
								String matchPreviouslyReported = originalTCR.get("MatchType");

								if(!instPreviouslyReported.matches("Shares") && venuePreviouslyReported.matches("O")
										&& (matchPreviouslyReported.matches("9") || matchPreviouslyReported.matches("1"))){
									tagValue = "N";
									//originalTCR.replace("PreviouslyReported", "N");
									originalTCR.put("PreviouslyReported", "N");
								}
								break;



							case("TrdType"):
								// OTC & SI validation
								String instTrdType = currectInst.get("Sub Asset Class");
								String venueTypeTrdType = originalTCR.get("VenueType");
								String matchTypeTrdType = originalTCR.get("MatchType");

								if(instTrdType.matches("Shares") && venueTypeTrdType.matches("O")
										&& (matchTypeTrdType.matches("9") || matchTypeTrdType.matches("1"))){
									tagValue = "0";
									//originalTCR.replace("TrdType", "0");
									originalTCR.put("TrdType", "0");
								}
								break;

							case("AlgorithmicTradeIndicator"):
								// OTC & SI validation
								String venueTypeAlgorithmicTradeIndicator = originalTCR.get("VenueType");
								String matchTypeAlgorithmicTradeIndicator = originalTCR.get("MatchType");

								if(venueTypeAlgorithmicTradeIndicator.matches("O")
										&& (matchTypeAlgorithmicTradeIndicator.matches("1") || matchTypeAlgorithmicTradeIndicator.matches("9"))){
									tagValue = "0";
									//originalTCR.replace("AlgorithmicTradeIndicator", "0");
									originalTCR.put("AlgorithmicTradeIndicator", "0");
								}
								break;

							case("FirmTradeID"):
								String firmTradeID = eElement2.getElementsByTagName("original").item(0).getTextContent();
								String[] firmTradeIDArr = firmTradeID.split(";");

								if(!isNewFirmTradeID) tagValue = firmTradeIDArr[1];
								else tagValue = firmTradeIDArr[2];
								break;

							case("TradeReportID"): //
								String tradeReportID = eElement2.getElementsByTagName("original").item(0).getTextContent();
								String[] tradeReportIDArr = tradeReportID.split(";");
								if(tradeReportIDArr.length > 2) if(isBuy) tagValue = tradeReportIDArr[1]; else tagValue = tradeReportIDArr[2];
								break;

							case("DelayToTime"):
								//System.out.println("DelayToTime");
								//System.out.println("Orig: " + originalTCR);
								//System.out.println("Amend: " + amendTCR);

								if(originalTCR.get("DelayToTime").matches("") && !amendTCR.get("DelayToTime").matches("")){
									//System.out.println("IN1: " + amendTCR.get("Reference"));
									switch(amendTCR.get("Reference")){
										case("asdb"): tagValue = "${asdb.DelayToTime}"; break;
		 								case("asda"): tagValue = "${asda.DelayToTime}"; break;
									}

								} else if(!originalTCR.get("DelayToTime").matches("") && amendTCR.get("DelayToTime").matches("")){
									System.out.println("IN2: " + amendTCR.get("Reference"));
									switch(amendTCR.get("Reference")){
										case("amdb"): tagValue = "#"; break;
		 								case("amda"): tagValue = "#"; break;
									}
								} else {
									//System.out.println("IN3:");
									String delayToTime = eElement2.getElementsByTagName("original").item(0).getTextContent();
									String[] delayToTimeArr = delayToTime.split(";");
									//System.out.println("delayToTimeArr:" + Arrays.toString(delayToTimeArr));
									tagValue = delayToTimeArr[1];
								}

								//System.out.println("tagValue: " + tagValue);
								break;


							case("RptTime"):
								//System.out.println("RptTime");
								//System.out.println("Orig: " + originalTCR);
								//System.out.println("Amend: " + amendTCR);

								if(originalTCR.get("DelayToTime").matches("") && !amendTCR.get("DelayToTime").matches("")){
									//System.out.println("IN1: " + amendTCR.get("Reference"));
									switch(amendTCR.get("Reference")){
										case("asdb"): tagValue = "${asdb.DelayToTime}"; break;
		 								case("asda"): tagValue = "${asda.DelayToTime}"; break;
									}

								} else if(!originalTCR.get("DelayToTime").matches("") && amendTCR.get("DelayToTime").matches("")){
									//System.out.println("IN2: " + amendTCR.get("Reference"));
									switch(amendTCR.get("Reference")){
										case("amdb"): tagValue = "*"; break;
		 								case("amda"): tagValue = "*"; break;
									}
								} else {
									//System.out.println("IN3:");
									String delayToTime = eElement2.getElementsByTagName("original").item(0).getTextContent();
									String[] delayToTimeArr = delayToTime.split(";");
									//System.out.println("delayToTimeArr:" + Arrays.toString(delayToTimeArr));
									tagValue = delayToTimeArr[1];
								}

								//System.out.println("tagValue: " + tagValue);
								break;



//							case("DelayToTime"): //
//								String delayToTime = eElement2.getElementsByTagName("original").item(0).getTextContent();
//								String[] delayToTimeArr = delayToTime.split(";");
//
//		        					if(!amendTCR.get("DelayToTime").matches("")) tagValue = delayToTimeArr[1]; else tagValue = delayToTimeArr[2];
//		        					break;
						}
					}

					//CHECK FOR PERSISTENCE
					if(isPersistence){
						if(tagValue.contains("si.")) tagValue = tagValue.replace("si.", persistenceName + ".");
						if(tagValue.contains("si_a.")) tagValue = tagValue.replace("si_a.", persistenceName + "Ack.");
						if(tagValue.contains("si_e.")) tagValue = tagValue.replace("si_e.", persistenceName + ".");

						if(tagValue.contains("su.")) tagValue = tagValue.replace("su.", persistenceName + ".");
						if(tagValue.contains("su_a.")) tagValue = tagValue.replace("su_a.", persistenceName + "Ack.");
						if(tagValue.contains("su_e.")) tagValue = tagValue.replace("su_e.", persistenceName + ".");

						if(tagValue.contains("ssd.")) tagValue = tagValue.replace("ssd.", persistenceName + ".");
						if(tagValue.contains("ssd_a.")) tagValue = tagValue.replace("ssd_a.", persistenceName + "Ack.");
						if(tagValue.contains("ssd_e.")) tagValue = tagValue.replace("ssd_e.", persistenceName + ".");

						if(tagValue.contains("smd.")) tagValue = tagValue.replace("smd.", persistenceName + ".");
						if(tagValue.contains("smd_a.")) tagValue = tagValue.replace("smd_a.", persistenceName + "Ack.");
						if(tagValue.contains("smd_e.")) tagValue = tagValue.replace("smd_e.", persistenceName + ".");
					}


					Cell tempMessageCell = row.createCell(j, Cell.CELL_TYPE_STRING);
					tempMessageCell.setCellValue(tagValue);
					tempMessageCell.setCellStyle(style);
				}
			}

			//TTC MESSAGES
			if(RTFTTC.contains(cell6.getStringCellValue())){
//				int lastCell = sheet.getRow(i - 1).getLastCellNum();
				CellStyle style = row.getCell(0).getCellStyle();

				//for(int j = 14; j < lastCell; j++){
				for(int j = 14; j < rtfTagsNumber; j++){
					String tagValue = "";
					Element eElement1 = null;
					NodeList nList1 = doc.getElementsByTagName("message");
					boolean isBug = false;

					switch(cell6.getStringCellValue()){
	        			case("si_dss_a"): eElement1 = (Element) nList1.item(62); break;
	        			case("si_dss_e"): eElement1 = (Element) nList1.item(63); break;
	        			case("si_dss_c"): eElement1 = (Element) nList1.item(64); break;
		        		case("si_gtp"): eElement1 = (Element) nList1.item(65); break;
		        		case("su_dss_a"): eElement1 = (Element) nList1.item(66); break;
		        		case("su_dss_e"): eElement1 = (Element) nList1.item(67); break;
		        		case("su_dss_c"): eElement1 = (Element) nList1.item(68); break;
		        		case("ssd_dss_a"): eElement1 = (Element) nList1.item(69); break;
		        		case("ssd_dss_e"): eElement1 = (Element) nList1.item(70); break;
		        		case("ssd_dss_c"): eElement1 = (Element) nList1.item(71); break;
		        		case("psd_dss_e"): eElement1 = (Element) nList1.item(72); break;
		        		case("psd_dss_c"): eElement1 = (Element) nList1.item(73); break;
		        		case("psd_gtp"): eElement1 = (Element) nList1.item(74); break;
		        		case("smd_dss_a"): eElement1 = (Element) nList1.item(75); break;
		        		case("smd_dss_e"): eElement1 = (Element) nList1.item(76); break;
		        		case("smd_dss_c"): eElement1 = (Element) nList1.item(77); break;
		        		case("pmd_dss_e"): eElement1 = (Element) nList1.item(78); break;
		        		case("pmd_dss_c"): eElement1 = (Element) nList1.item(79); break;
		        		case("pmd_gtp"): eElement1 = (Element) nList1.item(80); break;
		        		case("ci_dss_a"): eElement1 = (Element) nList1.item(81); break;
		        		case("ci_dss_e"): eElement1 = (Element) nList1.item(82); break;
		        		case("ci_dss_c"): eElement1 = (Element) nList1.item(83); break;
		        		case("ci_gtp"): eElement1 = (Element) nList1.item(84); break;
		        		case("cu_dss_a"): eElement1 = (Element) nList1.item(85); break;
		        		case("cu_dss_e"): eElement1 = (Element) nList1.item(86); break;
		        		case("cu_dss_c"): eElement1 = (Element) nList1.item(87); break;
		        		case("csdb_dss_a"): eElement1 = (Element) nList1.item(88); break;
		        		case("csdb_dss_e"): eElement1 = (Element) nList1.item(89); break;
		        		case("csdb_dss_c"): eElement1 = (Element) nList1.item(90); break;
		        		case("cmdb_dss_a"): eElement1 = (Element) nList1.item(91); break;
		        		case("cmdb_dss_e"): eElement1 = (Element) nList1.item(92); break;
		        		case("cmdb_dss_c"): eElement1 = (Element) nList1.item(93); break;
		        		case("csda_dss_a"): eElement1 = (Element) nList1.item(94); break;
		        		case("csda_dss_e"): eElement1 = (Element) nList1.item(95); break;
		        		case("csda_dss_c"): eElement1 = (Element) nList1.item(96); break;
		        		case("csda_gtp"): eElement1 = (Element) nList1.item(97); break;
		        		case("cmda_dss_a"): eElement1 = (Element) nList1.item(98); break;
		        		case("cmda_dss_e"): eElement1 = (Element) nList1.item(99); break;
		        		case("cmda_dss_c"): eElement1 = (Element) nList1.item(100); break;
		        		case("cmda_gtp"): eElement1 = (Element) nList1.item(101); break;
		        		case("pri_dss_a"): eElement1 = (Element) nList1.item(102); break;
		        		case("pru_dss_a"): eElement1 = (Element) nList1.item(103); break;
		        		case("prsda_dss_a"): eElement1 = (Element) nList1.item(104); break;
		        		case("prmda_dss_a"): eElement1 = (Element) nList1.item(105); break;
		        		case("prsd_dss_a"): eElement1 = (Element) nList1.item(106); break;
		        		case("prsd_dss_e"): eElement1 = (Element) nList1.item(107); break;
		        		case("prsd_dss_c"): eElement1 = (Element) nList1.item(108); break;
		        		case("prsd_gtp"): eElement1 = (Element) nList1.item(109); break;
		        		case("prmd_dss_a"): eElement1 = (Element) nList1.item(110); break;
		        		case("prmd_dss_e"): eElement1 = (Element) nList1.item(111); break;
		        		case("prmd_dss_c"): eElement1 = (Element) nList1.item(112); break;
		        		case("prmd_gtp"): eElement1 = (Element) nList1.item(113); break;
		        		case("ai_dss_a"): eElement1 = (Element) nList1.item(114); break;
		        		case("ai_dss_e"): eElement1 = (Element) nList1.item(115); break;
		        		case("ai_dss_c"): eElement1 = (Element) nList1.item(116); break;
		        		case("ai_gtp"): eElement1 = (Element) nList1.item(117); break;
		        		case("au_dss_a"): eElement1 = (Element) nList1.item(118); break;
		        		case("au_dss_e"): eElement1 = (Element) nList1.item(119); break;
		        		case("au_dss_c"): eElement1 = (Element) nList1.item(120); break;
		        		case("asdb_dss_a"): eElement1 = (Element) nList1.item(121); break;
		        		case("asdb_dss_e"): eElement1 = (Element) nList1.item(122); break;
		        		case("asdb_dss_c"): eElement1 = (Element) nList1.item(123); break;
		        		case("amdb_dss_a"): eElement1 = (Element) nList1.item(124); break;
		        		case("amdb_dss_e"): eElement1 = (Element) nList1.item(125); break;
		        		case("amdb_dss_c"): eElement1 = (Element) nList1.item(126); break;
		        		case("asda_dss_a"): eElement1 = (Element) nList1.item(127); break;
		        		case("asda_dss_e"): eElement1 = (Element) nList1.item(128); break;
		        		case("asda_dss_c"): eElement1 = (Element) nList1.item(129); break;
		        		case("asda_gtp"): eElement1 = (Element) nList1.item(130); break;
		        		case("amda_dss_a"): eElement1 = (Element) nList1.item(131); break;
		        		case("amda_dss_e"): eElement1 = (Element) nList1.item(132); break;
		        		case("amda_dss_c"): eElement1 = (Element) nList1.item(133); break;
		        		case("amda_gtp"): eElement1 = (Element) nList1.item(134); break;
			        }

			        NodeList nList2 = eElement1.getElementsByTagName("tag");
			        Element eElement2 = (Element) nList2.item(j - 14);

			        // IF BUG is SET UP
			        //NEW BUG CHECK FOR CROSS BUGS
			        if(!isBug && originalTCR.get("Side1").matches("8")){
			        	if(eElement2.getElementsByTagName("crossBug").item(0) != null &&
			        			!eElement2.getElementsByTagName("crossBug").item(0).getTextContent().matches("")){
				        	tagValue = eElement2.getElementsByTagName("crossBug").item(0).getTextContent();
				        	isBug = true;
			        	}
			        }

			        //System.out.println("tag: " + i + ", " + j);
			        if(!isBug && !eElement2.getElementsByTagName("bug").item(0).getTextContent().matches("")){
			        	isBug = true;
						switch(j){
							case(31): //  ReferencePrice
								//System.out.println("ReferencePrice");
								if(originalTCR.get("NoTradePriceConditions").contains("NPFT") &&
										originalTCR.get("VenueType").matches("D") &&
										originalTCR.get("MatchType").matches("")){
									//System.out.println(originalTCR.get("NoTradePriceConditions"));
									tagValue = eElement2.getElementsByTagName("bug").item(0).getTextContent();
								} else { tagValue = eElement2.getElementsByTagName("original").item(0).getTextContent(); isBug = false; }


								break;

							case(99): // EmissionAllowanceType
								if(cell6.getStringCellValue().contains("_a")){
									switch(currectInst.get("Emission Allowance Type")){
										case("EUA"): case("CER"):case("ERU"):
											tagValue = eElement2.getElementsByTagName("original").item(0).getTextContent(); break;
										case("EUAA"): case("OTHR"):
											tagValue = eElement2.getElementsByTagName("bug").item(0).getTextContent(); break;
										case(""): tagValue = eElement2.getElementsByTagName("original").item(0).getTextContent(); break;
									}
								} else {
									switch(currectInst.get("Emission Allowance Type")){
										case("EUA"): case("CER"):case("ERU"):
											tagValue = eElement2.getElementsByTagName("bug").item(0).getTextContent(); break;
										case("EUAA"): case("OTHR"):
											tagValue = eElement2.getElementsByTagName("original").item(0).getTextContent(); break;
										case(""): tagValue = eElement2.getElementsByTagName("original").item(0).getTextContent(); break;
									}
								}
								break;

							case(109): //  DelegatedReport
								//System.out.println("DelegatedReport");
								String delegatedReportValue = originalTCR.get("AssistedReportAPA");
								switch(delegatedReportValue){
									case(""):
									case("2"): tagValue = "#{ExpectedEmpty().Bug(\"BSRSUP-218\", 0, \"Protocol issues\").Actual(#{toInteger(x)})}"; break;
									case("1"): tagValue = "#{ExpectedEmpty().Bug(\"BSRSUP-218\", 1, \"Protocol issues\").Actual(#{toInteger(x)})}"; break;
								}
								break;

							default: tagValue = eElement2.getElementsByTagName("bug").item(0).getTextContent(); break;
							}
			        }

			        if(!isBug) tagValue = eElement2.getElementsByTagName("original").item(0).getTextContent();


			        ///System.out.println("IS BUG: " + isBug + ", tagValue: " + tagValue);
			        //НЕ МЕНЯЕМ
			        if(!isBug && tagValue.equals("DON'T CHANGE")){
			        	switch(row.getCell(j).getCellType()){
			        		case(Cell.CELL_TYPE_STRING): tagValue = row.getCell(j).getStringCellValue(); break;
			        		case(Cell.CELL_TYPE_NUMERIC): tagValue = String.valueOf((int)row.getCell(j).getNumericCellValue()); break;
			        	}
			        }

			        //VALIDATIONS CONDITIONAL
			        if(!isBug && eElement2.getElementsByTagName("original").item(0).getTextContent().contains("CONDITIONAL")){
			        	isBuy = originalTCR.get("Side1").matches("BUY");
			        	//System.out.println("JJJJJJJJJJJJ: " + j);

			        	switch(j){
			        		case(21): // PriceNotation
			        			if(originalTCR.get("PriceType").matches("")) tagValue = "0";
			        			if(originalTCR.get("PriceType").matches("2")) tagValue = "0";
			        			if(originalTCR.get("PriceType").matches("1")) tagValue = "1";
			        			if(originalTCR.get("PriceType").matches("9")) tagValue = "2";
			        			if(originalTCR.get("PriceType").matches("22")) tagValue = "3";
			        			break;

			        		case(22): // ExecutedSize
			        			//System.out.println("Denominated Par Value: " + currectInst.get("Denominated Par Value"));
				        		String executedSize = eElement2.getElementsByTagName("original").item(0).getTextContent();
			        			String[] executedSizeArr = executedSize.split(";");
			        			tagValue = "";
			        			//System.out.println("Denominated Par Value: " + currectInst.get("Denominated Par Value"));
			        			if(currectInst.get("Denominated Par Value").matches("")){
			        				tagValue = executedSizeArr[1]; break;
			        			}

			        			switch(currectInst.get("Sub Asset Class")){
			        				case("Sovereign Bonds"):
			        				case("Other Public Bonds"):
			        				case("Convertible Bonds"):
			        				case("Covered Bonds"):
			        				case("Corporate Bonds"):
			        				case("Other Bonds"):
			        					if(currectInst.get("Denominated Par Value").matches("0")) tagValue = executedSizeArr[1];
			        					else tagValue = executedSizeArr[2];
			        					break;
			        				default: tagValue = executedSizeArr[1]; break;
			        			}
			        			break;

			        		case(31): // ReferencePrice
			        			//if(originalTCR.get("TradePublishIndicator").matches("0") && !currectInst.get("Reference Price").matches("")) tagValue = "#";
			        			//else tagValue =  "%{ReferencePrice} != '#' ? x == %{ReferencePrice}: x == null"; //REWRITE
							if(originalTCR.get("TradePublishIndicator").matches("0")) tagValue = "#";
							else tagValue =  "%{ReferencePrice} != '#' ? x == %{ReferencePrice}: x == null";
			        			break;

			        		case(37): // VenueType
			        			if(amendTCR.get("VenueType") == null || amendTCR.get("VenueType").contains("VenueType")){
				        			if(originalTCR.get("MatchType").matches("1")) tagValue = "#";
									if(originalTCR.get("MatchType").matches("3")) tagValue = "3";
									if(originalTCR.get("MatchType").matches("9")) tagValue = "#";
									if(originalTCR.get("MatchType").matches("") && originalTCR.get("VenueType").matches("D") &&
											originalTCR.get("NoParty1").matches("73")) tagValue = "OTF";
									if(originalTCR.get("MatchType").matches("") && originalTCR.get("VenueType").matches("D") &&
											originalTCR.get("NoParty1").matches("64")) tagValue = "MTF";
								} else {
				        			if(amendTCR.get("MatchType").matches("1")) tagValue = "#";
									if(amendTCR.get("MatchType").matches("3")) tagValue = "3";
									if(amendTCR.get("MatchType").matches("9")) tagValue = "#";
									if(amendTCR.get("MatchType").matches("") && amendTCR.get("VenueType").matches("D") &&
											originalTCR.get("NoParty1").matches("73")) tagValue = "OTF";
									if(amendTCR.get("MatchType").matches("") && amendTCR.get("VenueType").matches("D") &&
											originalTCR.get("NoParty1").matches("64")) tagValue = "MTF";
								}
								break;

			        		case(42): // Capacity
			        			if(isBuy && cell6.getStringCellValue().contains("_e")) tagValue = "0";
				        		if(isBuy && cell6.getStringCellValue().contains("_c")) tagValue = "#";
				        		if(!isBuy && cell6.getStringCellValue().contains("_e")) tagValue = "#";
				        		if(!isBuy && cell6.getStringCellValue().contains("_c")) tagValue = "0";
			        			break;

			        		case(43): // OriginalCapacity
			        			if(isBuy && cell6.getStringCellValue().contains("_e")) tagValue = "4";
				        		if(isBuy && cell6.getStringCellValue().contains("_c")) tagValue = "#";
				        		if(!isBuy && cell6.getStringCellValue().contains("_e")) tagValue = "#";
				        		if(!isBuy && cell6.getStringCellValue().contains("_c")) tagValue = "4";
			        			break;

			        		case(45): // IsAggressor
			        			if(isBuy && cell6.getStringCellValue().contains("_e")) tagValue = "1";
				        		if(isBuy && cell6.getStringCellValue().contains("_c")) tagValue = "#";
				        		if(!isBuy && cell6.getStringCellValue().contains("_e")) tagValue = "#";
				        		if(!isBuy && cell6.getStringCellValue().contains("_c")) tagValue = "1";
			        			if(originalTCR.get("Side1").matches("8") && cell6.getStringCellValue().contains("_e")) tagValue = "1";
			        			if(originalTCR.get("Side1").matches("8") && cell6.getStringCellValue().contains("_c")) tagValue = "#";
			        			break;

			        		case(56): // Side
			        			if(originalTCR.get("Side1").matches("8")) tagValue = "3";
			        			else {
			        				if(cell6.getStringCellValue().contains("_e")) tagValue = "1";
				        			if(cell6.getStringCellValue().contains("_c")) tagValue = "2";
				        			if(isBuy && cell6.getStringCellValue().contains("_a")) tagValue = "1";
				        			if(isBuy && cell6.getStringCellValue().contains("_gtp")) tagValue = "1";
				        			if(!isBuy && cell6.getStringCellValue().contains("_a")) tagValue = "2";
				        			if(!isBuy && cell6.getStringCellValue().contains("_gtp")) tagValue = "2";
			        			}
				        		break;

			        		case(60): // BookDefinitionID
			        			if(originalTCR.get("VenueType").matches("D")) tagValue = "0";
			        			else tagValue = "1";
			        			break;

			        		case(63): // ClientID
			        			String clientID = eElement2.getElementsByTagName("original").item(0).getTextContent();
			        			String[] clientIDArr = clientID.split(";");
			        			if(clientIDArr.length > 2) if(isBuy) tagValue = clientIDArr[1]; else tagValue = clientIDArr[2];
			        			break;

			        		case(65): // ContraClientID
			        			String contraClientID = eElement2.getElementsByTagName("original").item(0).getTextContent();
			        			String[] contraClientIDArr = contraClientID.split(";");
			        			if(contraClientIDArr.length > 2) if(isBuy) tagValue = contraClientIDArr[1]; else tagValue = contraClientIDArr[2];
			        			break;

			        		case(66): // ContraFirmPartyID
			        			if(originalTCR.get("Side1").matches("8")){
			        				if(cell6.getStringCellValue().contains("_a")) tagValue = "#";
				        			if(cell6.getStringCellValue().contains("_e")) tagValue = "#";
				        			if(cell6.getStringCellValue().contains("_c")) tagValue = "%{Executer}";
				        			if(cell6.getStringCellValue().contains("_gtp")) tagValue = "#";
			        			} else {
			        				String contra = "%{ConterParty}", executer = "%{Executer}";
			        				if(!originalTCR.get("NoSides").contains("Side2")) contra = "#";

			        				if(cell6.getStringCellValue().contains("_a")) tagValue = contra;
				        			if(isBuy && cell6.getStringCellValue().contains("_c")) tagValue = executer;
				        			if(isBuy && cell6.getStringCellValue().contains("_e")) tagValue = contra;
				        			if(!isBuy && cell6.getStringCellValue().contains("_c")) tagValue = contra;
				        			if(!isBuy && cell6.getStringCellValue().contains("_e")) tagValue = executer;
				        			if(cell6.getStringCellValue().contains("_gtp")) tagValue = contra;
			        			}
			        			break;

			        		case(67): // ContraOwnerPartyID
			        			if(originalTCR.get("Side1").matches("8")){
			        				if(cell6.getStringCellValue().contains("_a")) tagValue = "#";
				        			if(cell6.getStringCellValue().contains("_e")) tagValue = "#";
				        			if(cell6.getStringCellValue().contains("_c")) tagValue = "%{ExecuterFIX}";
				        			if(cell6.getStringCellValue().contains("_gtp")) tagValue = "#";
			        			} else {
			        				String contra = "%{ContraFIX}", executer = "%{ExecuterFIX}";
			        				if(!originalTCR.get("NoSides").contains("Side2")) contra = "#";

			        				if(cell6.getStringCellValue().contains("_a")) tagValue = contra;
				        			if(isBuy && cell6.getStringCellValue().contains("_c")) tagValue = executer;
				        			if(isBuy && cell6.getStringCellValue().contains("_e")) tagValue = contra;
				        			if(!isBuy && cell6.getStringCellValue().contains("_c")) tagValue = contra;
				        			if(!isBuy && cell6.getStringCellValue().contains("_e")) tagValue = executer;
				        			if(cell6.getStringCellValue().contains("_gtp")) tagValue = contra;
			        			}
			        			break;

			        		case(74): // ExecutingFirm
			        			if(isBuy && cell6.getStringCellValue().contains("_e")) tagValue = "%{Executer}";
				        		if(isBuy && cell6.getStringCellValue().contains("_c")) tagValue = "%{ConterParty}";
				        		if(!isBuy && cell6.getStringCellValue().contains("_e")) tagValue = "%{ConterParty}";
				        		if(!isBuy && cell6.getStringCellValue().contains("_c")) tagValue = "%{Executer}";

				        		if(originalTCR.get("Side1").matches("8") && cell6.getStringCellValue().contains("_e")) tagValue = "%{Executer}";
				        		if(originalTCR.get("Side1").matches("8") && cell6.getStringCellValue().contains("_c")) tagValue = "#";

				        		break;

			        		case(76): //FirmTradeID
			        			String firmTradeID = eElement2.getElementsByTagName("original").item(0).getTextContent();
			        			String[] firmTradeIDArr = firmTradeID.split(";");

			        			if(!isNewFirmTradeID) tagValue = firmTradeIDArr[1];
			        			else tagValue = firmTradeIDArr[2];
			        			break;

			        		case(83): // OwnerID
			        			if(isBuy && cell6.getStringCellValue().contains("_e")) tagValue = "%{ExecuterFIX}";
				        		if(isBuy && cell6.getStringCellValue().contains("_c")) tagValue = "%{ContraFIX}";
				        		if(!isBuy && cell6.getStringCellValue().contains("_e")) tagValue = "%{ContraFIX}";
				        		if(!isBuy && cell6.getStringCellValue().contains("_c")) tagValue = "%{ExecuterFIX}";
				        		if(originalTCR.get("Side1").matches("8") && cell6.getStringCellValue().contains("_e")) tagValue = "%{ExecuterFIX}";
				        		if(originalTCR.get("Side1").matches("8") && cell6.getStringCellValue().contains("_c")) tagValue = "#";
			        			break;

			        		case(86): // SettlementDate
			        			if(originalTCR.get("SettlDate").matches("")) tagValue = "#";
			        			else tagValue = "#{formatDate(${" + originalTCR.get("Reference") + ".SettlDate}, \"YYYYMMdd\")}";
			        			break;

			        		case(93): // VenueOfExecution
			        			if(originalTCR.get("Side1").matches("8")) tagValue = "%{MIC}";
			        			else if(originalTCR.get("MatchType").matches("1")) tagValue = "XOFF";
			        			else if(originalTCR.get("MatchType").matches("3")) tagValue = "%{MarketCode}";
			        			else if(originalTCR.get("MatchType").matches("9")) tagValue = "SINT";
				        		break;

			        		case(94): // VenueOfPublication
			        			if(originalTCR.get("Side1").matches("8")) tagValue = "%{MIC}"; //////////////////////!!!!!!!!!!!!!!!!!!!!!!!!!!
							else{
								String VenueOfPublication = eElement2.getElementsByTagName("original").item(0).getTextContent();
			        			String[] VenueOfPublicationArr = VenueOfPublication.split(";");
								tagValue = VenueOfPublicationArr[1];
							}
				        		break;

			        		case(105): // ClosingPrice
			        			//System.out.println("Current Inst: " + currectInst);
			        			//System.out.println("Current Inst ClosingPrice: " + currectInst.get("Closing Price"));
			        			if(currectInst.get("Closing Price").matches("")) tagValue = "#";
			        			else tagValue = "#{ExpectedEmpty().Bug(\"BSRSUP-218\", %{ClosingPrice}, \"Protocol issues\")}"; //!!!! REWRITE
				        		break;

			        		case(106): // ParValue
			        			if(currectInst.get("Par Value").matches("")) tagValue = "#";
			        			else tagValue = "#{ExpectedEmpty().Bug(\"BSRSUP-218\", %{ParValue}, \"Protocol issues\")}"; //!!!! REWRITE
				        		break;

			        		case(107): // ParValueCurrency
			        			if(currectInst.get("Par Value Currency").matches("")) tagValue = "#";
			        			else tagValue = "#{ExpectedEmpty().Bug(\"BSRSUP-218\", %{ParValueCurrency}, \"Protocol issues\")}"; //!!!! REWRITE
				        		break;

			        		case(109): // DelegatedReport
			        			if(originalTCR.get("AssistedReportAPA").matches("")) tagValue = "0";
			        			if(originalTCR.get("AssistedReportAPA").matches("1")) tagValue = "1";
			        			if(originalTCR.get("AssistedReportAPA").matches("2")) tagValue = "0";
				        		break;

			        		case(111): // OnBookTradingMode
			        			if(originalTCR.get("Side1").matches("8") && !originalTCR.get("VenueType").matches("O")){
			        				if(cell6.getStringCellValue().contains("_e") || cell6.getStringCellValue().contains("_c")){
			        					switch(originalTCR.get("TradingSessionSubId")){
			        						case("3"): tagValue = "5"; break;
			        					}
			        				}
			        			}
			        			else tagValue = "#";
				        		break;

			        		case(115): // TradeDetails
//			        			System.out.println("TradeDetails");
//			        			System.out.println("TrdType: " + originalTCR.get("TrdType"));
			        			switch(originalTCR.get("TrdType")){
				        			case(""):
				        			case("0"): tagValue = "0"; break;
				        			case("2"): tagValue = "3"; break;
				        			case("62"): tagValue = "1"; break;
				        			case("65"): tagValue = "2"; break;
			        			}
			        			break;

			        		case(121): // UnitQuantity
			        			String UnitQuantity = eElement2.getElementsByTagName("original").item(0).getTextContent();
			        			String[] UnitQuantityArr = UnitQuantity.split(";");
			        			if(currectInst.get("Denominated Par Value").matches("")){
								switch(currectInst.get("Sub Asset Class")){
									case("Sovereign Bonds"):
									case("Other Public Bonds"):
									case("Convertible Bonds"):
									case("Covered Bonds"):
									case("Corporate Bonds"):
									case("Other Bonds"): tagValue = UnitQuantityArr[1] + " / 100"; break;
									/*case("Shares"): 
										if(currectInst.get("MIFIR SubClass Identifier").matches("")) tagValue = UnitQuantityArr[1];
										else tagValue = "#{Expected(" + UnitQuantityArr[1] + ").Bug(\"#30899\", " + UnitQuantityArr[1] + " / 100B)}";
										break;*/
									default: tagValue = UnitQuantityArr[1]; break;
								}
			        			} else {
								switch(currectInst.get("Sub Asset Class")){
									case("Sovereign Bonds"):
									case("Other Public Bonds"):
									case("Convertible Bonds"):
									case("Covered Bonds"):
									case("Corporate Bonds"):
									case("Other Bonds"): tagValue = "#{round(" + UnitQuantityArr[1] + " / " + currectInst.get("Denominated Par Value") + ", 2)}"; break;
									//case("Shares"): tagValue = "#{Expected(" + UnitQuantityArr[1] + ").Bug(\"#30899\", " + UnitQuantityArr[1] + " / 100B)}"; break;
									default: tagValue = UnitQuantityArr[1]; break;
								}
			        				//tagValue = UnitQuantityArr[1] + " / " + currectInst.get("Denominated Par Value");
			        			}
			        			break;

			        		case(123): // DenominatedParValue
			        			//System.out.println("DenominatedParValue");
			        			//System.out.println(currectInst.get("Denominated Par Value"));
			        			//System.out.println(currectInst.get("Sub Asset Class"));
			        			if(currectInst.get("Denominated Par Value").matches("")){
								switch(currectInst.get("Sub Asset Class")){
									case("Sovereign Bonds"):
									case("Other Public Bonds"):
									case("Convertible Bonds"):
									case("Covered Bonds"):
									case("Corporate Bonds"):
									case("Other Bonds"): tagValue = "100"; break;
									/*case("Shares"): 
										if(currectInst.get("MIFIR SubClass Identifier").matches("")) tagValue = "1";
										else tagValue = "#{Expected(1B).Bug(\"#30899\", 100B)}"; 
										break;*/
									default: tagValue = "1"; break;
								}
			        				/*if(currectInst.get("Sub Asset Class").matches("Shares")) tagValue = "1";
												tagValue = "#{Expected(1B).Bug(\"#30899\", 100B)}";
			        				else tagValue = "100";*/
			        			} else {
								switch(currectInst.get("Sub Asset Class")){
									case("Sovereign Bonds"):
									case("Other Public Bonds"):
									case("Convertible Bonds"):
									case("Covered Bonds"):
									case("Corporate Bonds"):
									case("Other Bonds"): tagValue = currectInst.get("Denominated Par Value"); break;
									//case("Shares"): tagValue = "#{Expected(1B).Bug(\"#30899\", 100B)}"; break;
									default: tagValue = "1"; break;
								}
			        				//tagValue = currectInst.get("Denominated Par Value");
			        			}
			        			break;
			        	}
			        }

					//IF REJECT
					if(isReject && cell6.getStringCellValue().contains("_a")){
						switch(j){
							case(31): tagValue = "#"; break; // ReferencePrice
							case(35): tagValue = "2"; break; // EventType
							case(49): tagValue = originalTCR.get("ClearingIntention").matches("1") ? "1" : "0"; break; // TransactionToBeCleared !!!!!
							case(50): tagValue = "1"; break; // MatchStatus
							case(53): tagValue = "#"; break; // PublicationPending
							case(54): tagValue = "#"; break; // PublishIndicator
							case(85): tagValue = "(#{diffDateTime(${" + originalTCR.get("Reference") + "_a.header.SendingTime}, " +
									"#{toDateTime(x, \"yyyy-MM-dd'T'HH:mm:ss."
									+ "SSSSSS'Z'\")}, \"s\")} >= 0) && (#{diffDateTime(${" + originalTCR.get("Reference") + "_a.header.SendingTime}, " +
											"#{toDateTime(x, \"yyyy-MM-dd'T'HH:"
									+ "mm:ss.SSSSSS'Z'\")}, \"s\")} <= 1) "; break; // ReportedTime
							case(89): tagValue = "*"; break; // TradeMatchID
							case(90): tagValue = "#"; break; // TradeReportID
							case(100): tagValue = originalTCR.get("PxQtyReviewed").matches("Y") ? "1" : "0"; break; //PriceQuantityReviewed
						}
					}

			        //CHECK FOR PERSISTENCE
			        if(isPersistence){
			        	if(tagValue.contains("si.")) tagValue = tagValue.replace("si.", persistenceName + ".");
			        	if(tagValue.contains("si_a.")) tagValue = tagValue.replace("si_a.", persistenceName + "Ack.");
			        	if(tagValue.contains("si_e.")) tagValue = tagValue.replace("si_e.", persistenceName + ".");
			        	if(tagValue.contains("si_dss_a.")) tagValue = tagValue.replace("si_dss_a.", persistenceName + "RTF.");

			        	if(tagValue.contains("su.")) tagValue = tagValue.replace("su.", persistenceName + ".");
			        	if(tagValue.contains("su_a.")) tagValue = tagValue.replace("su_a.", persistenceName + "Ack.");
			        	if(tagValue.contains("su_e.")) tagValue = tagValue.replace("su_e.", persistenceName + ".");
			        	if(tagValue.contains("su_dss_a.")) tagValue = tagValue.replace("su_dss_a.", persistenceName + "RTF.");

			        	if(tagValue.contains("ssd.")) tagValue = tagValue.replace("ssd.", persistenceName + ".");
			        	if(tagValue.contains("ssd_a.")) tagValue = tagValue.replace("ssd_a.", persistenceName + "Ack.");
			        	if(tagValue.contains("ssd_e.")) tagValue = tagValue.replace("ssd_e.", persistenceName + ".");

			        	if(tagValue.contains("smd.")) tagValue = tagValue.replace("smd.", persistenceName + ".");
			        	if(tagValue.contains("smd_a.")) tagValue = tagValue.replace("smd_a.", persistenceName + "Ack.");
			        	if(tagValue.contains("smd_e.")) tagValue = tagValue.replace("smd_e.", persistenceName + ".");

					//if(tagValue.contains("psd_e.header.")) tagValue = tagValue.replace("psd_e.header.", persistenceName + ".header."); // transactTime 2.10.5
					//if(tagValue.contains("pmd_e.header.")) tagValue = tagValue.replace("pmd_e.header.", persistenceName + ".header."); // transactTime 2.10.6
					//System.out.println("tagValue: " + tagValue);
			        }

			        //System.out.println(originalTCR.get("instPrefix"));
			        //CHECK FOR PREFIX
			        if(!originalTCR.get("instPrefix").matches("")){
			        	switch(j){
			        		case(16): tagValue = tagValue.replaceAll("Origin", "Origin" + originalTCR.get("instPrefix")); break;
			        		case(22): tagValue = tagValue.replaceAll("DenominatedParValue", "DenominatedParValue" + originalTCR.get("instPrefix")); break;
			        		case(31): tagValue = tagValue.replaceAll("ReferencePrice", "ReferencePrice" + originalTCR.get("instPrefix")); break;
			        		case(33): tagValue = tagValue.replaceAll("StrikePrice", "StrikePrice" + originalTCR.get("instPrefix")); break;
			        		case(34): tagValue = tagValue.replaceAll("Yield", "Yield" + originalTCR.get("instPrefix")); break;
			        		case(47): tagValue = tagValue.replaceAll("LSEGClearingType", "LSEGClearingType" + originalTCR.get("instPrefix")); break;
			        		case(55): tagValue = tagValue.replaceAll("SecurityType", "SecurityType" + originalTCR.get("instPrefix")); break;
			        		case(61): tagValue = tagValue.replaceAll("CFICode", "CFICode" + originalTCR.get("instPrefix")); break;
			        		case(69): tagValue = tagValue.replaceAll("Currency", "Currency" + originalTCR.get("instPrefix")); break;
			        		case(75): tagValue = tagValue.replaceAll("ExpirationDate", "ExpirationDate" + originalTCR.get("instPrefix")); break;
			        		case(77): tagValue = "%{Instrument" + originalTCR.get("instPrefix") + "}"; break;
			        		case(78): tagValue = tagValue.replaceAll("InstrumentSource", "InstrumentSource" + originalTCR.get("instPrefix")); break;
			        		case(79): tagValue = tagValue.replaceAll("InstrumentCurrency", "InstrumentCurrency" + originalTCR.get("instPrefix")); break;
			        		case(84): tagValue = tagValue.replaceAll("Segment", "Segment" + originalTCR.get("instPrefix")); break;
			        		case(87): tagValue = tagValue.replaceAll("InstrumentISIN", "InstrumentISIN" + originalTCR.get("instPrefix")); break;
			        		case(88): tagValue = tagValue.replaceAll("Symbol", "Symbol" + originalTCR.get("instPrefix")); break;
			        		case(92): tagValue = tagValue.replaceAll("Underlying", "Underlying" + originalTCR.get("instPrefix")); break;
			        		case(93): tagValue = tagValue.replaceAll("MarketCode", "MarketCode" + originalTCR.get("instPrefix")); break;
			        		case(95): tagValue = tagValue.replaceAll("MarketSource", "MarketSource" + originalTCR.get("instPrefix")); break;
			        		case(99): tagValue = tagValue.replaceAll("EmissionAllowanceType", "EmissionAllowanceType" + originalTCR.get("instPrefix")); break;
			        		case(101): tagValue = tagValue.replaceAll("InstrumentStatus", "InstrumentStatus" + originalTCR.get("instPrefix")); break;
			        		case(102): tagValue = tagValue.replaceAll("MarketCode", "MarketCode" + originalTCR.get("instPrefix")); break;
			        		case(103): tagValue = tagValue.replaceAll("ShortName", "ShortName" + originalTCR.get("instPrefix")); break;
			        		case(104): tagValue = tagValue.replaceAll("ADT", "ADT" + originalTCR.get("instPrefix")); break;
			        		case(105): tagValue = tagValue.replaceAll("ClosingPrice", "ClosingPrice" + originalTCR.get("instPrefix")); break;
			        		case(106): tagValue = tagValue.replaceAll("ParValue", "ParValue" + originalTCR.get("instPrefix")); break;
			        		case(107): tagValue = tagValue.replaceAll("ParValueCurrency", "ParValueCurrency" + originalTCR.get("instPrefix")); break;
			        		case(108): tagValue = tagValue.replaceAll("NMS", "NMS" + originalTCR.get("instPrefix")); break;
			        		case(120): tagValue = tagValue.replaceAll("Equity", "Equity" + originalTCR.get("instPrefix")); break;
			        		default: break;
			        	}
			        }
			        Cell tempMessageCell = row.createCell(j, Cell.CELL_TYPE_STRING);
					tempMessageCell.setCellValue(tagValue);
					tempMessageCell.setCellStyle(style);
				}
			}
		}
	}

	public void newFIXHeaders() throws IOException{
		//LOAD FROM XLS
		//FileInputStream information = new FileInputStream("c:\\Users\\user\\Documents\\Information.xls"); // WINDOWS
		FileInputStream information = new FileInputStream(HOME + INFORMATION); // LINUX
		HSSFWorkbook book = new HSSFWorkbook(information);
		HSSFSheet infoSheet = book.getSheet("messages2");
		int infoSheetLastRow = infoSheet.getLastRowNum();

		String[] tcrcHeader = null, tcraHeader = null, tcrsHeader = null, rtfHeader = null;

		for(int i = 0; i < infoSheetLastRow; i++){
			Row row = infoSheet.getRow(i);
			LinkedHashSet<String> tmpSet = new LinkedHashSet<String>();
			switch(row.getCell(0).getStringCellValue()){
				case("TCR-C"):
					System.out.println("TCR-C");
					for(int j = 1; j < row.getLastCellNum(); j++)
						if(row.getCell(j).getCellType() != Cell.CELL_TYPE_BLANK){
							String value = row.getCell(j).getStringCellValue();
							if(value.matches("SecurityIDSource")) value = "Instrument";
							if(value.matches("SecurityID")) continue;
							if(value.matches("CountryOfIssue")) continue;
							tmpSet.add(value);
						}
					//System.out.println("Tag: " + row.getCell(j).getStringCellValue());
					//System.out.println("tmpSet: " + tmpSet);
					tcrcHeader = tmpSet.toArray(new String[0]);
					//System.out.println("tcrcHeader: " + Arrays.asList(tcrcHeader));
					break;
				case("TCR-Ack"):
					System.out.println("TCR-Ack");
					for(int j = 1; j < row.getLastCellNum(); j++)
						if(row.getCell(j).getCellType() != Cell.CELL_TYPE_BLANK){
							String value = row.getCell(j).getStringCellValue();
							if(value.matches("SecurityIDSource")) value = "Instrument";
							if(value.matches("SecurityID")) continue;
							if(value.matches("CountryOfIssue")) continue;
							tmpSet.add(value);
						}
					//System.out.println("Tag: " + row.getCell(j).getStringCellValue());
					//System.out.println("tmpSet: " + tmpSet);
					tcraHeader = tmpSet.toArray(new String[0]);
					//System.out.println("tcraHeader: " + Arrays.asList(tcraHeader));
					break;
				case("TCR-S"):
					System.out.println("TCR-S");
					for(int j = 1; j < row.getLastCellNum(); j++)
						if(row.getCell(j).getCellType() != Cell.CELL_TYPE_BLANK){
							String value = row.getCell(j).getStringCellValue();
							if(value.matches("SecurityIDSource")) value = "Instrument";
							if(value.matches("SecurityID")) continue;
							if(value.matches("CountryOfIssue")) continue;
							tmpSet.add(value);
						}
					//System.out.println("Tag: " + row.getCell(j).getStringCellValue());
					//System.out.println("tmpSet: " + tmpSet);
					tcrsHeader = tmpSet.toArray(new String[0]);
					//System.out.println("tcrsHeader: " + Arrays.asList(tcrsHeader));
					break;

				case("TTC"):
					System.out.println("TTC");
					for(int j = 1; j < row.getLastCellNum(); j++)
						if(row.getCell(j).getCellType() != Cell.CELL_TYPE_BLANK){
							String value = row.getCell(j).getStringCellValue();
							if(value.matches("SecurityIDSource")) value = "Instrument";
							if(value.matches("SecurityID")) continue;
							if(value.matches("CountryOfIssue")) continue;
							tmpSet.add(value);
						}
					//System.out.println("Tag: " + row.getCell(j).getStringCellValue());
					//System.out.println("tmpSet: " + tmpSet);
					rtfHeader = tmpSet.toArray(new String[0]);
					System.out.println("rtfHeader: " + Arrays.asList(rtfHeader));
					break;
				default: break;
			}
		}


		System.out.println("newHeaders");
		log.write("\nnewHeaders Func\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();

		String[] tcrcArr = { "si", "su", "ssd", "smd", "ci", "cu", "csda", "cmda", "csdb", "cmdb", "pri", "pru", "prsda", "prmda", "prsd",
				"prmd", "ai", "au", "asdb", "amdb", "asda", "amda" };
		HashSet<String> tcrcSet = new HashSet<String>(Arrays.asList(tcrcArr));

		String[] tcraArr = { "si_a", "su_a", "ssd_a", "smd_a", "ci_a", "cu_a", "csdb_a", "cmdb_a", "csda_a", "cmda_a", "pri_a", "pru_a",
				"prsda_a", "prmda_a", "prsd_a", "prmd_a", "ai_a", "au_a", "asdb_a", "amdb_a", "asda_a", "amda_a" };
		HashSet<String> tcraSet = new HashSet<String>(Arrays.asList(tcraArr));

		String[] tcrsArr = { "si_e", "si_c", "psd_e", "psd_c", "pmd_e", "pmd_c", "ai_e", "ai_c", "prsd_e", "prsd_c", "prmd_e", "prmd_c" };
		HashSet<String> tcrsSet = new HashSet<String>(Arrays.asList(tcrsArr));

		String[] mergetcrsArr = { "su_e", "ssd_e", "smd_e", "ci_e", "cu_e", "csda_e", "cmda_e", "csdb_e", "cmdb_e",
				"au_e", "asdb_e", "amdb_e", "asda_e", "amda_e", "prmd_e", "prsd_e" };

		HashSet<String> mergetcrsSet = new HashSet<String>(Arrays.asList(mergetcrsArr));


		String[] rtfMessages = { "si_dss_a", "su_dss_a", "ssd_dss_a", "smd_dss_a", "psd_dss_e", "pmd_dss_e", "ci_dss_a", "cu_dds_a", "csdb_dss_a", "smdb_dss_a", "csda_dss_a",
					 "cmda_dss_a","pri_dss_a", "pru_dss_a", "prsda_dss_a", "prmda_dss_a", "prsd_dss_a", "prmd_dss_a", "ai_dss_a", "au_dss_a", "asdb_dss_a", "amdb_dss_a", 
					"asda_dss_a", "amda_dss_a" };

		HashSet<String> rtfMessagesSet = new HashSet<String>(Arrays.asList(rtfMessages));

		boolean flag = false;
		for(int i = 0; i < lastRow; i++){
			Row row = sheet.getRow(i);
			if(row == null) continue;

			Cell cell8 = row.getCell(8);
			if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING) continue;

			if(!flag && cell8.getStringCellValue().matches("DefineHeader")){ flag = true; continue; }

			if(flag){
				Cell cell6 = row.getCell(6);
				if(cell6 == null || cell6.getCellType() != Cell.CELL_TYPE_STRING) continue;


				//IF RTF
				if(rtfMessagesSet.contains(cell6.getStringCellValue())){
					System.out.println("INSIDE RTF");

					Row headerRow = sheet.getRow(i - 1);
//					int lastCell = headerRow.getLastCellNum();
//					LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();

					/*for(int j = 14; j < lastCell; j++){
						//Collect old message
						Cell headerCell = headerRow.getCell(j);
						if(headerCell == null || headerCell.getCellType() != Cell.CELL_TYPE_STRING) continue;

						Cell messageCell = row.getCell(j);
						String messageCellContent = "";

						if(messageCell == null){ map.put(headerCell.getStringCellValue(), ""); continue; }

						switch(messageCell.getCellType()){
						case(Cell.CELL_TYPE_BLANK): break;
						case(Cell.CELL_TYPE_BOOLEAN): System.out.println(messageCell.getBooleanCellValue()); messageCellContent = Boolean.toString(messageCell.getBooleanCellValue()); break;
						case(Cell.CELL_TYPE_ERROR): break;
						case(Cell.CELL_TYPE_FORMULA):  System.out.println("Formula"); break;
						case(Cell.CELL_TYPE_NUMERIC):
							if(messageCell.getNumericCellValue() % 1 == 0){
								Integer tmp = (int)messageCell.getNumericCellValue();
								messageCellContent = Integer.toString(tmp);
							} else messageCellContent = Double.toString(messageCell.getNumericCellValue());
						break;
						case(Cell.CELL_TYPE_STRING): messageCellContent = messageCell.getStringCellValue(); break;
						}

						map.put(headerCell.getStringCellValue(), messageCellContent);
						headerRow.removeCell(headerCell);
						row.removeCell(messageCell);
					}

					System.out.println("RTF: " + map);*/

					//Replace header
					CellStyle headerStyle = sheet.getRow(i - 1).getCell(8).getCellStyle();
//					CellStyle messageStyle = sheet.getRow(i).getCell(8).getCellStyle();
					for(int j = 14, k = 0; j < 14 + rtfHeader.length; j++, k++){
						//System.out.println(rtfHeader[k]);
						Cell headerCell = headerRow.getCell(j);
						if(headerCell == null) { headerRow.createCell(j, Cell.CELL_TYPE_STRING); }

						headerRow.getCell(j).setCellValue(rtfHeader[k]);
						headerRow.getCell(j).setCellStyle(headerStyle);

						//Cell messageCell = row.createCell(j, Cell.CELL_TYPE_STRING);
						/*messageCell.setCellStyle(messageStyle);
						if(map.containsKey(rtfHeader[k])){
							messageCell.setCellValue(map.get(rtfHeader[k]));
						}else messageCell.setCellValue("");*/
					}
					flag = false;
					continue;
				}


				//IF TCR-C
				if(tcrcSet.contains(cell6.getStringCellValue())){
					Row headerRow = sheet.getRow(i - 1);
					int lastCell = headerRow.getLastCellNum();

					LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();

					for(int j = 14; j < lastCell; j++){
						//Collect
						Cell headerCell = headerRow.getCell(j);
						if(headerCell == null || headerCell.getCellType() != Cell.CELL_TYPE_STRING) continue;

						Cell messageCell = row.getCell(j);
						String messageCellContent = "";

						if(messageCell == null){ map.put(headerCell.getStringCellValue(), ""); continue; }

						switch(messageCell.getCellType()){
						case(Cell.CELL_TYPE_BLANK): break;
						case(Cell.CELL_TYPE_BOOLEAN): System.out.println(messageCell.getBooleanCellValue()); messageCellContent = Boolean.toString(messageCell.getBooleanCellValue()); break;
						case(Cell.CELL_TYPE_ERROR): break;
						case(Cell.CELL_TYPE_FORMULA):  System.out.println("Formula"); break;
						case(Cell.CELL_TYPE_NUMERIC):
							if(messageCell.getNumericCellValue() % 1 == 0){
								Integer tmp = (int)messageCell.getNumericCellValue();
								messageCellContent = Integer.toString(tmp);
							} else messageCellContent = Double.toString(messageCell.getNumericCellValue());
						break;
						case(Cell.CELL_TYPE_STRING): messageCellContent = messageCell.getStringCellValue(); break;
						}

						map.put(headerCell.getStringCellValue(), messageCellContent);
						headerRow.removeCell(headerCell);
						row.removeCell(messageCell);
					}

					//System.out.println("TCR-C: " + map);

					//Replace
					CellStyle headerStyle = sheet.getRow(i - 1).getCell(8).getCellStyle();
					CellStyle messageStyle = sheet.getRow(i).getCell(8).getCellStyle();
					for(int j = 14, k = 0; j < 14 + tcrcHeader.length; j++, k++){
						Cell headerCell = headerRow.getCell(j);
						if(headerCell == null) { headerRow.createCell(j, Cell.CELL_TYPE_STRING); }

						headerRow.getCell(j).setCellValue(tcrcHeader[k]);
						headerRow.getCell(j).setCellStyle(headerStyle);

						Cell messageCell = row.createCell(j, Cell.CELL_TYPE_STRING);
						messageCell.setCellStyle(messageStyle);
						if(map.containsKey(tcrcHeader[k])){
							messageCell.setCellValue(map.get(tcrcHeader[k]));
						}else messageCell.setCellValue("");
					}
					flag = false;
					continue;
				}


				//IF TCR-A
				if(tcraSet.contains(cell6.getStringCellValue())){
					System.out.println("BusinessMessageReject");
					Cell cell9 = row.getCell(9);
					if(cell9 != null){
						String cell9Content = cell9.getStringCellValue();
						System.out.println(cell9Content);
						switch(cell9Content){
							case("BusinessMessageReject"): continue;
							case("Reject"): continue;
						}
					}
					System.out.println("AFTER");


					Row headerRow = sheet.getRow(i - 1);
					int lastCell = headerRow.getLastCellNum();

					LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();

					for(int j = 14; j < lastCell; j++){
						//Collect
						Cell headerCell = headerRow.getCell(j);
						if(headerCell == null || headerCell.getCellType() != Cell.CELL_TYPE_STRING) continue;

						Cell messageCell = row.getCell(j);
						String messageCellContent = "";

						if(messageCell == null){ map.put(headerCell.getStringCellValue(), ""); }

						switch(messageCell.getCellType()){
						case(Cell.CELL_TYPE_BLANK): break;
						case(Cell.CELL_TYPE_BOOLEAN): messageCellContent = Boolean.toString(messageCell.getBooleanCellValue()); break;
						case(Cell.CELL_TYPE_ERROR): break;
						case(Cell.CELL_TYPE_FORMULA): break;
						case(Cell.CELL_TYPE_NUMERIC):
							if(messageCell.getNumericCellValue() % 1 == 0){
								Integer tmp = (int)messageCell.getNumericCellValue();
								messageCellContent = Integer.toString(tmp);
							} else messageCellContent = Double.toString(messageCell.getNumericCellValue());
						break;
						case(Cell.CELL_TYPE_STRING): messageCellContent = messageCell.getStringCellValue(); break;
						}

						map.put(headerCell.getStringCellValue(), messageCellContent);
						headerRow.removeCell(headerCell);
						row.removeCell(messageCell);
					}

					//System.out.println("TCR-ACK: " + map);

					//Replace
					CellStyle headerStyle = sheet.getRow(i - 1).getCell(8).getCellStyle();
					CellStyle messageStyle = sheet.getRow(i).getCell(8).getCellStyle();
					for(int j = 14, k = 0; j < 14 + tcraHeader.length; j++, k++){
						Cell headerCell = headerRow.getCell(j);
						if(headerCell == null) { headerRow.createCell(j, Cell.CELL_TYPE_STRING); }

						headerRow.getCell(j).setCellValue(tcraHeader[k]);
						headerRow.getCell(j).setCellStyle(headerStyle);

						Cell messageCell = row.createCell(j, Cell.CELL_TYPE_STRING);
						messageCell.setCellStyle(messageStyle);
						if(map.containsKey(tcraHeader[k])){
							messageCell.setCellValue(map.get(tcraHeader[k]));
						}else messageCell.setCellValue("");
					}
					flag = false;
					continue;
				}


				//IF TCR-S
				if(tcrsSet.contains(cell6.getStringCellValue())){
					Row headerRow = sheet.getRow(i - 1);
					int lastCell = headerRow.getLastCellNum();

					LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();

					for(int j = 14; j < lastCell; j++){
						//Collect
						Cell headerCell = headerRow.getCell(j);
						if(headerCell == null || headerCell.getCellType() != Cell.CELL_TYPE_STRING) continue;

						Cell messageCell = row.getCell(j);
						String messageCellContent = "";

						if(messageCell == null){ map.put(headerCell.getStringCellValue(), ""); continue; }

						switch(messageCell.getCellType()){
						case(Cell.CELL_TYPE_BLANK): break;
						case(Cell.CELL_TYPE_BOOLEAN): System.out.println("Boolean"); messageCellContent = Boolean.toString(messageCell.getBooleanCellValue()); break;
						case(Cell.CELL_TYPE_ERROR): break;
						case(Cell.CELL_TYPE_FORMULA): System.out.println("Formula"); messageCellContent = Boolean.toString(messageCell.getBooleanCellValue()); break;
						case(Cell.CELL_TYPE_NUMERIC):
							if(messageCell.getNumericCellValue() % 1 == 0){
								Integer tmp = (int)messageCell.getNumericCellValue();
								messageCellContent = Integer.toString(tmp);
							} else messageCellContent = Double.toString(messageCell.getNumericCellValue());
						break;
						case(Cell.CELL_TYPE_STRING): messageCellContent = messageCell.getStringCellValue(); break;
						}

						map.put(headerCell.getStringCellValue(), messageCellContent);
						headerRow.removeCell(headerCell);
						row.removeCell(messageCell);
					}

					//System.out.println("SINGLE TCR-S: " + map);

					//Replace
					CellStyle headerStyle = sheet.getRow(i - 1).getCell(8).getCellStyle();
					CellStyle messageStyle = sheet.getRow(i).getCell(8).getCellStyle();
					for(int j = 14, k = 0; j < 14 + tcrsHeader.length; j++, k++){
						Cell headerCell = headerRow.getCell(j);
						if(headerCell == null) { headerRow.createCell(j, Cell.CELL_TYPE_STRING); }

						headerRow.getCell(j).setCellValue(tcrsHeader[k]);
						headerRow.getCell(j).setCellStyle(headerStyle);

						//if(j < 14) continue;

						Cell messageCell = row.createCell(j, Cell.CELL_TYPE_STRING);
						messageCell.setCellStyle(messageStyle);
						if(map.containsKey(tcrsHeader[k])){
							messageCell.setCellValue(map.get(tcrsHeader[k]));
						}else messageCell.setCellValue("");
					}

					flag = false;
					continue;
				}


				//IF MERGE TCR-S
				if(mergetcrsSet.contains(cell6.getStringCellValue())){
					//System.out.println("in");

					Row headerRow = sheet.getRow(i - 1);
					int lastCell = headerRow.getLastCellNum();
					LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();
					LinkedHashMap<String, String> map2 = null;

					Row nextRow = sheet.getRow(i + 1);
					if(nextRow != null){
						Cell nextRowCell6 = nextRow.getCell(6);
						if(nextRowCell6 != null && nextRowCell6.getCellType() == Cell.CELL_TYPE_STRING
								&& !nextRowCell6.getStringCellValue().matches("")) map2 = new LinkedHashMap<String, String>();
					}

					//System.out.println("in");

					for(int j = 14; j < lastCell; j++){
						//Collect
						Cell headerCell = headerRow.getCell(j);
						if(headerCell == null || headerCell.getCellType() != Cell.CELL_TYPE_STRING) continue;

						Cell messageCell = row.getCell(j);
						String messageCellContent = "";

						if(messageCell == null){ map.put(headerCell.getStringCellValue(), ""); continue; }

						switch(messageCell.getCellType()){
						case(Cell.CELL_TYPE_BLANK): break;
						case(Cell.CELL_TYPE_BOOLEAN): messageCellContent = Boolean.toString(messageCell.getBooleanCellValue()); break;
						case(Cell.CELL_TYPE_ERROR): break;
						case(Cell.CELL_TYPE_FORMULA): messageCellContent = Boolean.toString(messageCell.getBooleanCellValue()); break;
						case(Cell.CELL_TYPE_NUMERIC):
							if(messageCell.getNumericCellValue() % 1 == 0){
								Integer tmp = (int)messageCell.getNumericCellValue();
								messageCellContent = Integer.toString(tmp);
							} else messageCellContent = Double.toString(messageCell.getNumericCellValue());
						break;
						case(Cell.CELL_TYPE_STRING): messageCellContent = messageCell.getStringCellValue(); break;
						}

						map.put(headerCell.getStringCellValue(), messageCellContent);
						row.removeCell(messageCell);


						if(map2 != null){
							Cell nextMessageCell = sheet.getRow(i + 1).getCell(j);
							String nextMessageCellContent = "";

							if(nextMessageCell == null){ map2.put(headerCell.getStringCellValue(), ""); }

							switch(nextMessageCell.getCellType()){
							case(Cell.CELL_TYPE_BLANK): break;
							case(Cell.CELL_TYPE_BOOLEAN): nextMessageCellContent = Boolean.toString(nextMessageCell.getBooleanCellValue()); break;
							case(Cell.CELL_TYPE_ERROR): break;
							case(Cell.CELL_TYPE_FORMULA): nextMessageCellContent = Boolean.toString(nextMessageCell.getBooleanCellValue()); break;
							case(Cell.CELL_TYPE_NUMERIC):
								if(nextMessageCell.getNumericCellValue() % 1 == 0){
									Integer tmp = (int)nextMessageCell.getNumericCellValue();
									nextMessageCellContent = Integer.toString(tmp);
								} else nextMessageCellContent = Double.toString(nextMessageCell.getNumericCellValue());
							break;
							case(Cell.CELL_TYPE_STRING): nextMessageCellContent = nextMessageCell.getStringCellValue(); break;
							}

							map2.put(headerCell.getStringCellValue(), nextMessageCellContent);
							sheet.getRow(i + 1).removeCell(nextMessageCell);
						}
						headerRow.removeCell(headerCell);
					}

					//System.out.println("MERGE MAP1: " + map);
					//System.out.println("MERGE MAP2: " + map2);

					//Replace
					CellStyle headerStyle = sheet.getRow(i - 1).getCell(8).getCellStyle();
					CellStyle messageStyle = sheet.getRow(i).getCell(8).getCellStyle();
					for(int j = 14, k = 0; j < 14 + tcrsHeader.length; j++, k++){
						Cell headerCell = headerRow.getCell(j);
						if(headerCell == null) { headerRow.createCell(j, Cell.CELL_TYPE_STRING); }

						headerRow.getCell(j).setCellValue(tcrsHeader[k]);
						headerRow.getCell(j).setCellStyle(headerStyle);

						Cell messageCell = row.createCell(j, Cell.CELL_TYPE_STRING);
						messageCell.setCellStyle(messageStyle);
						if(map.containsKey(tcrsHeader[k])){
							messageCell.setCellValue(map.get(tcrsHeader[k]));
						}else messageCell.setCellValue("");
					}

					if(map2 != null){
						for(int j = 14, k = 0; j < 14 + tcrsHeader.length; j++, k++){
							Cell nextMessageCell = sheet.getRow(i + 1).createCell(j, Cell.CELL_TYPE_STRING);
							nextMessageCell.setCellStyle(messageStyle);
							if(map.containsKey(tcrsHeader[k])){
								nextMessageCell.setCellValue(map2.get(tcrsHeader[k]));
							}else nextMessageCell.setCellValue("");
						}
					}
					flag = false;
					continue;
				}
			}
		}
	}

	public void mergeTCRforUnPublishedandCancel() throws IOException{
		System.out.println("removeCancelPreReleaseHeaders");
		log.write("\nremoveCancelPreReleaseHeaders Func\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();
		boolean flag = false;
		LinkedList<Integer> rows = new LinkedList<Integer>();

		String[] execArr = new String[]{ "ci_e", "cu_e", "csda_e", "cmda_e", "csdb_e", "cmdb_e", "su_e", "au_e", "ssd_e", "smd_e" };
		HashSet<String> execSet = new HashSet<String>(Arrays.asList(execArr));
		String[] contraArr = new String[]{ "ci_c", "cu_c", "csda_c", "cmda_c", "csdb_c", "cmdb_c", "su_c", "au_c" };
		HashSet<String> contraSet = new HashSet<String>(Arrays.asList(contraArr));

		for(int i = 0; i < lastRow; i++){
			Row row = sheet.getRow(i);
			if(row == null){
				System.out.println("Empty Row: " + i);
				rows.add(i);
				continue;
			}

			Cell cell6 = row.getCell(6);

			if(!flag){
				if(cell6 == null || cell6.getCellType() != Cell.CELL_TYPE_STRING) continue;
				String cell6Content = cell6.getStringCellValue();
				if(execSet.contains(cell6Content)){ flag = true; continue; }
			}

			if(flag){
				Cell cell8 = row.getCell(8);
				if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING) { rows.add(i); continue; }

				String cell8Content = cell8.getStringCellValue();
				switch(cell8Content){
					case(""): case("DefineHeader"): case("CheckMessage"): rows.add(i); continue;
				}

				if(cell8Content.matches("Sleep")){
					flag = false;
					continue;
				}

				if(cell8Content.matches("receive") && contraSet.contains(row.getCell(6).getStringCellValue())){
					flag = false;
					continue;
				}
			}
		}


		//Remove
		System.out.println("Remove Rows");
		System.out.println(rows);
		Iterator<Integer> iter = rows.descendingIterator();
		while(iter.hasNext()){
			int i = iter.next();
			//System.out.println("Remove Row: " + i);
			sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
		}
	}

	public void priceConditionsCorrection() throws IOException{ // Äîïèñàòü !!!
		System.out.println("priceConditionsCorrection");
		log.write("\npriceConditionsCorrection Func\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();

		TreeMap<String, String> map = new TreeMap<String, String>();

		for(int i = 0; i < lastRow; i++){
			Row row = sheet.getRow(i);
			if(row == null) continue;

			Cell cell9 = row.getCell(9);
			if(cell9 == null || cell9.getCellType() != Cell.CELL_TYPE_STRING) continue;

			String cell9Content = cell9.getStringCellValue();
			if(cell9Content.matches("TradeCaptureReport_NoTradePriceConditions")){
				System.out.println("IN");
				double type = sheet.getRow(i).getCell(15).getNumericCellValue();
				if(type == 13){
					System.out.println("SDIV");
					map.put(row.getCell(6).getStringCellValue(), "SDIV");
					row.getCell(6).setCellValue("SDIV");
					continue;
				}

				if(type == 14){
					System.out.println("RPRI");
					map.put(row.getCell(6).getStringCellValue(), "RPRI");
					row.getCell(6).setCellValue("RPRI");
					continue;
				}

				if(type == 15){
					System.out.println("NPFT");
					map.put(row.getCell(6).getStringCellValue(), "NPFT");
					row.getCell(6).setCellValue("NPFT");
					continue;
				}

				if(type == 16){
					System.out.println("TNCP");
					map.put(row.getCell(6).getStringCellValue(), "TNCP");
					row.getCell(6).setCellValue("TNCP");
					continue;
				}
			}

			//Replace
			if(cell9Content.matches("TradeCaptureReport") &&
					(row.getCell(8).getStringCellValue().matches("receive") || row.getCell(8).getStringCellValue().matches("send"))){
				//System.out.println("Found TradeCaptureReport");
				int lastCell = row.getLastCellNum();
				for(int j = 14; j < lastCell; j++){
					Cell tempCell = row.getCell(j);
					if(tempCell == null || tempCell.getCellType() != Cell.CELL_TYPE_STRING) continue;
					String tempCellContent = tempCell.getStringCellValue();

					for(String key : map.descendingKeySet()){
						if(tempCellContent.contains(key)){
							//System.out.println("FOUND GROUP!!!!!!!!!!!!!!");
							//System.out.println("tempCellContent: " + tempCellContent + ", Key: " + key + ", value: " + map.get(key));
							tempCellContent = tempCellContent.replace(key, map.get(key));
							//System.out.println("tempCellContent after: " + tempCellContent);
							tempCell.setCellValue(tempCellContent);
						}
					}
				}
			}
		}
		System.out.println(map);
	}

	public void noPartyNamesCorrection() throws IOException{
		System.out.println("noPartyCorrection");
		log.write("\nnoPartyCorrection Func\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();
		for(int i = 0; i < lastRow; i++){
			Row row = sheet.getRow(i);
			if(row == null) continue;
			Cell cell9 = row.getCell(9);
			if(cell9 == null || cell9.getCellType() != Cell.CELL_TYPE_STRING) continue;
			String cell9Content = cell9.getStringCellValue();

			if(cell9Content.matches("TradeCaptureReport_NoSides_NoPartyIDs")){
				Cell cell6 = row.getCell(6);
				if(cell6 == null || cell6.getCellType() != Cell.CELL_TYPE_STRING) continue;
				String cell6Content = cell6.getStringCellValue();
				cell6.setCellValue(cell6Content.replace("SRR", ""));
				continue;
			}

			if(cell9Content.matches("TradeCaptureReport_NoSides")){
				Cell cell18 = row.getCell(18);
				if(cell18 == null || cell18.getCellType() != Cell.CELL_TYPE_STRING) continue;
				String cell18Content = cell18.getStringCellValue();
				cell18.setCellValue(cell18Content.replaceAll("SRR", ""));
				continue;
			}
		}
	}

	public void leiNamesCorrection() throws IOException{
		System.out.println("leiCorrection");
		log.write("\nleiCorrection Func\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();
		for(int i = 0; i < lastRow; i++){
			Row row = sheet.getRow(i);
			if(row == null) continue;
			Cell cell14 = row.getCell(14);
			if(cell14 == null || cell14.getCellType() != Cell.CELL_TYPE_STRING) continue;
			if(cell14.getStringCellValue().matches("\\%\\{LEI\\}")){
				cell14.setCellValue("%{ExecuterLEI}");
			}
		}
	}

	public void correctTestNumbers() throws IOException{
		System.out.println("correctTestNumbers");
		log.write("\ncorrectTestNumbers Func\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();
		int count = 1;

		for(int i = 0; i < lastRow; i++){
			Row row = sheet.getRow(i);
			if(row == null) continue;
			Cell cell8 = row.getCell(8);
			if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING) continue;
			String cell8Content = cell8.getStringCellValue();
			if(cell8Content.matches("test case start")){
				Cell cell0 = row.getCell(0);
				if(cell0 == null || cell0.getCellType() != Cell.CELL_TYPE_STRING) continue;
				cell0.setCellValue("test" + count++);
			}
		}
	}

	public void correctFlagNames() throws IOException{
		System.out.println("correctFlagNames");
		log.write("\ncorrectFlagNames\n");

		TreeMap<String, String> map = new TreeMap<String, String>();

		HSSFSheet sheet = doc.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();

		for(int i = 0; i < lastRow; i++){
			Row row = sheet.getRow(i);
			if(row == null) continue;

			Cell cell9 = row.getCell(9);
			if(cell9 == null || cell9.getCellType() != Cell.CELL_TYPE_STRING) continue;

			String cell9Content = cell9.getStringCellValue();
			if(cell9Content.matches("TradeCaptureReport_NoTrdRegPublications")){
				double type = sheet.getRow(i).getCell(15).getNumericCellValue();
				double reason = sheet.getRow(i).getCell(16).getNumericCellValue();
				//System.out.println(str1 + ", " + str2);

				if(type == 0.0 && reason == 0.0) {
					System.out.println("NLIQ");
					map.put(row.getCell(6).getStringCellValue(), "NLIQ");
					row.getCell(6).setCellValue("NLIQ");
					continue; }
				if(type == 0.0 && reason == 1.0) {
					System.out.println("OLIQ");
					map.put(row.getCell(6).getStringCellValue(), "OLIQ");
					row.getCell(6).setCellValue("OLIQ");
					continue; }
				if(type == 0.0 && reason == 2.0) {
					System.out.println("PRIC");
					map.put(row.getCell(6).getStringCellValue(), "PRIC");
					row.getCell(6).setCellValue("PRIC");
					continue; }
				if(type == 0.0 && reason == 3.0) {
					System.out.println("RFPT");
					map.put(row.getCell(6).getStringCellValue(), "RFPT");
					row.getCell(6).setCellValue("RFPT");
					continue; }
				if(type == 0.0 && reason == 4.0) {
					System.out.println("PRE_ILQD");
					map.put(row.getCell(6).getStringCellValue(), "PRE_ILQD");
					row.getCell(6).setCellValue("PRE_ILQD");
					continue; }
				if(type == 0.0 && reason == 5.0) {
					System.out.println("PRE_SIZE");
					map.put(row.getCell(6).getStringCellValue(), "PRE_SIZE");
					row.getCell(6).setCellValue("PRE_SIZE");
					continue; }
				if(type == 1.0 && reason == 6.0) {
					System.out.println("DEF_LRGS");
					map.put(row.getCell(6).getStringCellValue(), "DEF_LRGS");
					row.getCell(6).setCellValue("DEF_LRGS");
					continue; }
				if(type == 1.0 && reason == 7.0) {
					System.out.println("DEF_ILQD");
					map.put(row.getCell(6).getStringCellValue(), "DEF_ILQD");
					row.getCell(6).setCellValue("DEF_ILQD");
					continue; }
				if(type == 1.0 && reason == 8.0) {
					System.out.println("DEF_SIZE");
					map.put(row.getCell(6).getStringCellValue(), "DEF_SIZE");
					row.getCell(6).setCellValue("DEF_SIZE");
					continue; }
			}

			//Replace
			if(cell9Content.matches("TradeCaptureReport") &&
					(row.getCell(8).getStringCellValue().matches("receive") || row.getCell(8).getStringCellValue().matches("send"))){
				//System.out.println("Found TradeCaptureReport");
				int lastCell = row.getLastCellNum();
				for(int j = 14; j < lastCell; j++){
					Cell tempCell = row.getCell(j);
					if(tempCell == null || tempCell.getCellType() != Cell.CELL_TYPE_STRING) continue;
					String tempCellContent = tempCell.getStringCellValue();

					for(String key : map.descendingKeySet()){
						if(tempCellContent.contains(key)){
							//System.out.println("FOUND GROUP!!!!!!!!!!!!!!");
							//System.out.println("tempCellContent: " + tempCellContent + ", Key: " + key + ", value: " + map.get(key));
							tempCellContent = tempCellContent.replace(key, map.get(key));
							//System.out.println("tempCellContent after: " + tempCellContent);
							tempCell.setCellValue(tempCellContent);
							//break;
							//if(key.matches("LRGS")) map.remove(key);
						}
					}
				}
			}
		}
		System.out.println(map);
	}

	public void correctFixCounts() throws IOException{ //íîâàÿ âåðñèÿ
		System.out.println("correctFixCounts");
		log.write("\ncorrectFixCounts Func\n");

		HSSFSheet sheet = doc.getSheetAt(0);
		boolean flag = false;
		int count = 0, apa2Count = 0;

		CharSequence charSeq = "";

		for(int i = 0; i < sheet.getLastRowNum(); i++){
			System.out.println("Excel: " + (i + 1) + ", COUNT: " + count + ", FLAG: " + flag);
			Row row = sheet.getRow(i);
			if(row == null) {
				sheet.createRow(i);
				row = sheet.getRow(i);
			}

			Cell cell8 = row.getCell(8);
			if(!flag && (cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING)) continue;

			if(!flag && cell8.getStringCellValue().matches("send")){ charSeq = row.getCell(4).getStringCellValue(); continue; }

			if(flag && (cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING || !cell8.getStringCellValue().contains("count"))){
				if(count > 4){ flag = false; count = 0; apa2Count = 0; continue; }
				if(count == 4){
					Row tempRow = sheet.getRow(i - 1);
					Cell cell4 = tempRow.getCell(4);
					Cell cell10 = tempRow.getCell(10);

					String service = cell4.getStringCellValue();
					String checkPoint = cell10.getStringCellValue();

					//System.out.println(cell4.getStringCellValue() + ", " + cell10.getStringCellValue());

					sheet.shiftRows(i, sheet.getLastRowNum(), 2);

					Row tmpRow = sheet.createRow(i - 1);
					tmpRow.createCell(0, Cell.CELL_TYPE_STRING).setCellValue("0");
					tmpRow.createCell(4, Cell.CELL_TYPE_STRING).setCellValue(service);
					tmpRow.createCell(7, Cell.CELL_TYPE_STRING).setCellValue("0");
					tmpRow.createCell(8, Cell.CELL_TYPE_STRING).setCellValue("count");
					tmpRow.createCell(9, Cell.CELL_TYPE_STRING).setCellValue("TradeCaptureReport");
					tmpRow.createCell(10, Cell.CELL_TYPE_STRING).setCellValue(checkPoint);
					tmpRow.createCell(11, Cell.CELL_TYPE_STRING).setCellValue(Integer.toString(apa2Count));


					tmpRow = sheet.createRow(i);
					tmpRow.createCell(0, Cell.CELL_TYPE_STRING).setCellValue("0");
					tmpRow.createCell(4, Cell.CELL_TYPE_STRING).setCellValue(service);
					tmpRow.createCell(7, Cell.CELL_TYPE_STRING).setCellValue("0");
					tmpRow.createCell(8, Cell.CELL_TYPE_STRING).setCellValue("count");
					tmpRow.createCell(9, Cell.CELL_TYPE_STRING).setCellValue("TradeCaptureReportAck");
					tmpRow.createCell(10, Cell.CELL_TYPE_STRING).setCellValue(checkPoint);
					tmpRow.createCell(11, Cell.CELL_TYPE_STRING).setCellValue("0");

					tmpRow = sheet.createRow(i + 1);
					tmpRow.createCell(0, Cell.CELL_TYPE_STRING).setCellValue("0");
					tmpRow.createCell(4, Cell.CELL_TYPE_STRING).setCellValue(service);
					tmpRow.createCell(7, Cell.CELL_TYPE_STRING).setCellValue("0");
					tmpRow.createCell(8, Cell.CELL_TYPE_STRING).setCellValue("countApp");
					tmpRow.createCell(9, Cell.CELL_TYPE_STRING).setCellValue("");
					tmpRow.createCell(10, Cell.CELL_TYPE_STRING).setCellValue(checkPoint);
					tmpRow.createCell(11, Cell.CELL_TYPE_STRING).setCellValue(Integer.toString(apa2Count));

					//System.out.println("APA2Counts: " + apa2Count);
					flag = false;
					count = 0;
					apa2Count = 0;
					i += 2;
					continue;
				}
			}

			if(sheet.getRow(i).getCell(8) == null) { continue; }
			String cell8Content = sheet.getRow(i).getCell(8).getStringCellValue();

			//apa counts
			if(!flag && (cell8Content.matches("receive") && sheet.getRow(i).getCell(4).getStringCellValue().matches("apa2"))){
				apa2Count++;
			}

			Cell tempCell4 = sheet.getRow(i).getCell(4);

			if(!flag && (cell8Content.contains("count") && (tempCell4.getStringCellValue().contains("apa") ||
					tempCell4.getStringCellValue().contains(charSeq)))){ flag = true; count++; continue; }
			if(flag && (cell8Content.contains("count") && (tempCell4.getStringCellValue().contains("apa") ||
					tempCell4.getStringCellValue().contains(charSeq)))) { count++; continue; }
			if(flag && !cell8Content.contains("count")){
				if(count > 4){ flag = false; count = 0; apa2Count = 0; continue; }
				if(count == 4){
					flag = false;
					count = 0;
					apa2Count = 0;

//					Row tempRow = sheet.getRow(i - 1);
					i += 2;
					continue;
				}
			}
		}
	}

	public void correctFixCounts14032018() throws IOException{
		System.out.println("correctFixCounts");
		log.write("\ncorrectFixCounts Func\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		boolean flag = false;
		int count = 0;

		for(int i = 0; i < sheet.getLastRowNum(); i++){
			System.out.println("Excel: " + (i + 1) + ", COUNT: " + count + ", FLAG: " + flag);
			Row row = sheet.getRow(i);

			if(row == null) {
				System.out.println("Row NULL");
				continue;
			}

			Cell cell8 = row.getCell(8);
			if(cell8 == null && count == 0) continue;
			else{
				System.out.println("Cell NULL");
				flag = false;
			}

			if(flag == false && count > 4){
				//System.out.println("In a row: " + count);
				count = 0;
			}

			//insert
			if(flag == false && count == 4){
				System.out.println("Insert Mode");
				count = 0;
				int last = sheet.getLastRowNum();

				//Row tempRow = sheet.getRow(i - 1);
				int cellNum = i - 1;
				Row tempRow = sheet.getRow(cellNum);
				if(tempRow == null) continue;
				Cell cell4 = tempRow.getCell(4);
				Cell cell10 = tempRow.getCell(10);
				System.out.println("Cell NUll at: " + cellNum + "Cell Value: " + sheet.getRow(cellNum).getCell(8).getStringCellValue());


				if(cell4 == null || cell4.getCellType() != Cell.CELL_TYPE_STRING || cell10 == null || cell10.getCellType() != Cell.CELL_TYPE_STRING) {
					//System.out.println("Cell NUll at: " + i + "Cell Value: " + sheet.getRow(i).getCell(8).getStringCellValue());

					continue;
				}

				String service = cell4.getStringCellValue();
				String checkPoint = cell10.getStringCellValue();

				//System.out.println("Service: " + service + ", checkpoint: " + checkPoint);

				sheet.shiftRows(i, last, 2);

				Row tmpRow = sheet.createRow(i - 1);
				tmpRow.createCell(0, Cell.CELL_TYPE_STRING).setCellValue("0");
				tmpRow.createCell(4, Cell.CELL_TYPE_STRING).setCellValue(service);
				tmpRow.createCell(7, Cell.CELL_TYPE_STRING).setCellValue("0");
				tmpRow.createCell(8, Cell.CELL_TYPE_STRING).setCellValue("count");
				tmpRow.createCell(9, Cell.CELL_TYPE_STRING).setCellValue("TradeCaptureReport");
				tmpRow.createCell(10, Cell.CELL_TYPE_STRING).setCellValue(checkPoint);
				tmpRow.createCell(11, Cell.CELL_TYPE_STRING).setCellValue("1");


				tmpRow = sheet.createRow(i);
				tmpRow.createCell(0, Cell.CELL_TYPE_STRING).setCellValue("0");
				tmpRow.createCell(4, Cell.CELL_TYPE_STRING).setCellValue(service);
				tmpRow.createCell(7, Cell.CELL_TYPE_STRING).setCellValue("0");
				tmpRow.createCell(8, Cell.CELL_TYPE_STRING).setCellValue("count");
				tmpRow.createCell(9, Cell.CELL_TYPE_STRING).setCellValue("TradeCaptureReportAck");
				tmpRow.createCell(10, Cell.CELL_TYPE_STRING).setCellValue(checkPoint);
				tmpRow.createCell(11, Cell.CELL_TYPE_STRING).setCellValue("0");

				tmpRow = sheet.createRow(i + 1);
				tmpRow.createCell(0, Cell.CELL_TYPE_STRING).setCellValue("0");
				tmpRow.createCell(4, Cell.CELL_TYPE_STRING).setCellValue(service);
				tmpRow.createCell(7, Cell.CELL_TYPE_STRING).setCellValue("0");
				tmpRow.createCell(8, Cell.CELL_TYPE_STRING).setCellValue("countApp");
				tmpRow.createCell(9, Cell.CELL_TYPE_STRING).setCellValue("");
				tmpRow.createCell(10, Cell.CELL_TYPE_STRING).setCellValue(checkPoint);
				tmpRow.createCell(11, Cell.CELL_TYPE_STRING).setCellValue("1");

				//System.out.println("In a row: " + count);
				count = 0;
				i += 2;
				continue;
			}

			if(cell8.getStringCellValue().contains("count") && row.getCell(4).getStringCellValue().contains("apa")) {
				flag = true;
				//System.out.println(cell8.getStringCellValue());
				count++;
				continue;
			} else {
				//System.out.println("False at: " + i);
				flag = false;
				continue;
			}
		}
	}

	public void correctHeigthWidth() throws IOException{
		System.out.println("correctHeigthWidth");
		log.write("\ncorrectHeigthWidth Func\n");
		HSSFSheet sheet = doc.getSheetAt(0);

		int lastRow = sheet.getLastRowNum();
		for(int i = 0; i < lastRow; i++){ sheet.autoSizeColumn(i); }
	}

	public void fixSaveMessagesPossition() throws IOException{
		System.out.println("saveMessagesPossitions");
		log.write("\nsaveMessagesPossitions Func\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();
		for(int i = 0; i < lastRow; i++){
			Row row = sheet.getRow(i);
			if(row == null) continue;

			Cell cell8 = row.getCell(8);
			if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING) continue;
			String cell8Content = cell8.getStringCellValue();

			if(cell8Content.matches("SetPrefix")){
				//System.out.println("Found SetPrefix");
				Cell nextCell = sheet.getRow(i + 1).getCell(8);
				if(nextCell == null || nextCell.getCellType() != Cell.CELL_TYPE_STRING) continue;
				if(!nextCell.getStringCellValue().contains("Save")) continue;


				int lastCell = row.getLastCellNum();
				for(int j = 14; j < lastCell; j++){
					Cell tempCell = row.getCell(j);
					if(tempCell == null || tempCell.getCellType() != Cell.CELL_TYPE_STRING) continue;

					row.getCell(17).setCellValue(tempCell.getStringCellValue());
					if(i != 16) row.getCell(16).setCellType(Cell.CELL_TYPE_BLANK);
				}
			}

			if(cell8Content.matches("SaveExistMessage")){
				int lastCell = row.getLastCellNum();
				for(int j = 14; j < lastCell; j++){
					Cell tempCell = row.getCell(j);
					if(tempCell == null || tempCell.getCellType() != Cell.CELL_TYPE_STRING) continue;

					row.getCell(16).setCellValue(tempCell.getStringCellValue());
					if(j != 16) row.removeCell(row.getCell(j));
				}
			}
		}
	}

	public void fixUnusedVariables() throws IOException{
			System.out.println("fixUnusedVariables");
			log.write("\ffixUnusedVariables Func\n");

			HashMap<String, Integer> instruments = new HashMap<String, Integer>();
			TreeMap<String, Integer> noParties = new TreeMap<String, Integer>();
			TreeMap<String, TreeSet<Integer>> noSides = new TreeMap<String, TreeSet<Integer>>();
			//LinkedHashMap<String, Integer> repGroups = new LinkedHashMap<String, Integer>();
			TreeMap<String, Integer> repGroups = new TreeMap<String, Integer>();


			TreeSet<Integer> set = new TreeSet<Integer>(instruments.values());
			int begin = 0, end = 0;

			HSSFSheet sheet = doc.getSheetAt(0);
			int lastRow = sheet.getPhysicalNumberOfRows();

			for(int rowNum = 0; rowNum < lastRow; rowNum++){
	//			System.out.println("Row Num: " + rowNum);
				Row row = sheet.getRow(rowNum);
				if(row == null) continue;

				Cell cell8 = row.getCell(8);
				if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING) continue;
				String cell8Content = cell8.getStringCellValue();

				if(cell8Content.matches("test case start")){ begin = rowNum; continue; }

				if(!cell8Content.matches("test case end")) continue;

				end = rowNum;

				int start = 0;

				for(int i = begin; i < end; i++){
					row = sheet.getRow(i);
					if(row == null) continue;

					Cell cell9 = row.getCell(9);
					if(cell9 == null || cell9.getCellType() != Cell.CELL_TYPE_STRING) continue;
					String cell9Content = cell9.getStringCellValue();

					if(cell9Content.matches("TradeCaptureReport")){ start = i; break; }
					if(cell9Content.matches("Instrument")){ instruments.put('[' + row.getCell(6).getStringCellValue() + ']', i); continue; }
					if(cell9Content.matches("TradeCaptureReport_NoSides_NoPartyIDs")){ noParties.put(row.getCell(6).getStringCellValue(), i); continue; }
					if(cell9Content.matches("TradeCaptureReport_NoSides")){
						String cell6Contant = row.getCell(6).getStringCellValue();
						if(noSides.containsKey(cell6Contant)){
							noSides.get(cell6Contant).add(i);
						} else { noSides.put(cell6Contant, new TreeSet<Integer>()); noSides.get(cell6Contant).add(i); }

						continue; }
					if(cell9Content.matches("TradeCaptureReport_NoTrdRegPublications")){ repGroups.put(row.getCell(6).getStringCellValue(), i); continue; }
				}
				System.out.println(repGroups);

				//Check for instrument
				for(int i = start; i < end; i++){
					row = sheet.getRow(i);
					if(row == null) continue;

					cell8 = row.getCell(8);
					if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING) continue;

					if(cell8.getStringCellValue().matches("send")){
						Cell cell17 = row.getCell(17);
						String cell17Content = cell17.getStringCellValue();

						Iterator<Map.Entry<String, Integer>> iter = instruments.entrySet().iterator();
						while(iter.hasNext()){
							Map.Entry<String, Integer> entry = iter.next();
							if(cell17Content.contains(entry.getKey())){ iter.remove(); continue; }
						}
						continue;
					}

					if(cell8.getStringCellValue().matches("receive")){
						Cell cell19 = row.getCell(19);
						if(cell19 == null) continue;

						String cell19Content;

						switch(cell19.getCellType()){
							case(Cell.CELL_TYPE_NUMERIC):
								cell19Content = Double.toString(cell19.getNumericCellValue());
								row.createCell(19, Cell.CELL_TYPE_STRING).setCellValue(cell19Content);
								break;
							case(Cell.CELL_TYPE_STRING):
								cell19Content = cell19.getStringCellValue();
								break;
							default: cell19Content = ""; break;
						}

						//System.out.println("Check for : " + i);

						Iterator<Map.Entry<String, Integer>> iter = instruments.entrySet().iterator();
						while(iter.hasNext()){
							Map.Entry<String, Integer> entry = iter.next();
							if(cell19Content.contains(entry.getKey())){ iter.remove(); continue; }
						}

						//if reject check
						Cell cell16 = row.getCell(16);
						String cell16Content = "";
						switch(cell16.getCellType()){
							case(Cell.CELL_TYPE_NUMERIC): cell16Content = Double.toString(cell16.getNumericCellValue());
								row.createCell(16, Cell.CELL_TYPE_STRING).setCellValue(cell16Content);
								break;
							case(Cell.CELL_TYPE_STRING): cell16Content = cell16.getStringCellValue(); break;
							default: cell16Content = ""; break;
						}

						iter = instruments.entrySet().iterator();
						while(iter.hasNext()){
							Map.Entry<String, Integer> entry = iter.next();
							if(cell16Content.contains(entry.getKey())){ iter.remove(); continue; }
						}
					}
				}
				System.out.println("instruments: " + instruments);

				//Check for NoSides
				TreeMap<String, TreeSet<Integer>> noSidesToDeletion = new TreeMap<String, TreeSet<Integer>>(noSides);
				TreeSet<Integer> noSidesToSave = new TreeSet<Integer>();
				for(int i = start; i < end; i++){
					row = sheet.getRow(i);
					if(row == null) continue;

					cell8 = row.getCell(8);
					if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING) continue;
					if(cell8.getStringCellValue().matches("send")){
						Cell cell46 = row.getCell(46);
						String cell46Content = cell46.getStringCellValue();

						for(String key : noSides.descendingKeySet()){
							if(cell46Content.contains(key)){
								//System.out.println("Remove Key: " + key);
								noSidesToDeletion.remove(key);
								noSidesToSave.addAll(noSides.get(key));
								continue; }
						}
					}

					if(cell8.getStringCellValue().matches("receive")){
						Cell cell52 = row.getCell(52);
						if(cell52 == null || cell52.getCellType() != Cell.CELL_TYPE_STRING) continue;
						String cell52Content = cell52.getStringCellValue();

						for(String key : noSides.descendingKeySet()){
							if(cell52Content.contains(key)){
								//System.out.println("Remove Key: " + key);
								noSidesToDeletion.remove(key);
								noSidesToSave.addAll(noSides.get(key));
								continue; }
						}
					}
				}

				noSides = noSidesToDeletion;

				//Check for repGroup !!! 47 - 48 CELLSSSSSSSSSSSSSSSS
				for(int i = start; i < end; i++){
					row = sheet.getRow(i);
					if(row == null) continue;

					cell8 = row.getCell(8);
					if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING) continue;

					if(cell8.getStringCellValue().matches("send")){
						System.out.println("SEND: " + i);
						for(int j = 47; j < row.getLastCellNum(); j++){
							Row headerRow = sheet.getRow(i - 1);
							if(headerRow == null) continue;

							Cell headerTempCell = headerRow.getCell(j);
							if(headerTempCell == null) continue;

							if(headerTempCell.getStringCellValue().contains("NoTrdRegPublications")){
								//System.out.println("Header Row: " + repGroups);
								Cell tempCell = row.getCell(j);
								if(tempCell == null || tempCell.getCellType() != Cell.CELL_TYPE_STRING) continue;
								String tempCellContent = tempCell.getStringCellValue();

								Iterator<Map.Entry<String, Integer>> iter = repGroups.descendingMap().entrySet().iterator();
								while(iter.hasNext()){
									Map.Entry<String, Integer> entry = iter.next();
									if(tempCellContent.contains(entry.getKey())){ iter.remove(); continue; }
								}
							}
						}
					}

					if(cell8.getStringCellValue().matches("receive")){
						//int cellCount = row.getLastCellNum();
						for(int j = 54; j < 56; j++){
							Cell tempCell = row.getCell(j);
							if(tempCell == null || tempCell.getCellType() != Cell.CELL_TYPE_STRING) continue;
							String tempCellContent = tempCell.getStringCellValue();

							Iterator<Map.Entry<String, Integer>> iter = repGroups.entrySet().iterator();
							while(iter.hasNext()){
								Map.Entry<String, Integer> entry = iter.next();
								if(tempCellContent.contains(entry.getKey())){ System.out.println("Exclude row: " + i); iter.remove(); continue; }
							}
						}
					}
				}

				//NoParty Remove
				for(Integer tempRowNum : noSidesToSave){
					Iterator<Map.Entry<String, Integer>> iter = noParties.entrySet().iterator();
					while(iter.hasNext()){
						Map.Entry<String, Integer> entry = iter.next();
						Row tempRow = sheet.getRow(tempRowNum);
						String str = tempRow.getCell(18).getStringCellValue().substring(1, tempRow.getCell(18).getStringCellValue().length() - 1);

						for(String tmp : str.split(", ")){
							if(tmp.matches(entry.getKey())){ iter.remove(); break; }
						}
					}
				}

				set.addAll(instruments.values());
				set.addAll(noParties.values());
				for(String key : noSides.keySet()) set.addAll(noSides.get(key));
				set.addAll(repGroups.values());

				rowNum = end;
				noSides = new TreeMap<String, TreeSet<Integer>>();
			}

			System.out.println("ROW TO DELETION: " + set);
			Iterator it = set.descendingIterator();
			while(it.hasNext()){
				Integer i = (Integer)it.next();
				sheet.removeRow(sheet.getRow(i));
				sheet.shiftRows(i + 1, sheet.getLastRowNum(), -1);
			}
		}

	public void correctFonts(String fontName, Short fontSize) throws IOException{
		System.out.println("correctFonts");
		log.write("\ncorrectFonts\n");

		Font font = doc.createFont();
		font.setFontName(fontName);
		Integer tmpValue = 20 * fontSize;
		font.setFontHeight(tmpValue.shortValue());

		HSSFSheet sheet = doc.getSheetAt(0);
		int rowCount = sheet.getLastRowNum();
		for(int i = 0; i < rowCount; i++){
			Row row = sheet.getRow(i);
			if(row == null) continue;

			int cellCount = row.getLastCellNum();
			for(int j = 0; j < cellCount; j++){
				Cell cell = row.getCell(j);
				if(cell == null) continue;

				CellStyle style = cell.getCellStyle();
				style.setFont(font);
				cell.setCellStyle(style);
			}
		}
	}

	public void correctMessageDescription() throws IOException{
		System.out.println("correctMessageDescription");
		log.write("\ncorrectMessageDescription\n");

		HSSFSheet sheet = doc.getSheetAt(0);
		int rowCount = sheet.getLastRowNum();
		for(int i = 0; i < rowCount; i++){
			Row row = sheet.getRow(i);
			if(row == null) continue;

			Cell cell6 = row.getCell(6);
			if(cell6 == null || cell6.getCellType() != Cell.CELL_TYPE_STRING) continue;

			String cell6Content = cell6.getStringCellValue();
			if(cell6Content.matches("si_gtp")){ row.getCell(3).setCellValue("Submission immediately GTP");  continue; }
			if(cell6Content.matches("si_dss_c")){ row.getCell(3).setCellValue("Submission immediately DSS Contra"); continue; }
			if(cell6Content.matches("si_dss_e")){ row.getCell(3).setCellValue("Submission immediately DSS Executer");continue;  }
			if(cell6Content.matches("si_dss_a")){ row.getCell(3).setCellValue("Submission immediately DSS Ack"); continue; }
			if(cell6Content.matches("si_c")){ row.getCell(3).setCellValue("Submission immediately FIX Contra"); continue; }
			if(cell6Content.matches("si_e")){ row.getCell(3).setCellValue("Submission immediately FIX Executer"); continue; }
			if(cell6Content.matches("si_a")){ row.getCell(3).setCellValue("Submission immediately FIX Ack"); continue; }
			if(cell6Content.matches("si")){ row.getCell(3).setCellValue("Submission immediately"); continue; }

			if(cell6Content.matches("su_dss_c")){ row.getCell(3).setCellValue("Submission unpublished DSS Contra"); continue; }
			if(cell6Content.matches("su_dss_e")){ row.getCell(3).setCellValue("Submission unpublished DSS Executer"); continue; }
			if(cell6Content.matches("su_dss_a")){ row.getCell(3).setCellValue("Submission unpublished DSS Ack"); continue; }
			if(cell6Content.matches("su_c")){ row.getCell(3).setCellValue("Submission unpublished FIX Contra"); continue; }
			if(cell6Content.matches("su_e")){ row.getCell(3).setCellValue("Submission unpublished FIX Executer"); continue; }
			if(cell6Content.matches("su_a")){ row.getCell(3).setCellValue("Submission unpublished FIX Ack"); continue; }
			if(cell6Content.matches("su")){ row.getCell(3).setCellValue("Submission unpublished"); continue; }

			if(cell6Content.matches("psd_gtp")){ row.getCell(3).setCellValue("Publication system set deferral GTP"); continue; }
			if(cell6Content.matches("psd_dss_c")){ row.getCell(3).setCellValue("Publication system set deferral DSS Contra"); continue; }
			if(cell6Content.matches("psd_dss_e")){ row.getCell(3).setCellValue("Publication system set deferral DSS Executer"); continue; }
			if(cell6Content.matches("psd_c")){ row.getCell(3).setCellValue("Publication system set deferral FIX Contra"); continue; }
			if(cell6Content.matches("psd_e")){ row.getCell(3).setCellValue("Publication system set deferral FIX Executer"); continue; }
			if(cell6Content.matches("ssd_dss_c")){ row.getCell(3).setCellValue("Submission system set deferral DSS Contra"); continue; }
			if(cell6Content.matches("ssd_dss_e")){ row.getCell(3).setCellValue("Submission system set deferral DSS Executer"); continue; }
			if(cell6Content.matches("ssd_dss_a")){ row.getCell(3).setCellValue("Submission system set deferral DSS Ack"); continue; }
			if(cell6Content.matches("ssd_c")){ row.getCell(3).setCellValue("Submission system set deferral FIX Contra"); continue; }
			if(cell6Content.matches("ssd_e")){ row.getCell(3).setCellValue("Submission system set deferral FIX Executer"); continue; }
			if(cell6Content.matches("ssd_a")){ row.getCell(3).setCellValue("Submission system set deferral FIX Ack"); continue; }
			if(cell6Content.matches("ssd")){ row.getCell(3).setCellValue("Submission system set deferral"); continue; }

			if(cell6Content.matches("pmd_gtp")){ row.getCell(3).setCellValue("Publication manual set deferral GTP"); continue; }
			if(cell6Content.matches("pmd_dss_c")){ row.getCell(3).setCellValue("Publication manual set deferral DSS Contra"); continue; }
			if(cell6Content.matches("pmd_dss_e")){ row.getCell(3).setCellValue("Publication manual set deferral DSS Executer"); continue; }
			if(cell6Content.matches("pmd_c")){ row.getCell(3).setCellValue("Publication manual set deferral FIX Contra"); continue; }
			if(cell6Content.matches("pmd_e")){ row.getCell(3).setCellValue("Publication manual set deferral FIX Executer"); continue; }
			if(cell6Content.matches("smd_dss_c")){ row.getCell(3).setCellValue("Submission manual set deferral DSS Contra"); continue; }
			if(cell6Content.matches("smd_dss_e")){ row.getCell(3).setCellValue("Submission manual set deferral DSS Executer"); continue; }
			if(cell6Content.matches("smd_dss_a")){ row.getCell(3).setCellValue("Submission manual set deferral DSS Ack"); continue; }
			if(cell6Content.matches("smd_c")){ row.getCell(3).setCellValue("Submission manual set deferral FIX Contra"); continue; }
			if(cell6Content.matches("smd_e")){ row.getCell(3).setCellValue("Submission manual set deferral FIX Executer"); continue; }
			if(cell6Content.matches("smd_a")){ row.getCell(3).setCellValue("Submission manual set deferral FIX Ack"); continue; }
			if(cell6Content.matches("smd")){ row.getCell(3).setCellValue("Submission manual set deferral"); continue; }

			if(cell6Content.matches("ci_gtp")){ row.getCell(3).setCellValue("Cancellation immediately GTP"); continue; }
			if(cell6Content.matches("ci_dss_c")){ row.getCell(3).setCellValue("Cancellation immediately DSS Contra"); continue; }
			if(cell6Content.matches("ci_dss_e")){ row.getCell(3).setCellValue("Cancellation immediately DSS Executer"); continue; }
			if(cell6Content.matches("ci_dss_a")){ row.getCell(3).setCellValue("Cancellation immediately DSS Ack"); continue; }
			if(cell6Content.matches("ci_c")){ row.getCell(3).setCellValue("Cancellation immediately FIX Contra"); continue; }
			if(cell6Content.matches("ci_e")){ row.getCell(3).setCellValue("Cancellation immediately FIX Executer"); continue; }
			if(cell6Content.matches("ci_a")){ row.getCell(3).setCellValue("Cancellation immediately FIX Ack"); continue; }
			if(cell6Content.matches("ci")){ row.getCell(3).setCellValue("Cancellation immediately"); continue; }

			if(cell6Content.matches("cu_dss_c")){ row.getCell(3).setCellValue("Cancellation unpublished DSS Contra"); continue; }
			if(cell6Content.matches("cu_dss_e")){ row.getCell(3).setCellValue("Cancellation unpublished DSS Executer"); continue; }
			if(cell6Content.matches("cu_dss_a")){ row.getCell(3).setCellValue("Cancellation unpublished DSS Ack"); continue; }
			if(cell6Content.matches("cu_c")){ row.getCell(3).setCellValue("Cancellation unpublished FIX Contra"); continue; }
			if(cell6Content.matches("cu_e")){ row.getCell(3).setCellValue("Cancellation unpublished FIX Executer"); continue; }
			if(cell6Content.matches("cu_a")){ row.getCell(3).setCellValue("Cancellation unpublished FIX Ack"); continue; }
			if(cell6Content.matches("cu")){ row.getCell(3).setCellValue("Cancellation unpublished"); continue; }

			if(cell6Content.matches("csdb_dss_c")){ row.getCell(3).setCellValue("Cancellation system set deferral before DSS Contra"); continue; }
			if(cell6Content.matches("csdb_dss_e")){ row.getCell(3).setCellValue("Cancellation system set deferral before DSS Executer"); continue; }
			if(cell6Content.matches("csdb_dss_a")){ row.getCell(3).setCellValue("Cancellation system set deferral before DSS Ack"); continue; }
			if(cell6Content.matches("csdb_c")){ row.getCell(3).setCellValue("Cancellation system set deferral before FIX Contra"); continue; }
			if(cell6Content.matches("csdb_e")){ row.getCell(3).setCellValue("Cancellation system set deferral before FIX Executer"); continue; }
			if(cell6Content.matches("csdb_a")){ row.getCell(3).setCellValue("Cancellation system set deferral before FIX Ack"); continue; }
			if(cell6Content.matches("csdb")){ row.getCell(3).setCellValue("Cancellation system set deferral before"); continue; }

			if(cell6Content.matches("cmdb_dss_c")){ row.getCell(3).setCellValue("Cancellation manual set deferral before DSS Contra"); continue; }
			if(cell6Content.matches("cmdb_dss_e")){ row.getCell(3).setCellValue("Cancellation manual set deferral before DSS Executer"); continue; }
			if(cell6Content.matches("cmdb_dss_a")){ row.getCell(3).setCellValue("Cancellation manual set deferral before DSS Ack"); continue; }
			if(cell6Content.matches("cmdb_c")){ row.getCell(3).setCellValue("Cancellation manual set deferral before FIX Contra"); continue; }
			if(cell6Content.matches("cmdb_e")){ row.getCell(3).setCellValue("Cancellation manual set deferral before FIX Executer"); continue; }
			if(cell6Content.matches("cmdb_a")){ row.getCell(3).setCellValue("Cancellation manual set deferral before FIX Ack"); continue; }
			if(cell6Content.matches("cmdb")){ row.getCell(3).setCellValue("Cancellation manual set deferral before"); continue; }

			if(cell6Content.matches("csda_gtp")){ row.getCell(3).setCellValue("Cancellation system set deferral after GTP"); continue; }
			if(cell6Content.matches("csda_dss_c")){ row.getCell(3).setCellValue("Cancellation system set deferral after DSS Contra"); continue; }
			if(cell6Content.matches("csda_dss_e")){ row.getCell(3).setCellValue("Cancellation system set deferral after DSS Executer"); continue; }
			if(cell6Content.matches("csda_dss_a")){ row.getCell(3).setCellValue("Cancellation system set deferral after DSS Ack"); continue; }
			if(cell6Content.matches("csda_c")){ row.getCell(3).setCellValue("Cancellation system set deferral after FIX Contra"); continue; }
			if(cell6Content.matches("csda_e")){ row.getCell(3).setCellValue("Cancellation system set deferral after FIX Executer"); continue; }
			if(cell6Content.matches("csda_a")){ row.getCell(3).setCellValue("Cancellation system set deferral after FIX Ack"); continue; }
			if(cell6Content.matches("csda")){ row.getCell(3).setCellValue("Cancellation system set deferral after"); continue; }

			if(cell6Content.matches("cmda_gtp")){ row.getCell(3).setCellValue("Cancellation manual set deferral after GTP"); continue; }
			if(cell6Content.matches("cmda_dss_c")){ row.getCell(3).setCellValue("Cancellation manual set deferral after DSS Contra"); continue; }
			if(cell6Content.matches("cmda_dss_e")){ row.getCell(3).setCellValue("Cancellation manual set deferral after DSS Executer"); continue; }
			if(cell6Content.matches("cmda_dss_a")){ row.getCell(3).setCellValue("Cancellation manual set deferral after DSS Ack"); continue; }
			if(cell6Content.matches("cmda_c")){ row.getCell(3).setCellValue("Cancellation manual set deferral after FIX Contra"); continue; }
			if(cell6Content.matches("cmda_e")){ row.getCell(3).setCellValue("Cancellation manual set deferral after FIX Executer"); continue; }
			if(cell6Content.matches("cmda_a")){ row.getCell(3).setCellValue("Cancellation manual set deferral after FIX Ack"); continue; }
			if(cell6Content.matches("cmda")){ row.getCell(3).setCellValue("Cancellation manual set deferral after"); continue; }

			if(cell6Content.matches("pri_dss_a")){ row.getCell(3).setCellValue("Pre-Release immediately DSS Ack"); continue; }
			if(cell6Content.matches("pri_a")){ row.getCell(3).setCellValue("Pre-Release immediately FIX Ack"); continue; }
			if(cell6Content.matches("pri")){ row.getCell(3).setCellValue("Pre-Release immediately"); continue; }

			if(cell6Content.matches("pru_dss_a")){ row.getCell(3).setCellValue("Pre-Release unpublished DSS Ack"); continue; }
			if(cell6Content.matches("pru_a")){ row.getCell(3).setCellValue("Pre-Release unpublished FIX Ack"); continue; }
			if(cell6Content.matches("pru")){ row.getCell(3).setCellValue("Pre-Release unpublished"); continue; }

			if(cell6Content.matches("prsda_dss_a")){ row.getCell(3).setCellValue("Pre-Release system set deferral after DSS Ack"); continue; }
			if(cell6Content.matches("prsda_a")){ row.getCell(3).setCellValue("Pre-Release system set deferral after FIX Ack"); continue; }
			if(cell6Content.matches("prsda")){ row.getCell(3).setCellValue("Pre-Release system set deferral after"); continue; }

			if(cell6Content.matches("prmda_dss_a")){ row.getCell(3).setCellValue("Pre-Release manual set deferral after DSS Ack"); continue; }
			if(cell6Content.matches("prmda_a")){ row.getCell(3).setCellValue("Pre-Release manual set deferral after FIX Ack"); continue; }
			if(cell6Content.matches("prmda")){ row.getCell(3).setCellValue("Pre-Release manual set deferral after"); continue; }

			if(cell6Content.matches("prsd_gtp")){ row.getCell(3).setCellValue("Pre-Release system set deferral before GTP"); continue; }
			if(cell6Content.matches("prsd_dss_c")){ row.getCell(3).setCellValue("Pre-Release system set deferral before DSS Contra"); continue; }
			if(cell6Content.matches("prsd_dss_e")){ row.getCell(3).setCellValue("Pre-Release system set deferral before DSS Executer"); continue; }
			if(cell6Content.matches("prsd_dss_a")){ row.getCell(3).setCellValue("Pre-Release system set deferral before DSS Ack"); continue; }
			if(cell6Content.matches("prsd_c")){ row.getCell(3).setCellValue("Pre-Release system set deferral before FIX Contra"); continue; }
			if(cell6Content.matches("prsd_e")){ row.getCell(3).setCellValue("Pre-Release system set deferral before FIX Executer"); continue; }
			if(cell6Content.matches("prsd_a")){ row.getCell(3).setCellValue("Pre-Release system set deferral before FIX Ack"); continue; }
			if(cell6Content.matches("prsd")){ row.getCell(3).setCellValue("Pre-Release system set deferral before"); continue; }

			if(cell6Content.matches("prmd_gtp")){ row.getCell(3).setCellValue("Pre-Release manual set deferral before GTP"); continue; }
			if(cell6Content.matches("prmd_dss_c")){ row.getCell(3).setCellValue("Pre-Release manual set deferral before DSS Contra"); continue; }
			if(cell6Content.matches("prmd_dss_e")){ row.getCell(3).setCellValue("Pre-Release manual set deferral before DSS Executer"); continue; }
			if(cell6Content.matches("prmd_dss_a")){ row.getCell(3).setCellValue("Pre-Release manual set deferral before DSS Ack"); continue; }
			if(cell6Content.matches("prmd_c")){ row.getCell(3).setCellValue("Pre-Release manual set deferral before FIX Contra"); continue; }
			if(cell6Content.matches("prmd_e")){ row.getCell(3).setCellValue("Pre-Release manual set deferral before FIX Executer"); continue; }
			if(cell6Content.matches("prmd_a")){ row.getCell(3).setCellValue("Pre-Release manual set deferral before FIX Ack"); continue; }
			if(cell6Content.matches("prmd")){ row.getCell(3).setCellValue("Pre-Release manual set deferral before"); continue; }

			if(cell6Content.matches("ai_gtp")){ row.getCell(3).setCellValue("Amendment immediately GTP"); continue; }
			if(cell6Content.matches("ai_dss_c")){ row.getCell(3).setCellValue("Amendment immediately DSS Contra"); continue; }
			if(cell6Content.matches("ai_dss_e")){ row.getCell(3).setCellValue("Amendment immediately DSS Executer"); continue; }
			if(cell6Content.matches("ai_dss_a")){ row.getCell(3).setCellValue("Amendment immediately DSS Ack"); continue; }
			if(cell6Content.matches("ai_c")){ row.getCell(3).setCellValue("Amendment immediately FIX Contra"); continue; }
			if(cell6Content.matches("ai_e")){ row.getCell(3).setCellValue("Amendment immediately FIX Executer"); continue; }
			if(cell6Content.matches("ai_a")){ row.getCell(3).setCellValue("Amendment immediately FIX Ack"); continue; }
			if(cell6Content.matches("ai")){ row.getCell(3).setCellValue("Amendment immediately"); continue; }

			if(cell6Content.matches("au_dss_c")){ row.getCell(3).setCellValue("Amendment unpublished DSS Contra"); continue; }
			if(cell6Content.matches("au_dss_e")){ row.getCell(3).setCellValue("Amendment unpublished DSS Executer"); continue; }
			if(cell6Content.matches("au_dss_a")){ row.getCell(3).setCellValue("Amendment unpublished DSS Ack"); continue; }
			if(cell6Content.matches("au_c")){ row.getCell(3).setCellValue("Amendment unpublished FIX Contra"); continue; }
			if(cell6Content.matches("au_e")){ row.getCell(3).setCellValue("Amendment unpublished FIX Executer"); continue; }
			if(cell6Content.matches("au_a")){ row.getCell(3).setCellValue("Amendment unpublished FIX Ack"); continue; }
			if(cell6Content.matches("au")){ row.getCell(3).setCellValue("Amendment unpublished"); continue; }

			if(cell6Content.matches("asdb_dss_c")){ row.getCell(3).setCellValue("Amendment system set deferral before DSS Contra"); continue; }
			if(cell6Content.matches("asdb_dss_e")){ row.getCell(3).setCellValue("Amendment system set deferral before DSS Executer"); continue; }
			if(cell6Content.matches("asdb_dss_a")){ row.getCell(3).setCellValue("Amendment system set deferral before DSS Ack"); continue; }
			if(cell6Content.matches("asdb_c")){ row.getCell(3).setCellValue("Amendment system set deferral before FIX Contra"); continue; }
			if(cell6Content.matches("asdb_e")){ row.getCell(3).setCellValue("Amendment system set deferral before FIX Executer"); continue; }
			if(cell6Content.matches("asdb_a")){ row.getCell(3).setCellValue("Amendment system set deferral before FIX Ack"); continue; }
			if(cell6Content.matches("asdb")){ row.getCell(3).setCellValue("Amendment system set deferral before"); continue; }

			if(cell6Content.matches("amdb_dss_c")){ row.getCell(3).setCellValue("Amendment manual set deferral before DSS Contra"); continue; }
			if(cell6Content.matches("amdb_dss_e")){ row.getCell(3).setCellValue("Amendment manual set deferral before DSS Executer"); continue; }
			if(cell6Content.matches("amdb_dss_a")){ row.getCell(3).setCellValue("Amendment manual set deferral before DSS Ack"); continue; }
			if(cell6Content.matches("amdb_c")){ row.getCell(3).setCellValue("Amendment manual set deferral before FIX Contra"); continue; }
			if(cell6Content.matches("amdb_e")){ row.getCell(3).setCellValue("Amendment manual set deferral before FIX Executer"); continue; }
			if(cell6Content.matches("amdb_a")){ row.getCell(3).setCellValue("Amendment manual set deferral before FIX Ack"); continue; }
			if(cell6Content.matches("amdb")){ row.getCell(3).setCellValue("Amendment manual set deferral before"); continue; }

			if(cell6Content.matches("asda_gtp")){ row.getCell(3).setCellValue("Amendment system set deferral after GTP"); continue; }
			if(cell6Content.matches("asda_dss_c")){ row.getCell(3).setCellValue("Amendment system set deferral after DSS Contra"); continue; }
			if(cell6Content.matches("asda_dss_e")){ row.getCell(3).setCellValue("Amendment system set deferral after DSS Executer"); continue; }
			if(cell6Content.matches("asda_dss_a")){ row.getCell(3).setCellValue("Amendment system set deferral after DSS Ack"); continue; }
			if(cell6Content.matches("asda_c")){ row.getCell(3).setCellValue("Amendment system set deferral after FIX Contra"); continue; }
			if(cell6Content.matches("asda_e")){ row.getCell(3).setCellValue("Amendment system set deferral after FIX Executer"); continue; }
			if(cell6Content.matches("asda_a")){ row.getCell(3).setCellValue("Amendment system set deferral after FIX Ack"); continue; }
			if(cell6Content.matches("asda")){ row.getCell(3).setCellValue("Amendment system set deferral after"); continue; }

			if(cell6Content.matches("amda_gtp")){ row.getCell(3).setCellValue("Amendment manual set deferral after GTP"); continue; }
			if(cell6Content.matches("amda_dss_c")){ row.getCell(3).setCellValue("Amendment manual set deferral after DSS Contra"); continue; }
			if(cell6Content.matches("amda_dss_e")){ row.getCell(3).setCellValue("Amendment manual set deferral after DSS Executer"); continue; }
			if(cell6Content.matches("amda_dss_a")){ row.getCell(3).setCellValue("Amendment manual set deferral after DSS Ack"); continue; }
			if(cell6Content.matches("amda_c")){ row.getCell(3).setCellValue("Amendment manual set deferral after FIX Contra"); continue; }
			if(cell6Content.matches("amda_e")){ row.getCell(3).setCellValue("Amendment manual set deferral after FIX Executer"); continue; }
			if(cell6Content.matches("amda_a")){ row.getCell(3).setCellValue("Amendment manual set deferral after FIX Ack"); continue; }
			if(cell6Content.matches("amda")){ row.getCell(3).setCellValue("Amendment manual set deferral after"); continue; }
		}
	}

	public void newClearFunc() throws IOException{
		System.out.println("newClearFunc");
		log.write("\nnewClearFunc\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		int rowCount = sheet.getLastRowNum();
		System.out.println(sheet.getLastRowNum());
		System.out.println(sheet.getPhysicalNumberOfRows());

		sheet.shiftRows(0, rowCount, rowCount);
		for(int i = 0; i < rowCount; i++){
			Row row = sheet.getRow(i);
			//if(row == null) continue;
			if(row == null) { sheet.createRow(i); }
			sheet.removeRow(sheet.getRow(i));
		}

		sheet.shiftRows(rowCount, sheet.getLastRowNum(), -rowCount);
	}

	public void fixDiff2() throws IOException{ //new
		//Depends on Dashes
		System.out.println("fixDiff");
		log.write("\nfixDiff\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		boolean flag = false;
		String tempString = "";

		String[] toSave = { "si_e", "psd_e", "pmd_e", "ai_e", "prsd_e" };
		HashSet<String> toSaveSet = new HashSet<String>(Arrays.asList(toSave));

		for(int i = 0; i < sheet.getLastRowNum(); i++){
			Row row = sheet.getRow(i);
			if(row == null) continue;

			Cell cell8 = row.getCell(8);
			if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING) continue;
			String cell8Content = cell8.getStringCellValue();


			if(cell8Content.matches("receive") && toSaveSet.contains(row.getCell(6).getStringCellValue())){
				flag = true; tempString = row.getCell(6).getStringCellValue(); continue;
			}

			if(flag && cell8Content.matches("CheckMessage")){
				System.out.println(tempString);
				row.createCell(1, Cell.CELL_TYPE_BLANK);
				row.createCell(14, Cell.CELL_TYPE_STRING).setCellValue("(#{diffDateTime(${" + tempString + ".header.SendingTime}, ${" + tempString +
						".RptTime}, \"s\")} >= 0) && (#{diffDateTime(${" + tempString + ".header.SendingTime}, ${" + tempString +
						".RptTime}, \"s\")} < 1)");
				tempString = "";
				flag = false;
				continue;
			}
		}
	}

	public void fixDiff14032018() throws IOException{
		System.out.println("fixDiff");
		log.write("\nfixDiff\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		int last = sheet.getPhysicalNumberOfRows();
		for(int rowNum = 0; rowNum < last; rowNum++){
			Row row = sheet.getRow(rowNum);
			if(row == null) continue;

			Cell cell8 = row.getCell(8);
			if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING) continue;

			String cell8Content = cell8.getStringCellValue();
			if(cell8Content.contains("CheckMessage")){
				Cell cell1 = row.getCell(1);
				if(cell1 == null) row.createCell(1, Cell.CELL_TYPE_BLANK);
				else cell1.setCellValue("");
				log.write(Calendar.getInstance().getTime() + "Remove 'n' from row:\t" + rowNum + "\tcell: 1\n");

				Cell cell14 = row.getCell(14);
				if(cell14 == null || cell14.getCellType() != Cell.CELL_TYPE_STRING) continue;

				String cell14Content = cell14.getStringCellValue();
				if(cell14Content.contains("{diff(")){
					cell14.setCellValue(cell14Content.replace("{diff(", "{diffDateTime("));
					log.write(Calendar.getInstance().getTime() + "Correct 'Diff Func' in row:\t" + rowNum + "\tcell: 14\n");
				}
			}
		}
	}

	public void usersRempacement(TreeMap<String, String> m) throws IOException{
		System.out.println("usersRempacement");
		HSSFSheet sheet = doc.getSheetAt(0);
		int i = 2;
		while(true){
			Row row = sheet.getRow(i);
			if(row == null){ i++; continue; }

			Cell cell8 = row.getCell(8);
			if(cell8 == null){ i++; continue; }
			if(cell8.getStringCellValue().matches("Global Block end")) break;

			Cell cell6 = row.getCell(6);
			if(cell6 == null){ i++; continue; }

			String cell6Content = cell6.getStringCellValue();
			switch(cell6Content){
				case("Executer"): row.getCell(13).setCellValue(m.get("Executer")); i++; break;
				case("ExecuterFIX"): row.getCell(13).setCellValue(m.get("ExecuterFIX")); i++; break;
				case("ConterParty"): row.getCell(13).setCellValue(m.get("ConterParty")); i++; break;
				case("ContraFIX"): row.getCell(13).setCellValue(m.get("ContraFIX")); i++; break;
				case("ExecuterLEI"): row.getCell(13).setCellValue(m.get("ExecuterLEI")); i++; break;
				case("ConterPartyLEI"): row.getCell(13).setCellValue(m.get("ConterPartyLEI")); i++; break;
				default: i++;
			}
		}
	}

	public void insertIntoGlobalBlock(String desc, String ref, String type, String value){
		System.out.println("insertIntoGlobalBlock");
		HSSFSheet sheet = doc.getSheetAt(0);
		int last = sheet.getPhysicalNumberOfRows();
		sheet.shiftRows(2, last, 1);
		Row row = sheet.createRow(2);
		row.createCell(3, Cell.CELL_TYPE_STRING).setCellValue(desc);
		row.createCell(6, Cell.CELL_TYPE_STRING).setCellValue(desc);
		row.createCell(8, Cell.CELL_TYPE_STRING).setCellValue("SetStatic");
		row.createCell(12, Cell.CELL_TYPE_STRING).setCellValue(type);
		row.createCell(13, Cell.CELL_TYPE_STRING).setCellValue(value);
	}

	public void replaceReferenceNames(){
		//Submission immideately
		LinkedHashMap<String, String> si = new LinkedHashMap<String, String>();
		si.put("tcr1_dsg_tr", "si_gtp");
		si.put("tcr1_dss_ex2", "si_dss_c");
		si.put("tcr1_dss_ex", "si_dss_e");
		si.put("tcr1_dss_ack", "si_dss_a");
		si.put("tcr1Ac", "si_c");
		si.put("tcr1Ae", "si_e");
		si.put("tcr1A_ack", "si_a");
		si.put("tcr1A", "si");
		System.out.println(si.keySet());

		//Submission unpublished
		LinkedHashMap<String, String> su = new LinkedHashMap<String, String>();
		su.put("tcr1_dss_ex2", "su_dss_c");
		su.put("tcr1_dss_ex", "su_dss_e");
		su.put("tcr1_dss_ack", "su_dss_a");
		su.put("tcr1Ac", "su_c");
		su.put("tcr1Ae", "su_e");
		su.put("tcr1A_ack", "su_a");
		su.put("tcr1A", "su");


		//Submission/publication system deferred
		LinkedHashMap<String, String> ssd = new LinkedHashMap<String, String>();
		ssd.put("tcr2_dsg_tr", "psd_gtp");
		ssd.put("tcr2_dss_ex2", "psd_dss_c");
		ssd.put("tcr2_dss_ex", "psd_dss_e");
		ssd.put("tcr2Ac", "psd_c");
		ssd.put("tcr2Ae", "psd_e");
		ssd.put("tcr1_dss_ex2", "ssd_dss_c");
		ssd.put("tcr1_dss_ex", "ssd_dss_e");
		ssd.put("tcr1_dss_ack", "ssd_dss_a");
		ssd.put("tcr1Ac", "ssd_c");
		ssd.put("tcr1Ae", "ssd_e");
		ssd.put("tcr1A_ack", "ssd_a");
		ssd.put("tcr1A", "ssd");

		//Submission/publication manual deferred
		LinkedHashMap<String, String> smd = new LinkedHashMap<String, String>();
		smd.put("tcr2_dsg_tr", "pmd_gtp");
		smd.put("tcr2_dss_ex2", "pmd_dss_c");
		smd.put("tcr2_dss_ex", "pmd_dss_e");
		smd.put("tcr2Ac", "pmd_c");
		smd.put("tcr2Ae", "pmd_e");
		smd.put("tcr1_dss_ex2", "smd_dss_c");
		smd.put("tcr1_dss_ex", "smd_dss_e");
		smd.put("tcr1_dss_ack", "smd_dss_a");
		smd.put("tcr1Ac", "smd_c");
		smd.put("tcr1Ae", "smd_e");
		smd.put("tcr1A_ack", "smd_a");
		smd.put("tcr1A", "smd");

		//Cancellation immideately
		LinkedHashMap<String, String> ci = new LinkedHashMap<String, String>();
		ci.put("tcr2_dsg_tr", "ci_gtp");
		ci.put("tcr2_dss_ex2", "ci_dss_c");
		ci.put("tcr2_dss_ex", "ci_dss_e");
		ci.put("tcr2_dss_ack", "ci_dss_a");
		ci.put("tcr2Ac", "ci_c");
		ci.put("tcr2Ae", "ci_e");
		ci.put("tcr2A_ack", "ci_a");
		ci.put("tcr2A", "ci");
		ci.put("tcr1_dsg_tr", "si_gtp");
		ci.put("tcr1_dss_ex2", "si_dss_c");
		ci.put("tcr1_dss_ex", "si_dss_e");
		ci.put("tcr1_dss_ack", "si_dss_a");
		ci.put("tcr1Ac", "si_c");
		ci.put("tcr1Ae", "si_e");
		ci.put("tcr1A_ack", "si_a");
		ci.put("tcr1A", "si");

		//Cancellation unpublished
		LinkedHashMap<String, String> cu = new LinkedHashMap<String, String>();
		cu.put("tcr2_dss_ex2", "cu_dss_c");
		cu.put("tcr2_dss_ex", "cu_dss_e");
		cu.put("tcr2_dss_ack", "cu_dss_a");
		cu.put("tcr2Ac", "cu_c");
		cu.put("tcr2Ae", "cu_e");
		cu.put("tcr2A_ack", "cu_a");
		cu.put("tcr2A", "cu");
		cu.put("tcr1_dss_ex2", "su_dss_c");
		cu.put("tcr1_dss_ex", "su_dss_e");
		cu.put("tcr1_dss_ack", "su_dss_a");
		cu.put("tcr1Ac", "su_c");
		cu.put("tcr1Ae", "su_e");
		cu.put("tcr1A_ack", "su_a");
		cu.put("tcr1A", "su");

		//Cancellation system deferred before publication
		LinkedHashMap<String, String> csdb = new LinkedHashMap<String, String>();
		csdb.put("tcr2_dss_ex2", "csdb_dss_c");
		csdb.put("tcr2_dss_ex", "csdb_dss_e");
		csdb.put("tcr2_dss_ack", "csdb_dss_a");
		csdb.put("tcr2Ac", "csdb_c");
		csdb.put("tcr2Ae", "csdb_e");
		csdb.put("tcr2A_ack", "csdb_a");
		csdb.put("tcr2A", "csdb");
		csdb.put("tcr1_dss_ex2", "ssd_dss_c");
		csdb.put("tcr1_dss_ex", "ssd_dss_e");
		csdb.put("tcr1_dss_ack", "ssd_dss_a");
		csdb.put("tcr1Ac", "ssd_c");
		csdb.put("tcr1Ae", "ssd_e");
		csdb.put("tcr1A_ack", "ssd_a");
		csdb.put("tcr1A", "ssd");

		//Cancellation manual deferred before publication
		LinkedHashMap<String, String> cmdb = new LinkedHashMap<String, String>();
		cmdb.put("tcr2_dss_ex2", "cmdb_dss_c");
		cmdb.put("tcr2_dss_ex", "cmdb_dss_e");
		cmdb.put("tcr2_dss_ack", "cmdb_dss_a");
		cmdb.put("tcr2Ac", "cmdb_c");
		cmdb.put("tcr2Ae", "cmdb_e");
		cmdb.put("tcr2A_ack", "cmdb_a");
		cmdb.put("tcr2A", "cmdb");
		cmdb.put("tcr1_dss_ex2", "smd_dss_c");
		cmdb.put("tcr1_dss_ex", "smd_dss_e");
		cmdb.put("tcr1_dss_ack", "smd_dss_a");
		cmdb.put("tcr1Ac", "smd_c");
		cmdb.put("tcr1Ae", "smd_e");
		cmdb.put("tcr1A_ack", "smd_a");
		cmdb.put("tcr1A", "smd");

		//Cancellation system deferred after publication
		LinkedHashMap<String, String> csda = new LinkedHashMap<String, String>();
		csda.put("tcr3_dsg_tr", "csda_gtp");
		csda.put("tcr3_dss_ex2", "csda_dss_c");
		csda.put("tcr3_dss_ex", "csda_dss_e");
		csda.put("tcr3_dss_ack", "csda_dss_a");
		csda.put("tcr3Ac", "csda_c");
		csda.put("tcr3Ae", "csda_e");
		csda.put("tcr3A_ack", "csda_a");
		csda.put("tcr3A", "csda");
		csda.put("tcr2_dsg_tr", "psd_gtp");
		csda.put("tcr2_dss_ex2", "psd_dss_c");
		csda.put("tcr2_dss_ex", "psd_dss_e");
		csda.put("tcr2Ac", "psd_c");
		csda.put("tcr2Ae", "psd_e");
		csda.put("tcr1_dss_ex2", "ssd_dss_c");
		csda.put("tcr1_dss_ex", "ssd_dss_e");
		csda.put("tcr1_dss_ack", "ssd_dss_a");
		csda.put("tcr1Ac", "ssd_c");
		csda.put("tcr1Ae", "ssd_e");
		csda.put("tcr1A_ack", "ssd_a");
		csda.put("tcr1A", "ssd");

		//Cancellation manual deferred after publication
		LinkedHashMap<String, String> cmda = new LinkedHashMap<String, String>();
		cmda.put("tcr3_dsg_tr", "cmda_gtp");
		cmda.put("tcr3_dss_ex2", "cmda_dss_c");
		cmda.put("tcr3_dss_ex", "cmda_dss_e");
		cmda.put("tcr3_dss_ack", "cmda_dss_a");
		cmda.put("tcr3Ac", "cmda_c");
		cmda.put("tcr3Ae", "cmda_e");
		cmda.put("tcr3A_ack", "cmda_a");
		cmda.put("tcr3A", "cmda");
		cmda.put("tcr2_dsg_tr", "pmd_gtp");
		cmda.put("tcr2_dss_ex2", "pmd_dss_c");
		cmda.put("tcr2_dss_ex", "pmd_dss_e");
		cmda.put("tcr2Ac", "pmd_c");
		cmda.put("tcr2Ae", "pmd_e");
		cmda.put("tcr1_dss_ex2", "smd_dss_c");
		cmda.put("tcr1_dss_ex", "smd_dss_e");
		cmda.put("tcr1_dss_ack", "smd_dss_a");
		cmda.put("tcr1Ac", "smd_c");
		cmda.put("tcr1Ae", "smd_e");
		cmda.put("tcr1A_ack", "smd_a");
		cmda.put("tcr1A", "smd");

		//Pre-Release immideately
		LinkedHashMap<String, String> pri = new LinkedHashMap<String, String>();
		pri.put("tcr2_dss_ack", "pri_dss_a");
		pri.put("tcr2A_ack", "pri_a");
		pri.put("tcr2A", "pri");
		pri.put("tcr1_dsg_tr", "si_gtp");
		pri.put("tcr1_dss_ex2", "si_dss_c");
		pri.put("tcr1_dss_ex", "si_dss_e");
		pri.put("tcr1_dss_ack", "si_dss_a");
		pri.put("tcr1Ac", "si_c");
		pri.put("tcr1Ae", "si_e");
		pri.put("tcr1A_ack", "si_a");
		pri.put("tcr1A", "si");
		System.out.println(pri.keySet());

		//Submission unpublished
		LinkedHashMap<String, String> pru = new LinkedHashMap<String, String>();
		pru.put("tcr2_dss_ack", "pru_dss_a");
		pru.put("tcr2A_ack", "pru_a");
		pru.put("tcr2A", "pru");
		pru.put("tcr1_dss_ex2", "su_dss_c");
		pru.put("tcr1_dss_ex", "su_dss_e");
		pru.put("tcr1_dss_ack", "su_dss_a");
		pru.put("tcr1Ac", "su_c");
		pru.put("tcr1Ae", "su_e");
		pru.put("tcr1A_ack", "su_a");
		pru.put("tcr1A", "su");


		//Pre-Release system deferred after publication
		LinkedHashMap<String, String> prsda = new LinkedHashMap<String, String>();
		prsda.put("tcr3_dss_ack", "prsda_dss_a");
		prsda.put("tcr3A_ack", "prsda_a");
		prsda.put("tcr3A", "prsda");
		prsda.put("tcr2_dsg_tr", "psd_gtp");
		prsda.put("tcr2_dss_ex2", "psd_dss_c");
		prsda.put("tcr2_dss_ex", "psd_dss_e");
		prsda.put("tcr2Ac", "psd_c");
		prsda.put("tcr2Ae", "psd_e");
		prsda.put("tcr1_dss_ex2", "ssd_dss_c");
		prsda.put("tcr1_dss_ex", "ssd_dss_e");
		prsda.put("tcr1_dss_ack", "ssd_dss_a");
		prsda.put("tcr1Ac", "ssd_c");
		prsda.put("tcr1Ae", "ssd_e");
		prsda.put("tcr1A_ack", "ssd_a");
		prsda.put("tcr1A", "ssd");

		//Pre-Release manual deferred after publication
		LinkedHashMap<String, String> prmda = new LinkedHashMap<String, String>();
		prmda.put("tcr3_dss_ack", "prmda_dss_a");
		prmda.put("tcr3A_ack", "prmda_a");
		prmda.put("tcr3A", "prmda");
		prmda.put("tcr2_dsg_tr", "pmd_gtp");
		prmda.put("tcr2_dss_ex2", "pmd_dss_c");
		prmda.put("tcr2_dss_ex", "pmd_dss_e");
		prmda.put("tcr2Ac", "pmd_c");
		prmda.put("tcr2Ae", "pmd_e");
		prmda.put("tcr1_dss_ex2", "smd_dss_c");
		prmda.put("tcr1_dss_ex", "smd_dss_e");
		prmda.put("tcr1_dss_ack", "smd_dss_a");
		prmda.put("tcr1Ac", "smd_c");
		prmda.put("tcr1Ae", "smd_e");
		prmda.put("tcr1A_ack", "smd_a");
		prmda.put("tcr1A", "smd");

		//Pre-Release system deferral
		LinkedHashMap<String, String> prsd = new LinkedHashMap<String, String>();
		prsd.put("tcr2_dsg_tr", "prsd_gtp");
		prsd.put("tcr2_dss_ex2", "prsd_dss_c");
		prsd.put("tcr2_dss_ex", "prsd_dss_e");
		prsd.put("tcr2_dss_ack", "prsd_dss_a");
		prsd.put("tcr2Ac", "prsd_c");
		prsd.put("tcr2Ae", "prsd_e");
		prsd.put("tcr2A_ack", "prsd_a");
		prsd.put("tcr2A", "prsd");
		prsd.put("tcr1_dss_ex2", "ssd_dss_c");
		prsd.put("tcr1_dss_ex", "ssd_dss_e");
		prsd.put("tcr1_dss_ack", "ssd_dss_a");
		prsd.put("tcr1Ac", "ssd_c");
		prsd.put("tcr1Ae", "ssd_e");
		prsd.put("tcr1A_ack", "ssd_a");
		prsd.put("tcr1A", "ssd");

		//Pre-Release manual deferral
		LinkedHashMap<String, String> prmd = new LinkedHashMap<String, String>();
		prmd.put("tcr2_dsg_tr", "prmd_gtp");
		prmd.put("tcr2_dss_ex2", "prmd_dss_c");
		prmd.put("tcr2_dss_ex", "prmd_dss_e");
		prmd.put("tcr2_dss_ack", "prmd_dss_a");
		prmd.put("tcr2Ac", "prmd_c");
		prmd.put("tcr2Ae", "prmd_e");
		prmd.put("tcr2A_ack", "prmd_a");
		prmd.put("tcr2A", "prmd");
		prmd.put("tcr1_dss_ex2", "smd_dss_c");
		prmd.put("tcr1_dss_ex", "smd_dss_e");
		prmd.put("tcr1_dss_ack", "smd_dss_a");
		prmd.put("tcr1Ac", "smd_c");
		prmd.put("tcr1Ae", "smd_e");
		prmd.put("tcr1A_ack", "smd_a");
		prmd.put("tcr1A", "smd");

		//Amend immideately


		//Amend unpublished
		//Amend system deferral before publication
		//Amend manual deferral before publication
		//Amend system deferral after publication
		//Amend manual deferral after publication

		//Amend immideately
		LinkedHashMap<String, String> ai = new LinkedHashMap<String, String>();
		ai.put("tcr3_dsg_tr", "ai_gtp");
		ai.put("tcr3_dss_ex2", "ai_dss_c");
		ai.put("tcr3_dss_ex", "ai_dss_e");
		ai.put("tcr3_dss_ack", "ai_dss_a");
		ai.put("tcr3Ac", "ai_c");
		ai.put("tcr3Ae", "ai_e");
		ai.put("tcr3A_ack", "ai_a");
		ai.put("tcr3A", "ai");
		ai.put("tcr2_dsg_tr", "ci_gtp");
		ai.put("tcr2_dss_ex2", "ci_dss_c");
		ai.put("tcr2_dss_ex", "ci_dss_e");
		ai.put("tcr2_dss_ack", "ci_dss_a");
		ai.put("tcr2Ac", "ci_c");
		ai.put("tcr2Ae", "ci_e");
		ai.put("tcr2A_ack", "ci_a");
		ai.put("tcr2A", "ci");
		ai.put("tcr1_dsg_tr", "si_gtp");
		ai.put("tcr1_dss_ex2", "si_dss_c");
		ai.put("tcr1_dss_ex", "si_dss_e");
		ai.put("tcr1_dss_ack", "si_dss_a");
		ai.put("tcr1Ac", "si_c");
		ai.put("tcr1Ae", "si_e");
		ai.put("tcr1A_ack", "si_a");
		ai.put("tcr1A", "si");

		//Amend unpublished
		LinkedHashMap<String, String> au = new LinkedHashMap<String, String>();
		au.put("tcr3_dss_ex2", "au_dss_c");
		au.put("tcr3_dss_ex", "au_dss_e");
		au.put("tcr3_dss_ack", "au_dss_a");
		au.put("tcr3Ac", "au_c");
		au.put("tcr3Ae", "au_e");
		au.put("tcr3A_ack", "au_a");
		au.put("tcr3A", "au");
		au.put("tcr2_dss_ex2", "cu_dss_c");
		au.put("tcr2_dss_ex", "cu_dss_e");
		au.put("tcr2_dss_ack", "cu_dss_a");
		au.put("tcr2Ac", "cu_c");
		au.put("tcr2Ae", "cu_e");
		au.put("tcr2A_ack", "cu_a");
		au.put("tcr2A", "cu");
		au.put("tcr1_dss_ex2", "su_dss_c");
		au.put("tcr1_dss_ex", "su_dss_e");
		au.put("tcr1_dss_ack", "su_dss_a");
		au.put("tcr1Ac", "su_c");
		au.put("tcr1Ae", "su_e");
		au.put("tcr1A_ack", "su_a");
		au.put("tcr1A", "su");

		//Amend system deferred before publication
		LinkedHashMap<String, String> asdb = new LinkedHashMap<String, String>();
		asdb.put("tcr3_dss_ex2", "asdb_dss_c");
		asdb.put("tcr3_dss_ex", "asdb_dss_e");
		asdb.put("tcr3_dss_ack", "asdb_dss_a");
		asdb.put("tcr3Ac", "asdb_c");
		asdb.put("tcr3Ae", "asdb_e");
		asdb.put("tcr3A_ack", "asdb_a");
		asdb.put("tcr3A", "asdb");
		asdb.put("tcr2_dss_ex2", "csdb_dss_c");
		asdb.put("tcr2_dss_ex", "csdb_dss_e");
		asdb.put("tcr2_dss_ack", "csdb_dss_a");
		asdb.put("tcr2Ac", "csdb_c");
		asdb.put("tcr2Ae", "csdb_e");
		asdb.put("tcr2A_ack", "csdb_a");
		asdb.put("tcr2A", "csdb");
		asdb.put("tcr1_dss_ex2", "ssd_dss_c");
		asdb.put("tcr1_dss_ex", "ssd_dss_e");
		asdb.put("tcr1_dss_ack", "ssd_dss_a");
		asdb.put("tcr1Ac", "ssd_c");
		asdb.put("tcr1Ae", "ssd_e");
		asdb.put("tcr1A_ack", "ssd_a");
		asdb.put("tcr1A", "ssd");

		//Amend manual deferred before publication
		LinkedHashMap<String, String> amdb = new LinkedHashMap<String, String>();
		amdb.put("tcr3_dss_ex2", "amdb_dss_c");
		amdb.put("tcr3_dss_ex", "amdb_dss_e");
		amdb.put("tcr3_dss_ack", "amdb_dss_a");
		amdb.put("tcr3Ac", "amdb_c");
		amdb.put("tcr3Ae", "amdb_e");
		amdb.put("tcr3A_ack", "amdb_a");
		amdb.put("tcr3A", "amdb");
		amdb.put("tcr2_dss_ex2", "cmdb_dss_c");
		amdb.put("tcr2_dss_ex", "cmdb_dss_e");
		amdb.put("tcr2_dss_ack", "cmdb_dss_a");
		amdb.put("tcr2Ac", "cmdb_c");
		amdb.put("tcr2Ae", "cmdb_e");
		amdb.put("tcr2A_ack", "cmdb_a");
		amdb.put("tcr2A", "cmdb");
		amdb.put("tcr1_dss_ex2", "smd_dss_c");
		amdb.put("tcr1_dss_ex", "smd_dss_e");
		amdb.put("tcr1_dss_ack", "smd_dss_a");
		amdb.put("tcr1Ac", "smd_c");
		amdb.put("tcr1Ae", "smd_e");
		amdb.put("tcr1A_ack", "smd_a");
		amdb.put("tcr1A", "smd");

		//Amend system deferred after publication
		LinkedHashMap<String, String> asda = new LinkedHashMap<String, String>();
		asda.put("tcr4_dsg_tr", "asda_gtp");
		asda.put("tcr4_dss_ex2", "asda_dss_c");
		asda.put("tcr4_dss_ex", "asda_dss_e");
		asda.put("tcr4_dss_ack", "asda_dss_a");
		asda.put("tcr4Ac", "asda_c");
		asda.put("tcr4Ae", "asda_e");
		asda.put("tcr4A_ack", "asda_a");
		asda.put("tcr4A", "asda");
		asda.put("tcr3_dsg_tr", "csda_gtp");
		asda.put("tcr3_dss_ex2", "csda_dss_c");
		asda.put("tcr3_dss_ex", "csda_dss_e");
		asda.put("tcr3_dss_ack", "csda_dss_a");
		asda.put("tcr3Ac", "csda_c");
		asda.put("tcr3Ae", "csda_e");
		asda.put("tcr3A_ack", "csda_a");
		asda.put("tcr3A", "csda");
		asda.put("tcr2_dsg_tr", "psd_gtp");
		asda.put("tcr2_dss_ex2", "psd_dss_c");
		asda.put("tcr2_dss_ex", "psd_dss_e");
		asda.put("tcr2Ac", "psd_c");
		asda.put("tcr2Ae", "psd_e");
		asda.put("tcr1_dss_ex2", "ssd_dss_c");
		asda.put("tcr1_dss_ex", "ssd_dss_e");
		asda.put("tcr1_dss_ack", "ssd_dss_a");
		asda.put("tcr1Ac", "ssd_c");
		asda.put("tcr1Ae", "ssd_e");
		asda.put("tcr1A_ack", "ssd_a");
		asda.put("tcr1A", "ssd");

		//Cancellation manual deferred after publication
		LinkedHashMap<String, String> amda = new LinkedHashMap<String, String>();
		amda.put("tcr4_dsg_tr", "amda_gtp");
		amda.put("tcr4_dss_ex2", "amda_dss_c");
		amda.put("tcr4_dss_ex", "amda_dss_e");
		amda.put("tcr4_dss_ack", "amda_dss_a");
		amda.put("tcr4Ac", "amda_c");
		amda.put("tcr4Ae", "amda_e");
		amda.put("tcr4A_ack", "amda_a");
		amda.put("tcr4A", "amda");
		amda.put("tcr3_dsg_tr", "cmda_gtp");
		amda.put("tcr3_dss_ex2", "cmda_dss_c");
		amda.put("tcr3_dss_ex", "cmda_dss_e");
		amda.put("tcr3_dss_ack", "cmda_dss_a");
		amda.put("tcr3Ac", "cmda_c");
		amda.put("tcr3Ae", "cmda_e");
		amda.put("tcr3A_ack", "cmda_a");
		amda.put("tcr3A", "cmda");
		amda.put("tcr2_dsg_tr", "pmd_gtp");
		amda.put("tcr2_dss_ex2", "pmd_dss_c");
		amda.put("tcr2_dss_ex", "pmd_dss_e");
		amda.put("tcr2Ac", "pmd_c");
		amda.put("tcr2Ae", "pmd_e");
		amda.put("tcr1_dss_ex2", "smd_dss_c");
		amda.put("tcr1_dss_ex", "smd_dss_e");
		amda.put("tcr1_dss_ack", "smd_dss_a");
		amda.put("tcr1Ac", "smd_c");
		amda.put("tcr1Ae", "smd_e");
		amda.put("tcr1A_ack", "smd_a");
		amda.put("tcr1A", "smd");

		LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();

		HSSFSheet sheet = doc.getSheetAt(0);
		int lastRow = sheet.getPhysicalNumberOfRows();
		boolean flag = false;
		for(int i = 0; i < lastRow; i++){
			Cell cell = sheet.getRow(i).getCell(8);
			if(cell == null || cell.getCellType() != Cell.CELL_TYPE_STRING) continue;

			if(cell.getStringCellValue().matches("test case start")){
				System.out.println("Start");

				Cell cell3 = sheet.getRow(i).getCell(3);
				String cell3Content = cell3.getStringCellValue();

				if(cell3Content.contains("type=si")){ map = si; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=su")){ map = su; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=ssd")){ map = ssd; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=smd")){ map = smd; flag = true; System.out.println(map.toString()); }

				if(cell3Content.contains("type=ci")){ map = ci; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=cu")){ map = cu; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=csdb")){ map = csdb; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=cmdb")){ map = cmdb; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=csda")){ map = csda; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=cmda")){ map = cmda; flag = true; System.out.println(map.toString()); }

				if(cell3Content.contains("type=prsd")){ map = prsd; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=prmd")){ map = prmd; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=pri")){ map = pri; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=pru")){ map = pru; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=prsda")){ map = prsda; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=prmda")){ map = prmda; flag = true; System.out.println(map.toString()); }

				if(cell3Content.contains("type=ai")){ map = ai; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=au")){ map = au; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=asdb")){ map = asdb; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=amdb")){ map = amdb; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=asda")){ map = asda; flag = true; System.out.println(map.toString()); }
				if(cell3Content.contains("type=amda")){ map = amda; flag = true; System.out.println(map.toString()); }

				//String cell2 = sheet.getRow(i).getCell(3).getStringCellValue();
			}

			if(cell.getStringCellValue().matches("test case end")) {
				flag = false;
				map = null;
				System.out.println("End");
				continue;
			}

			if(flag){
				int lastCell = sheet.getRow(i).getLastCellNum();
				for(int j = 6; j < lastCell; j++){
					Cell tempCell = sheet.getRow(i).getCell(j);
					if(tempCell == null || tempCell.getCellType() != Cell.CELL_TYPE_STRING) continue;

					if(j == 6){
						for(String key : map.keySet()){
							String cellValue = tempCell.getStringCellValue();

							if((cellValue + ".").matches(key)){
								System.out.println("Found TCR");
								tempCell.setCellValue(map.get(key).substring(0, map.get(key).length() - 1));
							}
						}
						//continue;
					}


					for(String key : map.keySet()){
						String cellValue = tempCell.getStringCellValue();
						//System.out.println("Row: " + cellValue + ", key: " + key + ", " + cellValue.contains(key));

						if(cellValue.contains(key)){
							//System.out.println("Found, Row: " + i + ", cell: " + j);
							String strToReplace = cellValue.replace(key, map.get(key));
							tempCell.setCellValue(strToReplace);
						}
					}
				}
			}
		}
	}

	public void fixFIXcounts(){
		System.out.println("fixFIXcounts Func ");
		HSSFSheet sheet = doc.getSheetAt(0);
		int last = sheet.getPhysicalNumberOfRows();

		for(int i = 0; i < last; i++){

			Row row = sheet.getRow(i);
			Cell cell = row.getCell(8);
			if(cell == null){
				row.createCell(8, Cell.CELL_TYPE_BLANK);
				continue;
			}

			if(cell.getStringCellValue().matches("countApp")){
				Cell tempCell = row.getCell(i + 2);
				//System.out.println("Next cell: " + tempCell);
				//if(tempCell.getStringCellValue().matches(""))
				if(tempCell == null){
					System.out.println("NEED ADD ROWS " + i);
					//System.out.println("!!! " + sheet.getRow(i + 1).getCell(4));
					if(sheet.getRow(i + 1).getCell(4) == null) {
						sheet.getRow(i + 1).createCell(4, Cell.CELL_TYPE_STRING).setCellValue("");
					}

					String service = sheet.getRow(i + 1).getCell(4).getStringCellValue();

					if(sheet.getRow(i + 1).getCell(10) == null) {
						sheet.getRow(i + 1).createCell(10, Cell.CELL_TYPE_STRING).setCellValue("");
					}
					String checkPoint = sheet.getRow(i + 1).getCell(10).getStringCellValue();
					//System.out.println(service + ". " + checkPoint);

					sheet.shiftRows(i + 2, last, 2);

					Row tmpRow = sheet.createRow(i + 2);
					tmpRow.createCell(0, Cell.CELL_TYPE_STRING).setCellValue("0");
					tmpRow.createCell(4, Cell.CELL_TYPE_STRING).setCellValue(service);
					tmpRow.createCell(7, Cell.CELL_TYPE_STRING).setCellValue("0");
					tmpRow.createCell(8, Cell.CELL_TYPE_STRING).setCellValue("count");
					tmpRow.createCell(9, Cell.CELL_TYPE_STRING).setCellValue("TradeCaptureReportAck");
					tmpRow.createCell(10, Cell.CELL_TYPE_STRING).setCellValue(checkPoint);
					tmpRow.createCell(11, Cell.CELL_TYPE_STRING).setCellValue("0");

					tmpRow = sheet.createRow(i + 3);
					tmpRow.createCell(0, Cell.CELL_TYPE_STRING).setCellValue("0");
					tmpRow.createCell(4, Cell.CELL_TYPE_STRING).setCellValue(service);
					tmpRow.createCell(7, Cell.CELL_TYPE_STRING).setCellValue("0");
					tmpRow.createCell(8, Cell.CELL_TYPE_STRING).setCellValue("countApp");
					tmpRow.createCell(9, Cell.CELL_TYPE_STRING).setCellValue("");
					tmpRow.createCell(10, Cell.CELL_TYPE_STRING).setCellValue(checkPoint);
					tmpRow.createCell(11, Cell.CELL_TYPE_STRING).setCellValue("1");

					i += 3;
				}
			}
			last = sheet.getPhysicalNumberOfRows();
		}
	}

	public void correctHeaders() throws IOException{
		System.out.println("correctHeaders");
		log.write("\ncorrectHeaders\n");

		HSSFSheet sheet = doc.getSheetAt(0);
		Iterator<Row> rowIter = sheet.iterator();
		Set<Integer> sHeaders = new TreeSet<Integer>();
		ArrayList<Integer> lst = new ArrayList<Integer>();

		while(rowIter.hasNext()){
			Cell cell = rowIter.next().getCell(8);
			try{
			if(cell.getCellType() == Cell.CELL_TYPE_STRING && cell.getStringCellValue().matches("DefineHeader")){
				int i = cell.getRowIndex() + 1;
				Row tmpRow = sheet.getRow(i);
				if(tmpRow.getCell(8).getStringCellValue().matches("count")){
					sHeaders.add(i - 1);
					lst.add(i - 1);
				}
			}
			}catch(Exception e){ log.write("correctHeaders " + rowIter + "\n"); }
		}

		//System.out.println(sHeaders);
		//System.out.println(lst);

		//replace headers
		CharSequence charSeq1 = "dss", charSeq2 = "rtf";
		ListIterator<Integer> iter = lst.listIterator(lst.size());
		while(iter.hasPrevious()){
			int i = iter.previous();
			//System.out.println(i);
			Row row = sheet.createRow(i);
			row.createCell(8).setCellValue("DefineHeader");
			row.createCell(14).setCellValue("TradeID");

			Cell cell = sheet.getRow(i + 1).getCell(4);
			if(cell.getStringCellValue().contains(charSeq1) || cell.getStringCellValue().contains(charSeq2)){
				row.createCell(14).setCellValue("TradeMatchID");
				row.createCell(15).setCellValue("InstrumentID");
				row.createCell(16).setCellValue("message");
				row.createCell(17).setCellValue("prefix");
			}
		}
	}

	//Add count filters 08032018
	public void addCountFilters08032018() throws Exception{
		System.out.println("addCountFilters08032018");
		log.write("\naddCountFilters08032018\n");

		HSSFSheet sheet = doc.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();
		//String FirmTradeID = "";
		String TradeID = "";
		boolean flag = false;
		String instrumentFilter = "%{Instrument}";

		for(int i = 0; i < lastRow; i++){
			Row row = sheet.getRow(i);
			if(row == null) continue;

			Cell cell9 = row.getCell(9);
			if(cell9 != null && cell9.getCellType() == Cell.CELL_TYPE_STRING){
				String cell9Content = cell9.getStringCellValue();
				if(cell9Content.matches("Instrument")){
					instrumentFilter = row.getCell(14).getStringCellValue();
					if(instrumentFilter.contains("ISIN")) instrumentFilter = instrumentFilter.replaceAll("ISIN", "");
				}
			}



			Cell cell8 = row.getCell(8);
			if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING) continue;
			String cell8Content = cell8.getStringCellValue();

			if(!flag && cell8Content.matches("LoadMessage") && row.getCell(9).getStringCellValue().matches("TradeCaptureReport")){
				System.out.println("FOUND LOAD MESSAGE");
				//FirmTradeID = "${" + row.getCell(6).getStringCellValue() + ".FirmTradeID}";
				TradeID = "${" + row.getCell(6).getStringCellValue() + ".TradeID}";
				flag = true;
				continue;
			}

			//System.out.println("FIRMTRADEID!!!: " + TradeID);

			//recieve message
			if(cell8Content.matches("receive")){
				//System.out.println("FOUND receive");
				Cell cell6 = row.getCell(6);
				if(cell6 == null || cell6.getCellType() != Cell.CELL_TYPE_STRING) continue;
				String cell6Content = cell6.getStringCellValue();
				System.out.println("cell6Content: " + cell6Content);


				if(flag){
					switch(cell6Content){
					case("ai_a"): TradeID = "${ai_a.TradeID}"; break;
					case("au_a"): TradeID = "${au_a.TradeID}"; break;
					case("asdb_a"): TradeID = "${asdb_a.TradeID}"; break;
					case("amdb_a"): TradeID = "${amdb_a.TradeID}"; break;
					case("asda_a"): TradeID = "${asda_a.TradeID}"; break;
					case("amda_a"): TradeID = "${amda_a.TradeID}"; break;
					default: break;
					}
				} else{
					switch(cell6Content){
						case("si_a"): TradeID = "${si_a.TradeID}"; break;
						case("su_a"): TradeID = "${su_a.TradeID}"; break;
						case("ssd_a"): TradeID = "${ssd_a.TradeID}"; break;
						case("smd_a"): TradeID = "${smd_a.TradeID}"; break;

						case("ci_a"): TradeID = "${si_a.TradeID}"; break;
						case("cu_a"): TradeID = "${su_a.TradeID}"; break;
						case("csdb_a"): TradeID = "${ssd_a.TradeID}"; break;
						case("cmdb_a"): TradeID = "${smd_a.TradeID}"; break;
						case("csda_a"): TradeID = "${ssd_a.TradeID}"; break;
						case("cmda_a"): TradeID = "${smd_a.TradeID}"; break;

						case("pri_a"): TradeID = "${si_a.TradeID}"; break;
						case("pru_a"): TradeID = "${su_a.TradeID}"; break;
						case("prsda_a"): TradeID = "${ssd_a.TradeID}"; break;
						case("prmda_a"): TradeID = "${smd_a.TradeID}"; break;
						case("prsd_a"): TradeID = "${ssd_a.TradeID}"; break;
						case("prmd_a"): TradeID = "${smd_a.TradeID}"; break;

						case("ai_a"): TradeID = "${ai_a.TradeID}"; break;
						case("au_a"): TradeID = "${au_a.TradeID}"; break;
						case("asdb_a"): TradeID = "${asdb_a.TradeID}"; break;
						case("amdb_a"): TradeID = "${amdb_a.TradeID}"; break;
						case("asda_a"): TradeID = "${asda_a.TradeID}"; break;
						case("amda_a"): TradeID = "${amda_a.TradeID}"; break;
					}
				}



				//if(flag) continue;

				/*switch(cell6Content){
					case("si_a"): TradeID = "${si_a.TradeID}"; break;
					case("su_a"): TradeID = "${su_a.TradeID}"; break;
					case("ssd_a"): TradeID = "${ssd_a.TradeID}"; break;
					case("smd_a"): TradeID = "${smd_a.TradeID}"; break;

					case("ci_a"): TradeID = "${si_a.TradeID}"; break;
					case("cu_a"): TradeID = "${su_a.TradeID}"; break;
					case("csdb_a"): TradeID = "${ssd_a.TradeID}"; break;
					case("cmdb_a"): TradeID = "${smd_a.TradeID}"; break;
					case("csda_a"): TradeID = "${ssd_a.TradeID}"; break;
					case("cmda_a"): TradeID = "${smd_a.TradeID}"; break;

					case("pri_a"): TradeID = "${si_a.TradeID}"; break;
					case("pru_a"): TradeID = "${su_a.TradeID}"; break;
					case("prsda_a"): TradeID = "${ssd_a.TradeID}"; break;
					case("prmda_a"): TradeID = "${smd_a.TradeID}"; break;
					case("prsd_a"): TradeID = "${ssd_a.TradeID}"; break;
					case("prmd_a"): TradeID = "${smd_a.TradeID}"; break;

					case("ai_a"): TradeID = "${ai_a.TradeID}"; break;
					case("au_a"): TradeID = "${au_a.TradeID}"; break;
					case("asdb_a"): TradeID = "${asdb_a.TradeID}"; break;
					case("amdb_a"): TradeID = "${amdb_a.TradeID}"; break;
					case("asda_a"): TradeID = "${asda_a.TradeID}"; break;
					case("amda_a"): TradeID = "${amda_a.TradeID}"; break;
				}*/

				//System.out.println("TradeID: " + TradeID);


				//if(cell6Content.contains("ci_e")) { TradeID = "${si_a.TradeID}"; continue; }
				//if(cell6Content.contains("cu_e")) { TradeID = "${su_a.TradeID}"; continue; }
				//if(cell6Content.contains("csd_e")) { TradeID = "${ssd_a.TradeID}"; continue; }
				//if(cell6Content.contains("cmd_e")) { TradeID = "${smd_a.TradeID}"; continue; }

				//if(cell6Content.matches("si_e")) { TradeID = "${si_a.TradeID}"; continue; }
				//if(cell6Content.matches("ci_e")) { TradeID = "${si_a.TradeID}"; continue; }
			}


			//counts
			if(cell8Content.matches("count")){
				System.out.println("FOUND count");

				String tmpStr = row.getCell(4).getStringCellValue();
				if(tmpStr.contains("apa")){
					row.createCell(14).setCellValue(TradeID);
					continue;
				}

				if(tmpStr.contains("dss") || tmpStr.contains("gtp")){
					//remove old filters //'BZ'
					int lastCell = row.getLastCellNum();
					for(int j = 14; j < lastCell; j++){
						Cell tempCell = row.getCell(j);
						if(tempCell == null) continue;

						tempCell.setCellType(Cell.CELL_TYPE_BLANK);
					}

					row.createCell(14).setCellValue(TradeID);
					row.createCell(15).setCellValue(instrumentFilter);
					continue;
				}
			}

			if(cell8Content.matches("test case start")){
				flag = false;
				TradeID = "";
			}
		}
	}

	public void removeFixHeaders() throws IOException{
		System.out.println("removeFixHeaders");
		log.write("\nremoveFixHeaders\n");

		HSSFSheet sheet = doc.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();

		ArrayList<Integer> s = new ArrayList<Integer>();
		CharSequence charSeq = "";

		for(int i = 0; i < lastRow; i++){
			Row row = sheet.getRow(i);
			Cell cell8 = row.getCell(8);
			if(cell8 == null){ /*System.out.println("Row: " + i);*/ continue; }

			if(cell8.getStringCellValue().matches("send")){
				charSeq = row.getCell(4).getStringCellValue();
				continue;
			}

			if((row.getCell(8).getStringCellValue().matches("count") || row.getCell(8).getStringCellValue().matches("countApp"))
					&& (row.getCell(4).getStringCellValue().contains(charSeq) || row.getCell(4).getStringCellValue().contains("apa"))){
				s.add(row.getRowNum());
			}
		}

		//System.out.println("COunts: " + s);

		ListIterator<Integer> siter = s.listIterator(s.size());
		while(siter.hasPrevious()){
			int i = (int)siter.previous() + 1;
			if(sheet.getRow(i) == null){ sheet.createRow(i).createCell(8, Cell.CELL_TYPE_STRING).setCellValue(""); }

			int last = sheet.getLastRowNum();

			Cell cell = sheet.getRow(i).getCell(8);
			if(cell == null) continue;
			if(cell.getCellType() == Cell.CELL_TYPE_STRING){
				try{
					if(cell.getStringCellValue().matches("DefineHeader")){ sheet.shiftRows(i + 1, last, -1); }
				}catch(Exception e){ log.write("Row: " + i + "\n"); e.printStackTrace(); }
			}
		}
	}

	public void fillEmptyCells() throws IOException{
		System.out.println("fillEmptyCells");
		HSSFSheet sheet = doc.getSheetAt(0);

		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++){
			try{
				Row row = sheet.getRow(i);
				if(row == null){
					log.write(Calendar.getInstance().getTime() + "\tCreate Row:\t" + i + "\n");
					sheet.createRow(i);
					for(int j = 0; j < 256; j++){
						log.write(Calendar.getInstance().getTime() + "\tCreate Cell at:\tRow:\t" + i + "\tCell:\t" + j + "\n");
						sheet.getRow(i).createCell(j, Cell.CELL_TYPE_BLANK);
					}
					continue;
				}

				for(int j = 0; j < 121; j++){
				try{
					Cell cell = sheet.getRow(i).getCell(j);
					if(cell == null){
						log.write(Calendar.getInstance().getTime() + "\tCreate Cell at:\tRow:\t" + i + "\tCell:\t" + j + "\n");
						sheet.getRow(i).createCell(j, Cell.CELL_TYPE_BLANK);
					}
				}catch(Exception e){ e.printStackTrace(); }
			}
			}catch(Exception e){ e.printStackTrace(); }
		}
	}

	public void lineNumbers() throws Exception{
		System.out.println("lineNumbers");
		HSSFSheet sheet = doc.getSheetAt(0);

		//find start & end of test case
		caseStartEnd = new TreeSet<Integer>();
		Iterator<Row> rowIter = doc.getSheetAt(0).iterator();

		while(rowIter.hasNext()){
			Row row = rowIter.next();
			Iterator<Cell> cellIter = row.iterator();
			while(cellIter.hasNext()){
				Cell cell = cellIter.next();

				if(cell.getCellType() == Cell.CELL_TYPE_STRING){
					String tempStr = cell.getStringCellValue();
					if(tempStr.matches("test case start") || tempStr.matches("test case end")){
						caseStartEnd.add(cell.getRowIndex());
					}
				}
			}
		}

		//set numbers
		System.out.println("Test case bounds: " + caseStartEnd);

		Iterator<Integer> iter = caseStartEnd.iterator();
		while(iter.hasNext()){
			int counter = 0;
			int beg = iter.next() + 1;
			int end = iter.next() - 1;
			while(beg != end){
				try{
					beg++;
					Row row = sheet.getRow(beg);
					Cell eightCell = row.getCell(8);
					//System.out.println("Row: " + row.getRowNum() + ", " + zeroCell + ", " + eightCell);
					if(eightCell == null){ row.createCell(8, Cell.CELL_TYPE_BLANK); }

					if(row.getCell(8).getStringCellValue().matches("DefineHeader")){
						continue;
					} else {
						Cell tempCell = row.getCell(0);
						if(tempCell == null){
							//System.out.println(tempCell);
							tempCell = row.createCell(0, Cell.CELL_TYPE_STRING);
						}
						CellStyle style = tempCell.getCellStyle();
						style.setAlignment(CellStyle.ALIGN_LEFT);
						tempCell.setCellStyle(style);
						tempCell.setCellValue(counter++);
					}
				}catch(Exception e){ e.printStackTrace(); }
			}
		}
	}

	public void removeDashes() throws Exception{
		HSSFSheet sheet = doc.getSheetAt(0);
		CharSequence dash1 = "_1", dash2 = "_2", dash3 = "_3", dash4 = "_4";
		Iterator<Row> rowIter = sheet.iterator();
		while(rowIter.hasNext()){
			Row tempRow = rowIter.next();
			Iterator<Cell> cellIter = tempRow.iterator();
			while(cellIter.hasNext()){
				Cell tempCell = cellIter.next();
				if(tempCell.getCellType() == Cell.CELL_TYPE_STRING && (tempCell.getStringCellValue().contains(dash1)
						|| tempCell.getStringCellValue().contains(dash2)
						|| tempCell.getStringCellValue().contains(dash3)
						|| tempCell.getStringCellValue().contains(dash4)))
				{
					String tempStr = tempCell.getStringCellValue();
					tempStr = tempStr.replaceAll("_1", "").replaceAll("_2", "").replaceAll("_3", "").replaceAll("_4", "");
					//System.out.println(tempStr);
					tempCell.setCellValue(tempStr);
					log.write(tempCell.getRowIndex() + ", " + tempCell.getColumnIndex());
					//System.out.println(tempCell.getRowIndex() + ", " + tempCell.getColumnIndex());
				}
			}
		}
	}

	public void removeEmptyRows() throws Exception{
		HSSFSheet sheet = doc.getSheetAt(0);
		Iterator<Row> rowIter = sheet.iterator();
		ArrayList<Integer> lst = new ArrayList<Integer>();

		while(rowIter.hasNext()){
			boolean flag = true, firstOne = true;
			Row row = rowIter.next();
			Iterator<Cell> cellIter = row.iterator();

			//CellIterator cellIter = rowIter.next().cellIterator();
			while(cellIter.hasNext()){
				Cell cell = cellIter.next();
				if(firstOne){
					firstOne = false;
					continue;
				}

				switch(cell.getCellType()){
				case (Cell.CELL_TYPE_BLANK): continue;
				case(Cell.CELL_TYPE_STRING): {
					if(cell.getStringCellValue().length() == 0) continue;
					else flag = false;
					 break;
				}
				case (Cell.CELL_TYPE_BOOLEAN): case (Cell.CELL_TYPE_ERROR):
					case(Cell.CELL_TYPE_FORMULA): case (Cell.CELL_TYPE_NUMERIC) : flag = false;
				}
				//if()
			}
			if(flag) lst.add(row.getRowNum());
			//System.out.println("Row: " + row.getRowNum() + ", " + flag);

		}
		System.out.println(lst);

		for(int i = lst.size() - 1; i > 0; i--){
			int beg = lst.get(i - 1);
			int next = lst.get(i);
//			int last = sheet.getLastRowNum();

			int altLast = sheet.getPhysicalNumberOfRows();
			//System.out.println(beg + ", " + next + ", " + last + ", " + altLast + ": " + (beg == next - 1));

			if(beg == next - 1){
				//System.out.println("IN");
				try{
					sheet.removeRow(sheet.getRow(next));
					sheet.shiftRows(next, altLast, -1);
				}catch(Exception e){ e.printStackTrace(); }
			}
		}
	}

	public void fixBuySellId2() throws IOException{
		System.out.println("fixBuySellId2");
		log.write("\nfixBuySellId2\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		int lastRow = sheet.getLastRowNum();

		rowsToDelete = new TreeSet<Integer>();
		CharSequence charSeq1 = "BuyTrdId", charSeq2 = "SellTrdId";
		LinkedHashMap<String, String> map = new LinkedHashMap<String, String>();
		boolean flag = false;

		for(int i = 0; i < lastRow; i++){
			Row row =sheet.getRow(i);
			if(row == null) { continue; }

			Cell cell8 = row.getCell(8);
			if(!flag && (cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING)){ continue; }

			if(flag && (cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING)){
				rowsToDelete.add(i);
				continue;
			}

			String cell8Content = cell8.getStringCellValue();

			//Replace
			if(cell8Content.matches("receive")){
				int lastCell = row.getLastCellNum();
				for(int j = 0; j < lastCell; j++){
					Cell tempCell = row.getCell(j);
					if(tempCell == null || tempCell.getCellType() != Cell.CELL_TYPE_STRING) continue;

					String tempCellContent = tempCell.getStringCellValue();
					for(String str : map.keySet()){
						if(tempCellContent.contains(str)){
							String tmp = map.get(str);
							tmp = "\\" + tmp.replace("\"", "\'");
							tempCell.setCellValue(tempCell.getStringCellValue().replaceAll("\\%\\{" + str + "\\}", tmp));
						}
					}
				}
			}


			if(!flag && cell8Content.matches("SetStatic") && (row.getCell(6).getStringCellValue().contains(charSeq1) ||
					row.getCell(6).getStringCellValue().contains(charSeq2))){
				map.put(row.getCell(6).getStringCellValue(), row.getCell(13).getStringCellValue());
				rowsToDelete.add(i - 1);
				rowsToDelete.add(i);
				flag = true;
				continue;
			}

			if(flag && cell8Content.matches("SetStatic") && (row.getCell(6).getStringCellValue().contains(charSeq1) ||
					row.getCell(6).getStringCellValue().contains(charSeq2))){
				map.put(row.getCell(6).getStringCellValue(), row.getCell(13).getStringCellValue());
				rowsToDelete.add(i - 1);
				rowsToDelete.add(i);
				continue;
			}

			if(flag && cell8Content.equals("")){
				System.out.println("Empty Cell: " + i);
				rowsToDelete.add(i);
				continue;
			}

			if(flag && cell8Content.matches("DefineHeader")){
				flag = false;
				continue;
			}
		}

		System.out.println(map);
		System.out.println(rowsToDelete);

		//Remove
		ArrayList<Integer> lst = new ArrayList<Integer>();
		lst.addAll(rowsToDelete);
		ListIterator<Integer> it = lst.listIterator(lst.size());
		while(it.hasPrevious()){
			int beg = it.previous();
			sheet.removeRow(sheet.getRow(beg));
			sheet.shiftRows(beg + 1, sheet.getLastRowNum(), -1);
			log.write("Remove row: " + beg + "\n");
		}
	}

	public void fixBuySellIdOld() throws IOException{ //Old
		HSSFSheet sheet = doc.getSheetAt(0);
		rowsToDelete = new TreeSet<Integer>();

		//found
		Iterator<Row> rowIter = sheet.iterator();
		CharSequence charSeq1 = "BuyTrdId", charSeq2 = "SellTrdId";

		while(rowIter.hasNext()){
			Row row = rowIter.next();
			Iterator<Cell> cellIter = row.iterator();
			while(cellIter.hasNext()){
				Cell cell = cellIter.next();

				if(cell.getCellType() == Cell.CELL_TYPE_STRING){
					String tempStr = cell.getStringCellValue();
					if(tempStr.contains(charSeq1) || tempStr.contains(charSeq2)){
						if(cell.getColumnIndex() == 6){
							int i = cell.getRowIndex();
							mVariableNames.put(tempStr, "temp");
							rowsToDelete.add(i - 1);
							rowsToDelete.add(i);
						}
					}
				}
			}
		}

		//replace
		charSeq1 = "{BuyTrdId";
		charSeq2 = "{SellTrdId";
		rowIter = sheet.iterator();

		while(rowIter.hasNext()){
			Row row = rowIter.next();
			Iterator<Cell> cellIter = row.iterator();
			while(cellIter.hasNext()){
				Cell cell = cellIter.next();

				if(cell.getCellType() == Cell.CELL_TYPE_STRING){
					String tempString = cell.getStringCellValue();

					//setValues
					if(mVariableNames.containsKey(tempString)){
						mVariableNames.put(tempString, row.getCell(13).getStringCellValue());
					}

					//replace
					if(tempString.contains(charSeq1) || tempString.contains(charSeq2)){
						for(String s: mVariableNames.keySet()){
							String tmp = mVariableNames.get(s);
							tmp = "\\" + tmp.replace("\"", "\'");
							try{
								cell.setCellValue(cell.getStringCellValue().replaceAll("\\%\\{" + s + "\\}", tmp));
							}catch(Exception e){ e.printStackTrace(); }
						}
					}
				}
			}
		}

		//Remove
		System.out.println(rowsToDelete);
		ArrayList<Integer> lst = new ArrayList<Integer>();
		lst.addAll(rowsToDelete);
		ListIterator<Integer> it = lst.listIterator(lst.size());
		while(it.hasPrevious()){
			int beg = it.previous();
			sheet.createRow(beg);
			log.write("Remove row: " + beg + "\n");
		}
	}

	public void closeAll() throws Exception{
		doc.write(out);
		in.close();
		out.close();
		log.close();
	}

	public void fixPersistenceNums() throws IOException{
		System.out.println("fixPersistenceNums");
		log.write("\nfixPersistenceNums\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		String persCaseNum = "";
		Map<String, String> map = new TreeMap<String, String>();

		boolean replaceMode = false;

		int lastRow = sheet.getLastRowNum();
		for(int rowNum = 0; rowNum < lastRow; rowNum++){
			Cell cell8 = sheet.getRow(rowNum).getCell(8);
			if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING) continue;

			String cellContent = cell8.getStringCellValue();
			if(cellContent.matches("test case start")){
				persCaseNum = sheet.getRow(rowNum).getCell(0).getStringCellValue(); continue; }

			if(cellContent.matches("SaveExistMessage")){
				String cell6Content = sheet.getRow(rowNum).getCell(6).getStringCellValue();
				//String newValue = cell6Content.replaceAll("case\\d{1,2}", persCaseNum);
				String newValue = cell6Content.replaceAll("case\\d{1,2}", persCaseNum);
				map.put(cell6Content, newValue);
				sheet.getRow(rowNum).getCell(6).setCellValue(newValue);
				log.write(Calendar.getInstance().getTime() + "\tSaveExistMessage correction:\tRow:\t" + rowNum + "\tCell:\t6\tNewValue:\t" + newValue + "\n");
				continue;
			}

			if(replaceMode){
				int lastCell = sheet.getRow(rowNum).getLastCellNum();
				for(int cellNum = 0; cellNum < lastCell; cellNum++){
					Cell tempCell = sheet.getRow(rowNum).getCell(cellNum);
					if(tempCell == null || tempCell.getCellType() != Cell.CELL_TYPE_STRING) continue;

					String tempCellContent = tempCell.getStringCellValue();
					for(String key : map.keySet()){
						if(tempCellContent.contains(key + '.')){
							String newTempCellValue = tempCellContent.replaceAll(key, map.get(key));
							tempCell.setCellValue(newTempCellValue);
							log.write(Calendar.getInstance().getTime() + "\tCorrect Rers#: \tRow:\t" + rowNum + "\tCell:\t" + cellNum + "\tNewValue:\t" + newTempCellValue + "\n");
						}
					}
				}
			}


			if(cellContent.matches("LoadMessage")){
				replaceMode = true;
				String cell6Content = sheet.getRow(rowNum).getCell(6).getStringCellValue();
				for(String key : map.keySet()){
					if(cell6Content.contains(key)){
						String newCell6Content = cell6Content.replaceAll(key, map.get(key));
						sheet.getRow(rowNum).getCell(6).setCellValue(newCell6Content);
					}
				}
				continue;
			}
		}
	}

	public void addEmptyLineToTheEnd() throws IOException{
		System.out.println("addEmptyLineToTheEnd");
		log.write("\naddEmptyLineToTheEnd Func\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		for(int rowNum = 0; rowNum < sheet.getPhysicalNumberOfRows(); rowNum++){
			Cell cell8 = sheet.getRow(rowNum).getCell(8);
			if(cell8 == null || cell8.getCellType() != Cell.CELL_TYPE_STRING) continue;

			String cellContent = cell8.getStringCellValue();
			if(cellContent.matches("test case end")){
				Cell tempCell = sheet.getRow(rowNum - 1).getCell(8);
				if(tempCell == null || tempCell.getCellType() != Cell.CELL_TYPE_STRING) continue;

				String tempCellContent = tempCell.getStringCellValue();
				if(tempCellContent.matches("SaveExistMessage") || tempCellContent.matches("count")){
					log.write(Calendar.getInstance().getTime() + "\tAdd Row: " + rowNum + "\n");
					sheet.shiftRows(rowNum, sheet.getPhysicalNumberOfRows(), 1);
					sheet.createRow(rowNum);
				}
			}
		}
	}

	public void replaceKnownBug() throws IOException{
		System.out.println("replaceKnownBug");
		log.write("\nreplaceKnownBug\n");
		HSSFSheet sheet = doc.getSheetAt(0);
		int lastRow = sheet.getPhysicalNumberOfRows();
		for(int rowNum = 0; rowNum < lastRow; rowNum++){
			Row row = sheet.getRow(rowNum);
			if(row == null) continue;

			int lastCell = row.getLastCellNum();
			for(int columnNum = 0; columnNum < lastCell; columnNum++){
				Cell cell = sheet.getRow(rowNum).getCell(columnNum);
				if(cell == null || cell.getCellType() != Cell.CELL_TYPE_STRING) continue;

				String cellContent = cell.getStringCellValue();

				CharSequence expectEmpty = "KnownBug(\"" + '#' + "\")";
				CharSequence expectAny = "KnownBug(\"" + '*' + "\")";
				CharSequence check1 = ".Check(x)";
				CharSequence check2 = ".Check(";
				CharSequence null1 = ",null";
				CharSequence null2 = ", null";
				CharSequence knownBug = "KnownBug(";
				CharSequence knownBug2 = "KnownBug()";

				if(cellContent.contains("Known")){
					String newValue = cellContent;
					System.out.println(newValue);

					if(cellContent.contains(check1)) newValue = cellContent.replace(check1, "");

					if(newValue.contains(check2)) newValue = cellContent.replace(check2, ".Actual(");

					if(cellContent.contains(null1)){
						newValue = newValue.replace(null1, "");
						newValue = newValue.replace(".Bug", ".BugEmpty");
					}

					if(cellContent.contains(null2)){
						newValue = newValue.replace(null2, "");
						newValue = newValue.replace(".Bug", ".BugEmpty");
					}

					if(cellContent.contains(expectEmpty)) newValue = newValue.replace(expectEmpty, "ExpectedEmpty()");
					if(cellContent.contains(expectAny)) newValue = newValue.replace(expectAny, "ExpectedAny()");
					if(cellContent.contains(knownBug2)) newValue = newValue.replace(knownBug2, "ExpectedEmpty()");
					if(cellContent.contains(knownBug)) newValue = newValue.replace(knownBug, "Expected(");
					//System.out.println("Before Try: " + newValue);

					try{
						//Found Bug
						int comaStart = newValue.indexOf(".Bug(\"#");
						//System.out.println("comaStart: " + comaStart);

						if(comaStart == -1) {
							comaStart = newValue.indexOf(".BugEmpty(\"#");
							if(comaStart == -1) continue;
							comaStart += 11;
						} else { comaStart += 6; }
						//System.out.println("comaStart: " + comaStart);

						int comaEnd = newValue.indexOf("\"", comaStart);
						//System.out.println("comaEnd: " + comaEnd);


						if(!newValue.substring(comaStart, comaEnd).contains(".")) {
							cell.setCellValue(newValue);
							continue;
						}

						//Dot Found
						int dotPos = newValue.indexOf(".", comaStart);
						newValue = newValue.substring(0, dotPos) + newValue.substring(comaEnd);
						cell.setCellValue(newValue);
					}catch(Exception e){}
					//System.out.println("After Try: " + newValue);

					cell.setCellValue(newValue);
					log.write(Calendar.getInstance().getTime() + "\tRow:\t" + rowNum + ", Column:\t" + columnNum + "\tOldVal: " + cellContent + "\tNewVal: " + newValue + "\n");
				}
			}
		}
	}
}
