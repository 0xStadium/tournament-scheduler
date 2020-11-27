/*
 *June 2018 - August 2018
 *This program was made by Seong Su Park for the Orange County Badminton Club.
 *It creates excel / word documents of court schedules for tournaments.
 *These documents are designed to be printed and placed between courts
 *so that the players can easily see when the they need to play their games.
 *200+ players compete in these tournaments which are hosted by KBFSA (http://kbfsa.com/).
 */

import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.awt.Font;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFColor;
import javax.swing.*;

public class BadmintonScheduler {

	private JFrame frame;
	private JLabel errorLabel;
	private JButton seedButton;
	private ArrayList<Duo> partnerList;
	private ArrayList<ArrayList<Integer>> matchList;
	private ArrayList<JTextField> textList;
	
	public BadmintonScheduler() {

		// Setup GUI
		frame = new JFrame("Badminton Scheduler");
		frame.setSize(800, 800);
		frame.setLayout(new GridLayout(11, 3));
		ImageIcon img = new ImageIcon(getClass().getResource("/images/birdie.png"));
		frame.setIconImage(img.getImage());
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		// Fills in the 11 x 3 GridLayout
		createTextFields();

		// Activates Seed Button
		setUpSeedButton();

		// UI
		try {
			UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		// Makes Frame Visible
		frame.setVisible(true);
		
	}

	// Creates JTextFields
	public void createTextFields() {
		// Fonts
		Font teamfont = new Font("Microsoft Sans Serif", Font.BOLD, 14);
		Font namefont = new Font("Dialog", Font.BOLD, 20);

		// List of JextFields
		textList = new ArrayList<JTextField>();

		for (int i = 0; i < 20; i++) {
			// Up to 10 Teams can register , 20/2 = 10
			if (i % 2 == 0) { // This loops 10 times to set the Team Labels
				int teamnumber = i / 2 + 1;
				JTextField teamlabel = new JTextField("Team " + teamnumber + " ");
				teamlabel.setFont(teamfont);
				teamlabel.setHorizontalAlignment(JTextField.CENTER);
				teamlabel.setFocusable(false);
				teamlabel.setEditable(false);
				frame.add(teamlabel);
			}

			// 10 teams x 2 per team = 20, namefield is created 20 times
			JTextField namefield = new JTextField("");
			namefield.setFont(namefont);
			namefield.setHorizontalAlignment(JTextField.CENTER);

			/*
			 * Clicking on the namefield sets it to blank for faster typing, not useful
			 * right now, maybe for future changes namefield.addMouseListener(new
			 * MouseAdapter() {
			 * 
			 * @Override public void mouseClicked(MouseEvent e) { namefield.setText(""); }
			 * });
			 */

			frame.add(namefield);
			textList.add(namefield);
		}
	}

	public void setUpSeedButton() {
		// Fonts
		Font f = new Font("Microsoft Sans Serif", Font.BOLD, 24);
		Font fa = new Font("Microsoft Sans Serif", Font.BOLD, 22);

		// Blank JTextField for Looks
		JTextField blanklabel = new JTextField("");
		blanklabel.setEditable(false);
		frame.add(blanklabel);

		// Finish Button
		seedButton = new JButton("Finish");
		seedButton.setFont(f);

		// Error Label
		errorLabel = new JLabel("");
		errorLabel.setHorizontalAlignment(JLabel.CENTER);
		errorLabel.setFont(fa);

		// Create PartnerList
		partnerList = new ArrayList<Duo>();

		seedButton.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent event) {

				// First checks number of blank spaces
				int i = 0;
				for (int x = 0; x < textList.size(); x++) {
					if (!textList.get(x).getText().isEmpty()) {
						i++;
					}
				}

				if (i == 0 || i % 2 == 1) { // Error
					errorLabel.setText("Error");
				} else { // Valid input of partner names
					for (int j = 0; j < textList.size(); j += 2) {
						if (textList.get(j).getText().isEmpty()) {
							j += textList.size();
						} else {
							// Creates Teams
							Duo duo = new Duo(textList.get(j).getText(), textList.get(j + 1).getText());
							partnerList.add(duo);
						}
					}

					if (partnerList.size() == 1) { // There can't be 1 team in a tournament
						errorLabel.setText("Error");
					} else {
						// Puts each team in their respective matches
						makeSeed();
					}
				}
			}
		});
		frame.add(seedButton);
		frame.add(errorLabel);
	}

	/*
	 * Creates sequential matches between 2 to 10 teams in a single round robin
	 * tournament so that there are few occasions where teams must have consecutive
	 * matches because this is unfair for those teams due to them not getting a
	 * break between their games
	 */
	public void makeSeed() {
		// https://en.wikipedia.org/wiki/Round-robin_tournament#Scheduling_algorithm
		// matchList holds arrayLists of Integers where Integers represent the team
		// number matchList is in order of sequential matches
		matchList = new ArrayList<ArrayList<Integer>>();

		int numberOfTeams = partnerList.size();
		int totalMatches = 0;
		int roundDiv = 0;
		ArrayList<Integer> rotateList = new ArrayList<Integer>();
		if (numberOfTeams % 2 == 0) {
			totalMatches = ((numberOfTeams) * (numberOfTeams - 1)) / 2;
			roundDiv = numberOfTeams - 1;
		} else {
			totalMatches = ((numberOfTeams + 1) * (numberOfTeams)) / 2;
			roundDiv = numberOfTeams;
		}

		for (int i = 2; i <= numberOfTeams; i++) {
			rotateList.add(i);
		}

		if (numberOfTeams % 2 != 0) {
			rotateList.add(-1);
		}

		for (int j = 0; j < roundDiv; j++) {
			if (rotateList.get(rotateList.size() - 1) != -1) {
				ArrayList<Integer> firstRoundMatch = new ArrayList<Integer>();
				firstRoundMatch.add(1);
				firstRoundMatch.add(rotateList.get(rotateList.size() - 1));
				matchList.add(firstRoundMatch);
			}

			for (int x = 0; x < (rotateList.size() - 1) / 2; x++) {
				if (rotateList.get(x) != -1 && rotateList.get(rotateList.size() - 2 - x) != -1) {
					ArrayList<Integer> match = new ArrayList<Integer>();
					match.add(rotateList.get(x));
					match.add(rotateList.get(rotateList.size() - 2 - x));
					matchList.add(match);
				}
			}
			int end = rotateList.remove(rotateList.size() - 1);
			rotateList.add(0, end);
		}

		if (numberOfTeams == 5) {
			matchList.add(matchList.remove(8));
			matchList.add(7, matchList.remove(6));
		}
		
		// Creates Excel and Word Documents
		try {
			// Create Excel document
			XSSFWorkbook workbook = new XSSFWorkbook();

			int totalSheets;
			int calc = matchList.size() / 15;
			if (matchList.size() % 15 >= 1) {
				totalSheets = calc + 1;
			} else {
				totalSheets = calc;
			}

			int matchCount = 0;

			for (int i = 1; i <= totalSheets; i++) {
				// Create Excel Sheet
				XSSFSheet sheet = workbook.createSheet("Sheet" + i);

				// Create Excel cellStyle
				XSSFCellStyle cellStyle = workbook.createCellStyle();

				// Black Borders
				cellStyle.setBorderBottom(BorderStyle.THIN);
				cellStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());
				cellStyle.setBorderLeft(BorderStyle.THIN);
				cellStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
				cellStyle.setBorderRight(BorderStyle.THIN);
				cellStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
				cellStyle.setBorderTop(BorderStyle.THIN);
				cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());

				// Create Excel Font
				XSSFFont font = workbook.createFont();
				font.setFontHeightInPoints((short) 13);
				font.setFontName("Calibri");
				font.setBold(true);
				cellStyle.setFont(font);

				// cellStyle centered
				cellStyle.setAlignment(HorizontalAlignment.CENTER);
				cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

				// Create row1
				XSSFRow row1 = sheet.createRow(0);
				row1.setHeight((short) 550);

				// Create row2
				XSSFRow row2 = sheet.createRow(1);
				row2.setHeight((short) 450);

				// First sheet will have title specifics
				if (i == 1) {

					// Row 1 position
					XSSFCell empty1 = (XSSFCell) row1.createCell(0);
					XSSFCell titleCell = (XSSFCell) row1.createCell(1);

					// Row 2 position
					XSSFCell emptyrow2a = (XSSFCell) row2.createCell(0);
					XSSFCell emptyrow2b = (XSSFCell) row2.createCell(1);
					XSSFCell subtitleCell = (XSSFCell) row2.createCell(2);
					XSSFCell extra = (XSSFCell) row2.createCell(4);

					// Merge cells for better title typing
					sheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 6));
					sheet.addMergedRegion(new CellRangeAddress(1, 1, 2, 3));
					sheet.addMergedRegion(new CellRangeAddress(1, 1, 4, 6));

					// New font for titles and extras
					XSSFCellStyle titleStyle = workbook.createCellStyle();
					XSSFFont titleFont = workbook.createFont();
					titleFont.setFontHeightInPoints((short) 20);
					titleFont.setFontName("Calibri");
					titleFont.setBold(true);
					titleStyle.setFont(titleFont);
					titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
					titleCell.setCellValue("COURT#[]");
					titleCell.setCellStyle(titleStyle);
					XSSFCellStyle subtitleStyle = workbook.createCellStyle();
					subtitleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
					XSSFFont subtitleFont = workbook.createFont();
					subtitleFont.setFontHeightInPoints((short) 14);
					subtitleFont.setFontName("Calibri");
					subtitleFont.setBold(true);
					subtitleStyle.setFont(subtitleFont);
					subtitleCell.setCellValue("**");
					subtitleCell.setCellStyle(subtitleStyle);

					XSSFCellStyle extraTitle = workbook.createCellStyle();
					XSSFFont extraFont = workbook.createFont();
					extraFont.setFontHeightInPoints((short) 12);
					extraFont.setFontName("Calibri");
					extraFont.setBold(true);
					extraTitle.setVerticalAlignment(VerticalAlignment.CENTER);
					extraTitle.setFont(extraFont);
					extra.setCellStyle(extraTitle);
					extra.setCellValue("[]");

				}

				// Create descriptionStyle
				XSSFCellStyle descriptionStyle = (XSSFCellStyle) cellStyle.clone();
				XSSFColor color = new XSSFColor();
				color.setARGBHex("00B0F0");
				descriptionStyle.setFillForegroundColor(color);
				descriptionStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

				// Create row3
				XSSFRow row3 = sheet.createRow(2);
				row3.setHeight((short) 1000);
				XSSFCell numberCell = (XSSFCell) row3.createCell(0);
				numberCell.setCellValue("#");
				numberCell.setCellStyle(descriptionStyle);

				XSSFCell winTeamCell = (XSSFCell) row3.createCell(1);
				winTeamCell.setCellValue("W");
				winTeamCell.setCellStyle(descriptionStyle);

				XSSFCell loseTeamCell = (XSSFCell) row3.createCell(2);
				loseTeamCell.setCellValue("L");
				loseTeamCell.setCellStyle(descriptionStyle);

				XSSFCell descriptionCell = (XSSFCell) row3.createCell(3);
				descriptionCell.setCellValue("Description");
				descriptionCell.setCellStyle(descriptionStyle);

				XSSFCell winnerCell = (XSSFCell) row3.createCell(4);
				winnerCell.setCellValue("Winner");
				winnerCell.setCellStyle(descriptionStyle);

				XSSFCell winScoreCell = (XSSFCell) row3.createCell(5);
				winScoreCell.setCellValue("WS");
				winScoreCell.setCellStyle(descriptionStyle);

				XSSFCell loseScoreCell = (XSSFCell) row3.createCell(6);
				loseScoreCell.setCellValue("LS");
				loseScoreCell.setCellStyle(descriptionStyle);

				// Set Excel Column Widths and Center
				sheet.setColumnWidth(0, 1200);
				sheet.setColumnWidth(1, 1200);
				sheet.setColumnWidth(2, 1200);
				sheet.setColumnWidth(3, 11700);
				sheet.setColumnWidth(4, 3500);
				sheet.setColumnWidth(5, 1200);
				sheet.setColumnWidth(6, 1200);

				sheet.setHorizontallyCenter(true);
				sheet.setVerticallyCenter(true);

				// 15 rows of matches per sheet

				for (int sheetMatch = 0; sheetMatch < 15; sheetMatch++) {

					// Create row and cells
					XSSFRow row = sheet.createRow(3 + sheetMatch);
					row.setHeight((short) 800);
					XSSFCell roundCell = (XSSFCell) row.createCell(0);
					roundCell.setCellStyle(cellStyle);

					XSSFCell winCell = (XSSFCell) row.createCell(1);
					winCell.setCellStyle(cellStyle);

					XSSFCell loseCell = (XSSFCell) row.createCell(2);
					loseCell.setCellStyle(cellStyle);

					XSSFCell descrCell = (XSSFCell) row.createCell(3);
					descrCell.setCellStyle(cellStyle);

					// Fills in descrCell with the match details if there are matches remaining,
					// leaves blank if not
					if (matchCount < matchList.size()) {

						roundCell.setCellValue(matchCount + 1);
						int firstTeam = matchList.get(matchCount).get(0);
						int secondTeam = matchList.get(matchCount).get(1);
						String s = firstTeam + " " + partnerList.get(firstTeam - 1).getPerson1() + "/"
								+ partnerList.get(firstTeam - 1).getPerson2() + "  vs  " + secondTeam + " "
								+ partnerList.get(secondTeam - 1).getPerson1() + "/"
								+ partnerList.get(secondTeam - 1).getPerson2();
						descrCell.setCellValue(s);
						descrCell.setCellStyle(cellStyle);

						matchCount++;

						// Ensures that the team player names don't get cut off if they are long
						if (descrCell.getCellStyle().getShrinkToFit() == false) {
							XSSFCellStyle shrinkStyle = (XSSFCellStyle) cellStyle.clone();
							shrinkStyle.setShrinkToFit(true);
							descrCell.setCellStyle(shrinkStyle);
						}
					}

					XSSFCell emptyWinner = (XSSFCell) row.createCell(4);
					emptyWinner.setCellStyle(cellStyle);

					XSSFCell emptyWon = (XSSFCell) row.createCell(5);
					emptyWon.setCellStyle(cellStyle);

					XSSFCell emptyLost = (XSSFCell) row.createCell(6);
					emptyLost.setCellStyle(cellStyle);
				}
			}
			// Completes Excel Document

			// Creates Word document
			XWPFDocument document = new XWPFDocument();

			// Court title section
			XWPFParagraph paragraph = document.createParagraph();
			XWPFRun run1 = paragraph.createRun();
			
			run1.setCapitalized(true);
			run1.setFontFamily("Calibri");
			run1.setFontSize(42);
			run1.setBold(true);
			run1.setUnderline(UnderlinePatterns.SINGLE);
			run1.setText("COURT #");

			// Large title section
			XWPFParagraph paragraph2 = document.createParagraph();
			paragraph2.setAlignment(ParagraphAlignment.CENTER);
			XWPFRun run2 = paragraph2.createRun();
			run2.setCapitalized(true);
			run2.setFontFamily("Calibri");
			run2.setFontSize(80);
			run2.setBold(true);
			run2.setText("A");

			// Team name section
			XWPFParagraph paragraph3 = document.createParagraph();
			paragraph3.setSpacingBetween(1.2);

			int partnerCount = 0;
			for (int i = 1; i <= 10; i++) {
				XWPFRun run3 = paragraph3.createRun();
				run3.setCapitalized(true);
				run3.setFontFamily("Calibri");
				run3.setFontSize(22);
				run3.setBold(true);

				if (partnerCount < partnerList.size()) {
					run3.addTab();
					run3.setText(i + "  " + partnerList.get(partnerCount).getPerson1() + " / "
							+ partnerList.get(partnerCount).getPerson2());
					run3.addBreak();
					partnerCount++;
				} else {
					run3.addBreak();
				}
				
				// Extra section
				if (i == 10) {
					run3.addTab();
					run3.setText("*   *");
				}
			}
			
			JFileChooser fileChooser = new JFileChooser();
			fileChooser.setDialogTitle("Choose a file");
			JFrame parentFrame = new JFrame();

			int userSelection = fileChooser.showSaveDialog(parentFrame);
			File fileToSave = null;
			if (userSelection == JFileChooser.APPROVE_OPTION) {
				fileToSave = fileChooser.getSelectedFile();
				System.out.println("Save as file: " + fileToSave.getAbsolutePath());
			}

			// Completes excel document
			FileOutputStream out = new FileOutputStream(fileToSave.getAbsolutePath() + ".xlsx");
			workbook.write(out);
			workbook.close();

			// Completes Word document
			out = new FileOutputStream(fileToSave.getAbsolutePath() + ".docx");
			document.write(out);
			out.close();
			document.close();
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		reset();
	}

	public void reset() {
		frame.getContentPane().removeAll();
		frame.setSize(800, 800);
		frame.setLayout(new GridLayout(11, 3));
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		// Fills in the 11 x 3 GridLayout
		createTextFields();

		// Activates Seed Button
		setUpSeedButton();

		// Makes Frame Visible
		frame.setVisible(true);
	}

	public static void main(String[] args) {
		BadmintonScheduler score = new BadmintonScheduler();
	}
}
