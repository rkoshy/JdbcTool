/**
 * JdbcTool: Command line pain relief for JDBC databases.
 * Copyright (C) 2007, Quuxo Software.
 * JdbcTool 2.0 - XLS, CSV, TEXT, HTML output and multi-tab features
 * Copyright 2009-2017 Renny Koshy
 *
 * This program is free software; you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation; either version 2 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License along
 * with this program; if not, write to the Free Software Foundation, Inc.,
 * 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
 *
 */

package com.quuxo.jdbctool;

import java.io.BufferedInputStream;
import java.io.BufferedReader;
import java.io.EOFException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.SQLWarning;
import java.sql.Statement;
import java.sql.Types;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.StringTokenizer;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.gnu.readline.Readline;
import org.gnu.readline.ReadlineLibrary;

import gnu.getopt.Getopt;

/**
 * Command line program to execute SQL statements interactively.
 * 
 * @author <a href="mailto:michael@quuxo.com">Michael Gratton</a>
 */
public class JdbcTool {
	private static final String DEFAULT_CSS_FILE = "style.css";
	protected String[] driverNames = new String[] { "org.hsqldb.jdbcDriver", "net.sourceforge.jtds.jdbc.Driver",
			"com.mysql.jdbc.Driver", "org.postgresql.Driver", "net.snowflake.client.jdbc.SnowflakeDriver" };
	protected Connection connection;
	protected Statement statement;
	protected char eol = System.getProperty("line.separator").charAt(0);
	protected String prompt;
	protected static OutputFormat outputFormat = OutputFormat.TEXT;
	private static String title = null;
	protected String history;
	private HSSFWorkbook workbook;
	private FileOutputStream outputStream;
	private HSSFSheet currentSheet;
	private HSSFCellStyle textCellStyle;
	private HSSFCellStyle floatingCellStyle;
	private HSSFCellStyle integerCellStyle;
	private HSSFCellStyle headerCellStyle;
	private HSSFFont titleFont;
	private HashMap<String, HSSFCellStyle> cellStyles = new HashMap<String, HSSFCellStyle>();
	private static String currentSheetName = "";
	private int resultSetNum = 0;
	private int currentSheetCounter = 0;
	private static boolean showResultsOnly = false;
	private static String cssFile;
	private static boolean quiet;
	private static String outputFile = null;

	private static PrintStream output = null;
	private static boolean headings = true;
	private static boolean append = false;
	private static List<String> tabNames = new ArrayList<String>();
	private static List<String> tabTitles = new ArrayList<String>();
	private static boolean incrementTab = false;

	private enum OutputFormat {
		TEXT, HTML, XLS, CSV
	};

	public JdbcTool(String url, String username, String password) throws SQLException, Exception {
		loadDrivers();
		this.connection = DriverManager.getConnection(url, username, password);
		setPrompt(url.startsWith("jdbc:") ? url.substring("jdbc:".length()) : url);
		this.statement = this.connection.createStatement();
		this.history = System.getProperty("user.home") + "/.jdbctool_history";

		// readline init
		try {
			Readline.load(ReadlineLibrary.GnuReadline);
		} catch (UnsatisfiedLinkError ule) {
			if (!quiet)
				System.err.println("Java-Readline not found, using simple stdin.");
		}
		Readline.initReadline("JdbcTool");
		try {
			Readline.readHistoryFile(this.history);
		} catch (Exception e) {
			e.printStackTrace();
			// oh well
		}

		if (outputFile != null) {
			if (outputFormat != OutputFormat.XLS) {
				// Everything else
				try {
					outputStream = new FileOutputStream(outputFile);
					output = new PrintStream(outputStream);
				} catch (Exception e) {
					System.err.println("Could not open output file '" + outputFile + "'");
					e.printStackTrace(System.err);
					throw e;
				}
			}
		} else {
			output = System.out;
		}
	}

	protected void loadDrivers() {
		for (String name : this.driverNames) {
			try {
				Class.forName(name);
			} catch (Exception e) {
				System.err.println("Did not load " + name);
			}
		}
	}

	public void setPrompt(String prompt) {
		if (prompt == null) {
			prompt = "jdbctool";
		}
		this.prompt = prompt + "> ";
	}

	public void close() throws Exception {
		this.connection.close();
		try {
			Readline.writeHistoryFile(this.history);
		} catch (Exception e) {
			// oh well
		}
		Readline.cleanup();
	}

	public void start() throws Exception {
		// let's go!
		boolean originalAppend = append;
		while (true) {
			try {
				String line;
				append = originalAppend;
				line = Readline.readline(quiet ? "" : this.prompt);
				if (line != null) {
					if (line.equalsIgnoreCase("quit") || line.equalsIgnoreCase("exit")) {
						throw new EOFException();
					}
					if (outputFormat == OutputFormat.XLS) {
						try {
							if (!append) {
								workbook = new HSSFWorkbook();
								outputStream = new FileOutputStream(outputFile);
							} else if (workbook == null) {
								try {
									POIFSFileSystem fileSystem = new POIFSFileSystem(new FileInputStream(outputFile));
									// Read in the data, since we're about to
									// overwrite
									workbook = new HSSFWorkbook(fileSystem);
									// Make sure we can open it for output
									outputStream = new FileOutputStream(outputFile);
									// set headers = off, turn off titles
									headings = false;
									title = null;
								} catch (IOException e) {
									// File does not exist, so try to create it
									originalAppend = true;
									append = false;
									workbook = new HSSFWorkbook();
									outputStream = new FileOutputStream(outputFile);
								}

							}
							textCellStyle = workbook.createCellStyle();
							// stextCellStyle.setAlignment(HSSFCellStyle.ALIGN_FILL);

							floatingCellStyle = workbook.createCellStyle();
							floatingCellStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
							HSSFDataFormat format = workbook.createDataFormat();
							floatingCellStyle.setDataFormat(format.getFormat("###,###,###,##0.00"));

							integerCellStyle = workbook.createCellStyle();
							integerCellStyle.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
							integerCellStyle.setDataFormat(format.getFormat("###########0"));

							titleFont = workbook.createFont();
							titleFont.setFontHeightInPoints((short) 16);
							titleFont.setFontName("Arial Black");
							titleFont.setItalic(true);

							headerCellStyle = workbook.createCellStyle();
							headerCellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
							headerCellStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
							headerCellStyle.setFont(titleFont);
						} catch (Exception e) {
							System.err.println("Could not open output file '" + outputFile + "'");
							e.printStackTrace(System.err);
							throw e;
						}
					}
					if (!incrementTab)
						resultSetNum = 0;

					execute(line);
					// If there is a workbook, write it out to file
					if (workbook != null) {
						workbook.write(outputStream);
						outputStream.flush();
					}
					// Close the output writer
					if (output != null) {
						output.close();
					}
				}
			} catch (EOFException e) {
				System.out.println();
				break;
			} catch (SQLException e) {
				printLineError(e.toString());
			}
		}
	}

	protected void execute(String line) throws SQLException {
		boolean hasResults = this.statement.execute(line);
		printResults();
	}

	protected void printResults() throws SQLException {
		printPageHeader();
		while (true) {
			ResultSet results = this.statement.getResultSet();
			if (results != null) {
				exhumeWarnings(results);

				resultSetNum++;

				ResultSetMetaData meta = results.getMetaData();
				int cols = meta.getColumnCount();

				String[] names = new String[cols];
				int[] types = new int[cols];
				for (int i = 0; i < cols; i++) {
					names[i] = meta.getColumnName(i + 1);
					types[i] = meta.getColumnType(i + 1);
					System.out.println(i + " type " + types[i]);
				}

				int[] widths = new int[cols];
				int maxRows = 50000;
				ArrayList<String[]> rows = new ArrayList<String[]>(maxRows);
				while (results != null) {
					// init widths
					for (int i = 0; i < cols; i++) {
						widths[i] = names[i].length();
					}

					// get values
					for (int row = 0; row < maxRows; row++) {
						if (!results.next()) {
							results.close();
							results = null;
							break;
						}
						String[] values = new String[cols];
						for (int i = 0; i < cols; i++) {
							values[i] = results.getString(i + 1);
							if (values[i] == null) {
								values[i] = "<NULL>";
							}
							widths[i] = Math.max(widths[i], values[i].length());
						}
						rows.add(values);
						exhumeWarnings(results);
					}

					// print it
					printHeaders(names, widths);
					for (String[] values : rows) {
						printRow(values, widths, types);
					}
					printFooter(widths);
				}
			} else {
				if (!showResultsOnly) {
					output.println();
					output.println("Updated: " + this.statement.getUpdateCount());
					output.println();
				}
			}
			// Advance and quit if done
			if ((this.statement.getMoreResults() == false) && (this.statement.getUpdateCount() == -1))
				break;
		}
		printPageFooter();
	}

	private void printPageHeader() {
		if (headings) {
			switch (outputFormat) {
			case HTML:
				output.println("<html><head><title>" + ((title != null) ? title : "")
						+ "</title><meta http-equiv=\"content-type\" content=\"text/html;charset=UTF-8\"/></head>");
				if (cssFile != null) {
					BufferedInputStream in;
					try {
						in = new BufferedInputStream(new FileInputStream(new File(cssFile)));
						BufferedReader reader = new BufferedReader(new InputStreamReader(in));
						String line = null;
						while ((line = reader.readLine()) != null) {
							output.println(line);
						}
					} catch (IOException x) {
						System.err.println(x);
					}
				} else {
					try {
						InputStream in = this.getClass().getClassLoader().getResourceAsStream(DEFAULT_CSS_FILE);
						BufferedReader reader = new BufferedReader(new InputStreamReader(in));
						String line = null;
						while ((line = reader.readLine()) != null) {
							output.println(line);
						}
					} catch (IOException x) {
						System.err.println(x);
					}
				}
				output.println("<body>");
				if (title != null) {
					output.println("<table width=\"100%\">");
					output.println("<tr><th class=\"title\">" + title + "</th>");
					output.println("</table>");
				}
				break;
			case CSV:
				if (title != null) {
					output.println(title);
				}
				break;
			case XLS:
				if (title != null) {
					// TODO: What do we do here?
				}
				break;
			}
		}
	}

	private void printPageFooter() {
		if (outputFormat == OutputFormat.HTML) {
			output.println("</body>");
			output.println("</html>");
		}
	}

	protected void printHeaders(String[] values, int[] widths) {
		switch (outputFormat) {
		case HTML:
			output.println("<table width=\"100%\">");
			output.print("<tr>");
			if (headings) {
				for (int i = 0; i < values.length; i++) {
					output.print("<th align=\"center\">");
					output.print(values[i]);
					output.print("</th>");
				}
				output.println("</tr>");
			}
			break;
		case CSV:
			if (headings) {
				for (int i = 0; i < values.length; i++) {
					if (i > 0) {
						output.print(",");
					}
					output.print(values[i]);
				}
			}
			break;
		case TEXT:
			if (headings) {
				printLine(widths);
				output.print("|");
				for (int i = 0; i < values.length; i++) {
					output.print(" ");
					output.print(values[i]);
					for (int w = values[i].length(); w < widths[i]; w++) {
						output.print(" ");
					}
					output.print(" |");
				}
				output.println();
				printLine(widths);
			}
			break;
		case XLS:
			// Start a new sheet, or pick the correct sheet
			boolean newSheet = true;
			// There is a map entry
			String myTitle = title;

			int mySheet = tabNames.indexOf(currentSheetName);
			if (mySheet == -1) {
				mySheet = resultSetNum - 1;
			}

			if (mySheet < tabTitles.size()) {
				myTitle = tabTitles.get(mySheet);
			}
			if (mySheet < tabNames.size()) {
				if (workbook.getNumberOfSheets() + 1 < mySheet) {
					// Create all intermediate sheets
					for (int i = workbook.getNumberOfSheets() + 1; i < mySheet; i++) {
						String name = tabNames.get(i);
						System.out.println("Looking for " + name);
						currentSheet = workbook.getSheet(name);
						if (currentSheet == null) {
							System.out.println("Creating " + name);
							currentSheet = workbook.createSheet(name);
							String title = null;
							if (i < tabTitles.size()) {
								title = tabTitles.get(i);
							}
							if (title != null) {
								HSSFRow row = currentSheet.createRow(0);
								setCellWithFormatting(row, 0, title);
							}
						}
					}
				}

				String name = tabNames.get(mySheet);
				currentSheet = workbook.getSheet(name);
				if (currentSheet == null) {
					currentSheet = workbook.createSheet(name);
				} else {
					newSheet = false;
				}
			} else {
				currentSheet = workbook.createSheet();
			}

			if ((!append || incrementTab) && newSheet) {
				if (myTitle != null) {
					// Add a row
					HSSFRow row = currentSheet.createRow(0);
					// HSSFCell cell = row.createCell(0);
					// cell.setCellValue(new HSSFRichTextString(myTitle));
					// cell.setCellStyle(headerCellStyle);

					setCellWithFormatting(row, 0, myTitle);
					//
					// currentSheet.addMergedRegion(new CellRangeAddress(0, //
					// first
					// // row
					// // (0-based)
					// 0, // last row (0-based)
					// 0, // first column (0-based)
					// values.length - 1)); // last column (0-based)
					// create the next row
					// currentSheet.createRow(currentSheet.getLastRowNum());
				}
				if (headings) {
					// Add a row
					HSSFRow row = currentSheet.createRow(currentSheet.getLastRowNum() + 1);
					for (int i = 0; i < values.length; i++) {
						row.createCell(i).setCellValue(new HSSFRichTextString(values[i]));
					}
				}
			}
			break;
		}
	}

	protected void printRow(String[] values, int[] widths, int[] types) {
		switch (outputFormat) {
		case HTML:
			output.print("<tr>");
			for (int i = 0; i < values.length; i++) {
				output.print("<td align=\"center\">");
				output.print(values[i]);
				output.print("</td>");
			}
			output.println("</tr>");
			break;
		case CSV:
			for (int i = 0; i < values.length; i++) {
				if (i > 0) {
					output.print(",");
				}
				output.print(values[i]);
			}
			output.println();
			break;
		case TEXT:
			output.print("|");
			for (int i = 0; i < values.length; i++) {
				output.print(" ");
				output.print(values[i]);
				for (int w = values[i].length(); w < widths[i]; w++) {
					output.print(" ");
				}
				output.print(" |");
			}
			output.println();
			break;
		case XLS:
			// Add a row
			HSSFRow row = currentSheet.createRow(currentSheet.getLastRowNum() + 1);
			for (int i = 0; i < values.length; i++) {
				if (!values[i].equals("<NULL>")) {
					// if (values[i].matches("^((-|\\+)?[0-9]+(\\.[0-9]+)?)$"))
					// {
					if ((types[i] == Types.BIGINT) || (types[i] == Types.DECIMAL) || (types[i] == Types.DOUBLE)
							|| (types[i] == Types.FLOAT) || (types[i] == Types.INTEGER) || (types[i] == Types.NUMERIC)
							|| (types[i] == Types.REAL) || (types[i] == Types.TINYINT)) {

						HSSFCell cell = row.createCell(i);
						cell.setCellValue(Double.parseDouble(values[i]));
						cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);

						if ((types[i] == Types.DECIMAL) || (types[i] == Types.DOUBLE) || (types[i] == Types.FLOAT)
								|| (types[i] == Types.NUMERIC) || (types[i] == Types.REAL)) {
							cell.setCellStyle(floatingCellStyle);
						} else {
							cell.setCellStyle(integerCellStyle);
						}
					} else {
						if (values[i].startsWith("{")) {
							setCellWithFormatting(row, i, values[i]);
						} else {
							HSSFCell cell = row.createCell(i);
							cell.setCellValue(values[i]);
							cell.setCellType(HSSFCell.CELL_TYPE_STRING);
							cell.setCellStyle(textCellStyle);
						}
					}
				}
			}
			break;
		}
	}

	private void setCellWithFormatting(HSSFRow row, int column, String value) {
		int index = 0;
		boolean bold = false;
		boolean underline = false;
		boolean italic = false;
		boolean center = false;
		int heading = 5; // Default "heading5" == 10pt which
		// is normal

		boolean mergeSpec = false;
		int mergeCells = 0;
		while ((++index < value.length()) && (value.charAt(index) != '}')) {
			switch (value.charAt(index)) {
			case 'u':
			case 'U':
				underline = true;
				break;
			case 'b':
			case 'B':
				bold = true;
				break;
			case 'i':
			case 'I':
				italic = true;
				break;
			case 'c':
			case 'C':
				center = true;
				break;
			case '>': // Merge columns
				mergeSpec = true;
				break;
			default:
				char c = value.charAt(index);
				if (!mergeSpec) {
					if ((c >= '1') && (c <= '4')) {
						heading = c - '0';
					}
				} else {
					mergeCells = c - '0';
				}
				break;

			}
		}
		if ((index + 1) < value.length()) {
			value = value.substring(index + 1);
		} else {
			value = "";
		}

		String key = ((heading > 0) ? String.valueOf((char) ('0' + heading)) : "") + (bold ? "B" : "")
				+ (italic ? "I" : "") + (underline ? "U" : "");

		HSSFCellStyle style = cellStyles.get(key);
		if (style == null) {
			// Create it
			style = workbook.createCellStyle();
			HSSFFont font = workbook.createFont();

			if (bold) {
				font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			}
			if (underline) {
				font.setUnderline(HSSFFont.U_SINGLE);
			}
			font.setItalic(italic);
			font.setFontHeightInPoints((short) (20 - 2 * heading));

			style.setFont(font);
			if (center) {
				style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			}
			cellStyles.put(key, style);
		}

		HSSFCell cell = row.createCell(column);
		cell.setCellValue(value);
		cell.setCellStyle(style);
		cell.setCellType(HSSFCell.CELL_TYPE_STRING);
		if (mergeSpec) {
			currentSheet.addMergedRegion(new CellRangeAddress(row.getRowNum(), row.getRowNum(), cell.getColumnIndex(),
					cell.getColumnIndex() + mergeCells));
		}
	}

	protected void printLine(int[] widths) {
		if (outputFormat == OutputFormat.TEXT) {
			output.print("-");
			for (int i = 0; i < widths.length; i++) {
				for (int w = -3; w < widths[i]; w++) {
					output.print("-");
				}
			}
			output.println();
		}
	}

	protected void printFooter(int[] widths) {
		switch (outputFormat) {
		case HTML:
			output.println("</table>");
			output.println();
			break;
		case TEXT:
			output.print("-");
			for (int i = 0; i < widths.length; i++) {
				for (int w = -3; w < widths[i]; w++) {
					output.print("-");
				}
			}
			output.println();
			break;
		case XLS:
			for (int i = 0; i < widths.length; i++) {
				currentSheet.autoSizeColumn((short) i); // adjust width of the
				// first column
			}
		}
	}

	protected void exhumeWarnings(ResultSet results) throws SQLException {
		SQLWarning w = results.getWarnings();
		while (w != null) {
			printLineWarning(w.toString());
			w = w.getNextWarning();
		}
		results.clearWarnings();
	}

	protected static String promptUser(String message) {
		try {
			BufferedReader in = new BufferedReader(new InputStreamReader(System.in));
			if (!quiet)
				System.err.println(message + ": ");
			return in.readLine();
		} catch (IOException ioe) {
			printError(ioe.getMessage());
			return null;
		}
	}

	protected void printLineWarning(String message) {
		System.out.println("Warning: " + message);
	}

	protected void printLineError(String message) {
		System.out.println("Error: " + message);
	}

	protected static void printError(String message) {
		System.err.println(message);
	}

	public static void main(String[] argv) {
		String password = null;
		String user = null;

		Getopt g = new Getopt("JdbcTool", argv, "p:Pu:hf:o:t:s:rqaT:S:i");
		int c;
		String tabs = null;
		while ((c = g.getopt()) != -1) {
			switch (c) {
			case 'p':
				password = g.getOptarg();
				break;
			case 'q':
				quiet = true;
				break;
			case 'h':
				headings = false;
				break;
			case 'f':
				String temp = g.getOptarg();
				if (temp.equalsIgnoreCase("xls")) {
					outputFormat = OutputFormat.XLS;
				} else if (temp.equalsIgnoreCase("html")) {
					outputFormat = OutputFormat.HTML;
				} else if (temp.equalsIgnoreCase("CSV")) {
					outputFormat = OutputFormat.CSV;
				}
				break;
			case 'a':
				append = true;
				break;
			case 'o':
				outputFile = g.getOptarg();
				break;
			case 't':
				title = g.getOptarg();
				if (!title.startsWith("{")) {
					title = "{BUC3>6}" + title;
				}
				break;
			case 'T':
				tabs = g.getOptarg();
				break;
			case 's':
				cssFile = g.getOptarg();
				break;
			case 'S':
				currentSheetName = g.getOptarg();
				break;
			case 'P':
				password = promptUser("Enter password");
				break;
			case 'r':
				showResultsOnly = true;
				break;
			case 'i':
				incrementTab = true;
				break;
			case 'u':
				user = g.getOptarg();
				break;
			default:
				printError("JdbcTool: unknown option `" + c + "'");
				System.exit(-1);
			}
		}

		int i = g.getOptind();
		if (i >= argv.length) {
			printError("No JDBC URL specified.");
			System.exit(-1);
		}

		String url = argv[i];

		if ((outputFormat == OutputFormat.XLS) && (tabs != null) && (tabs.matches("(\\[[^\\[]*\\])*"))) {
			// tab-names are specified
			// Correct format for tab names
			tabs = tabs.substring(1, tabs.length() - 1);
			StringTokenizer t = new StringTokenizer(tabs, "][");
			while (t.hasMoreTokens()) {
				// Found a tab... are the query's specified?
				StringTokenizer t2 = new StringTokenizer(t.nextToken(), "|");
				String name = t2.nextToken();
				tabNames.add(name); // Name
				if (t2.hasMoreTokens()) {
					String title = t2.nextToken();
					if (title.startsWith("{")) {
						tabTitles.add(title);
					} else {
						tabTitles.add("{BUC3>6}" + title);
					}
				} else {
					tabTitles.add("{BUC3>6}" + name);
				}
			}
			for (String s : tabNames) {
				System.err.println("Tab: " + s);
			}
			for (String s : tabTitles) {
				System.err.println("Tab Title: " + s);
			}
		}

		JdbcTool jt = null;
		try {
			jt = new JdbcTool(url, user, password);
		} catch (Exception sqle) {
			printError("Unable to connect to database: " + sqle);
			System.exit(-2);
		}

		try {
			jt.start();
		} catch (Exception e) {
			jt.printError(e.getMessage());
			e.printStackTrace(System.err);
			System.exit(-3);
		}

		try {
			jt.close();
		} catch (Exception sqle) {
			printError("Error closing connection: " + sqle);
			System.exit(-4);
		}

	}

}
