package com.vaadin.addon.tableexport;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.logging.Logger;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.vaadin.ui.Grid;

public class CsvExport extends ExcelExport {
	private static final long serialVersionUID = 935966816321924835L;
	private static Logger LOGGER = Logger.getLogger(CsvExport.class.getName());

	public CsvExport(Grid<?> grid) {
		super(grid);
	}

	public CsvExport(Grid<?> grid, String sheetName) {
		super(grid, sheetName);
	}

	public CsvExport(Grid<?> grid, String sheetName, String reportTitle) {
		super(grid, sheetName, reportTitle);
	}

	public CsvExport(Grid<?> grid, String sheetName, String reportTitle, String exportFileName) {
		super(grid, sheetName, reportTitle, exportFileName);
	}

	public CsvExport(Grid<?> grid, String sheetName, String reportTitle, String exportFileName, boolean hasTotalsRow) {
		super(grid, sheetName, reportTitle, exportFileName, hasTotalsRow);
	}

	public CsvExport(TableHolder<?> tableHolder) {
		super(tableHolder);
	}

	public CsvExport(TableHolder<?> tableHolder, String sheetName) {
		super(tableHolder, sheetName);
	}

	public CsvExport(TableHolder<?> tableHolder, String sheetName, String reportTitle) {
		super(tableHolder, sheetName, reportTitle);
	}

	public CsvExport(TableHolder<?> tableHolder, String sheetName, String reportTitle, String exportFileName) {
		super(tableHolder, sheetName, reportTitle, exportFileName);
	}

	public CsvExport(TableHolder<?> tableHolder, String sheetName, String reportTitle, String exportFileName,
			boolean hasTotalsRow) {
		super(tableHolder, sheetName, reportTitle, exportFileName, hasTotalsRow);
	}

	public File writeToTempFile() {
		if (null == mimeType) {
			setMimeType(CSV_MIME_TYPE);
		}
		File tempXlsFile, tempCsvFile;
		try {
			tempXlsFile = File.createTempFile("tmp", ".xls");
			FileOutputStream fileOut = new FileOutputStream(tempXlsFile);
			workbook.write(fileOut);
			FileInputStream fis = new FileInputStream(tempXlsFile);
			POIFSFileSystem fs = new POIFSFileSystem(fis);
			tempCsvFile = File.createTempFile("tmp", ".csv");
			PrintStream p = new PrintStream(new BufferedOutputStream(new FileOutputStream(tempCsvFile, true)));

			XLS2CSVmra xls2csv = new XLS2CSVmra(fs, p, -1);
			xls2csv.process();
			p.close();
			return tempCsvFile;
		} catch (IOException e) {
			LOGGER.warning("Converting to CSV failed with IOException " + e);
			return null;
		}
	}
}
