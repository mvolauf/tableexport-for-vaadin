package com.vaadin.addon.tableexport;

import java.io.File;
import java.io.Serializable;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.function.Consumer;

public abstract class TableExport implements Serializable {

	private static final long serialVersionUID = -2972527330991334117L;

	public static String XLS_MIME_TYPE = "application/vnd.ms-excel";
	public static String XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
	public static String CSV_MIME_TYPE = "text/csv";

	/** The Tableholder to export. */
	private TableHolder<?> tableHolder;

	protected Set<Object> excludedColumns = new HashSet<>();

	protected String mimeType;

	public TableExport(TableHolder<?> tableHolder) {
		this.tableHolder = tableHolder;
	}

	public TableHolder<?> getTableHolder() {
		return tableHolder;
	}

	public List<String> getColumnIds() {
		List<String> columnIds = tableHolder.getColumnIds();
		columnIds.removeAll(excludedColumns);
		return columnIds;
	}

	public void setTableHolder(TableHolder<?> tableHolder) {
		this.tableHolder = tableHolder;
	}

	public boolean isHierarchical() {
		return tableHolder.isHierarchical();
	}

	public abstract void convertTable();

	public abstract File writeToTempFile();

	/**
	 * Perform the export into temporary file
	 * 
	 * @return temporary file with exported data
	 */
	public File exportToTempFile() {
		convertTable();
		return writeToTempFile();
	}

	/**
	 * Create the export source
	 * 
	 * @param onCloseCallback (optional) callback invoked after the stream was
	 *                        closed (just prior to temp file delete)
	 * @return stream source providing the exported data
	 */
	public TemporaryFileStreamSource createExportSource(Consumer<File> onCloseCallback) {
		return new TemporaryFileStreamSource(this::exportToTempFile, onCloseCallback);
	}

	public String getMimeType() {
		return this.mimeType;
	}

	public void setMimeType(String mimeType) {
		this.mimeType = mimeType;
	}

	/**
	 * Removes a column from the exported table.
	 *
	 * If some of the displayed columns should not be exported, they can be removed
	 * with calls to this method.
	 *
	 * @param excludedColumns The columns to exclude
	 */
	public void setExcludedColumns(Set<Object> excludedColumns) {
		this.excludedColumns = excludedColumns;
	}

}
