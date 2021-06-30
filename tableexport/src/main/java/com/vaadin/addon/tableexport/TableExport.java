package com.vaadin.addon.tableexport;

import java.io.File;
import java.io.Serializable;
import java.util.List;
import java.util.function.Consumer;
import java.util.logging.Logger;

public abstract class TableExport implements Serializable {

    private static final long serialVersionUID = -2972527330991334117L;
    private static final Logger LOGGER = Logger.getLogger(TableExport.class.getName());

    public static String EXCEL_MIME_TYPE = "application/vnd.ms-excel";
    public static String CSV_MIME_TYPE = "text/csv";

    /** The Tableholder to export. */
    private TableHolder tableHolder;

    /** The window to send the export result */
    protected String exportWindow = "_self";

    protected String mimeType;

    public TableExport(TableHolder tableHolder) {
        this.tableHolder = tableHolder;
    }

    public TableHolder getTableHolder() {
        return tableHolder;
    }

    public List<?> getPropIds() {
        return tableHolder.getPropIds();
    }

    public void setTableHolder(final TableHolder tableHolder) {
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
  	 * @param onCloseCallback (optional) callback invoked after the stream was closed (just prior to temp file delete)
     * @return stream source providing the exported data  
     */
    public TemporaryFileStreamSource createExportSource(Consumer<File> onCloseCallback) {
        return new TemporaryFileStreamSource(this::exportToTempFile, onCloseCallback);
    }

    public String getExportWindow() {
        return this.exportWindow;
    }

    public void setExportWindow(final String exportWindow) {
        this.exportWindow = exportWindow;
    }

    public String getMimeType() {
        return this.mimeType;
    }

    public void setMimeType(final String mimeType) {
        this.mimeType = mimeType;
    }

}
