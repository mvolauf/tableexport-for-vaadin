package com.vaadin.addon.tableexport;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.function.Consumer;
import java.util.logging.Logger;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.ss.util.WorkbookUtil;

import com.vaadin.ui.Grid;

/**
 * The Class ExcelExport. Implementation of TableExport to export Vaadin Tables to Excel .xls files.
 *
 * @author jnash
 * @version $Revision: 1.2 $
 */
public class ExcelExport extends TableExport {

    private static final long serialVersionUID = -8404407996727936497L;

    private static Logger LOGGER = Logger.getLogger(ExcelExport.class.getName());

    /**
     * The name of the sheet in the workbook the table contents will be written to.
     */
    protected String sheetName;

    /**
     * The title of the "report" of the table contents.
     */
    protected String reportTitle;

    /**
     * The filename of the workbook that will be sent to the user.
     */
    protected String exportFileName;

    /**
     * Flag indicating whether we will add a totals row to the Table. A totals row in the Table is
     * typically implemented as a footer and therefore is not part of the data source.
     */
    protected boolean displayTotals;

    /**
     * Flag indicating whether the first column should be treated as row headers. They will then be
     * formatted either like the column headers or a special row headers CellStyle can be specified.
     */
    protected boolean rowHeaders = false;

    /**
     * The workbook that contains the sheet containing the report with the table contents.
     */
    protected Workbook workbook;

    /**
     * The Sheet object that will contain the table contents report.
     */
    protected Sheet sheet;
    protected Sheet hierarchicalTotalsSheet = null;

    /**
     * The POI cell creation helper.
     */
    protected CreationHelper createHelper;
    protected DataFormat dataFormat;

    /**
     * Various styles that are used in report generation. These can be set by the user if the
     * default style is not desired to be used.
     */
    protected CellStyle dateCellStyle, doubleCellStyle, integerCellStyle, totalsDoubleCellStyle,
            totalsIntegerCellStyle, columnHeaderCellStyle, titleCellStyle;
    protected Short dateDataFormat, doubleDataFormat, integerDataFormat;
    protected Map<Short, CellStyle> dataFormatCellStylesMap = new HashMap<Short, CellStyle>();

    /**
     * The default row header style is null and, if row headers are specified with
     * setRowHeaders(true), then the column headers style is used. setRowHeaderStyle() allows the
     * user to specify a different row header style.
     */
    protected CellStyle rowHeaderCellStyle = null;

    /**
     * The totals row.
     */
    protected Row titleRow, headerRow, totalsRow;
    protected Row hierarchicalTotalsRow;
    // This let's the user specify the data format of the column in case the formatting of the column
    // will not be properly identified by the class of the column. In this case, the specified format is
    // used.  However, all other cell stylings will be those of the
    protected Map<String, String> columnExcelFormatMap = new HashMap<>();

    /**
     * At minimum, we need a Grid to export. Everything else has default settings.
     *
     * @param grid the grid
     */
    public ExcelExport(Grid<?> grid) {
        this(new DefaultGridHolder<>(grid), null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param grid      the grid
     * @param sheetName the sheet name
     */
    public ExcelExport(Grid<?> grid, String sheetName) {
        this(new DefaultGridHolder<>(grid), sheetName, null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param grid         the grid
     * @param sheetName   the sheet name
     * @param reportTitle the report title
     */
    public ExcelExport(Grid<?> grid, String sheetName, String reportTitle) {
        this(new DefaultGridHolder<>(grid), sheetName, reportTitle, null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param grid           the grid
     * @param sheetName      the sheet name
     * @param reportTitle    the report title
     * @param exportFileName the export file name
     */
    public ExcelExport(Grid<?> grid, String sheetName, String reportTitle,
                       String exportFileName) {
        this(new DefaultGridHolder<>(grid), sheetName, reportTitle, exportFileName, true);
    }

    /**
     * Instantiates a new TableExport class. This is the constructor that all other
     * constructors end up calling. If the other constructors were called then they pass in the
     * default parameters.
     *
     * @param grid           the grid
     * @param sheetName      the sheet name
     * @param reportTitle    the report title
     * @param exportFileName the export file name
     * @param hasTotalsRow   flag indicating whether we should create a totals row
     */
    public ExcelExport(Grid<?> grid, String sheetName, String reportTitle,
                       String exportFileName, boolean hasTotalsRow) {
        this(new DefaultGridHolder<>(grid), new HSSFWorkbook(), sheetName, reportTitle, exportFileName, hasTotalsRow);
    }

    public ExcelExport(Grid<?> grid, Workbook wkbk, String shtName, String rptTitle,
                       String xptFileName, boolean hasTotalsRow) {
        this(new DefaultGridHolder<>(grid), wkbk, shtName, rptTitle, xptFileName, hasTotalsRow);
    }

    /**
     * At minimum, we need a TableHolder to export. Everything else has default settings.
     *
     * @param tableHolder the tableHolder
     */
    public ExcelExport(TableHolder<?> tableHolder) {
        this(tableHolder, null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param tableHolder the tableHolder
     * @param sheetName   the sheet name
     */
    public ExcelExport(TableHolder<?> tableHolder, String sheetName) {
        this(tableHolder, sheetName, null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param tableHolder the tableHolder
     * @param sheetName   the sheet name
     * @param reportTitle the report title
     */
    public ExcelExport(TableHolder<?> tableHolder, String sheetName, String reportTitle) {
        this(tableHolder, sheetName, reportTitle, null);
    }

    /**
     * Instantiates a new TableExport class.
     *
     * @param tableHolder    the tableHolder
     * @param sheetName      the sheet name
     * @param reportTitle    the report title
     * @param exportFileName the export file name
     */
    public ExcelExport(TableHolder<?> tableHolder, String sheetName, String reportTitle,
                       String exportFileName) {
        this(tableHolder, sheetName, reportTitle, exportFileName, true);
    }

    /**
     * Instantiates a new TableExport class. This is the constructor that all other
     * constructors end up calling. If the other constructors were called then they pass in the
     * default parameters.
     *
     * @param tableHolder    the tableHolder
     * @param sheetName      the sheet name
     * @param reportTitle    the report title
     * @param exportFileName the export file name
     * @param hasTotalsRow   flag indicating whether we should create a totals row
     */
    public ExcelExport(TableHolder<?> tableHolder, String sheetName, String reportTitle,
                       String exportFileName, boolean hasTotalsRow) {
        this(tableHolder, new HSSFWorkbook(), sheetName, reportTitle, exportFileName, hasTotalsRow);
    }

    public ExcelExport(TableHolder<?> tableHolder, Workbook wkbk, String shtName,
                       String rptTitle, String xptFileName, boolean hasTotalsRow) {
        super(tableHolder);
        this.workbook = wkbk;
        init(shtName, rptTitle, xptFileName, hasTotalsRow);
    }

    private void init(String shtName, String rptTitle, String xptFileName,
                      boolean hasTotalsRow) {
        if ((null == shtName) || ("".equals(shtName))) {
            this.sheetName = "Table Export";
        } else {
            this.sheetName = shtName;
        }
        if (null == rptTitle) {
            this.reportTitle = "";
        } else {
            this.reportTitle = rptTitle;
        }
        if ((null == xptFileName) || ("".equals(xptFileName))) {
            this.exportFileName = "Table-Export.xls";
        } else {
            this.exportFileName = xptFileName;
        }
        this.displayTotals = hasTotalsRow;

        this.sheet = this.workbook.createSheet(this.sheetName);
        this.createHelper = this.workbook.getCreationHelper();
        this.dataFormat = this.workbook.createDataFormat();
        this.dateDataFormat = defaultDateDataFormat(this.workbook);
        this.doubleDataFormat = defaultDoubleDataFormat(this.workbook);
        this.integerDataFormat = defaultIntegerDataFormat(this.workbook);

        this.doubleCellStyle = defaultDataCellStyle(this.workbook);
        this.doubleCellStyle.setDataFormat(doubleDataFormat);
        this.dataFormatCellStylesMap.put(doubleDataFormat, doubleCellStyle);

        this.integerCellStyle = defaultDataCellStyle(this.workbook);
        this.integerCellStyle.setDataFormat(integerDataFormat);
        this.dataFormatCellStylesMap.put(integerDataFormat, integerCellStyle);

        this.dateCellStyle = defaultDataCellStyle(this.workbook);
        this.dateCellStyle.setDataFormat(this.dateDataFormat);
        this.dataFormatCellStylesMap.put(this.dateDataFormat, this.dateCellStyle);

        this.totalsDoubleCellStyle = defaultTotalsDoubleCellStyle(this.workbook);
        this.totalsIntegerCellStyle = defaultTotalsIntegerCellStyle(this.workbook);
        this.columnHeaderCellStyle = defaultHeaderCellStyle(this.workbook);
        this.titleCellStyle = defaultTitleCellStyle(this.workbook);
    }

    public void setNextTableHolder(TableHolder<?> tableHolder, String sheetName) {
        setTableHolder(tableHolder);
        sheet = workbook.createSheet(WorkbookUtil.createSafeSheetName(sheetName));
    }

    /*
     * This will exclude columns from the export that are not visible due to them being collapsed.
     * This should be called before convertTable() is called.
     */
    public void excludeCollapsedColumns() {
        Iterator<String> iterator = getColumnIds().iterator();
        while (iterator.hasNext()) {
            String columnId = iterator.next();
            if (getTableHolder().isColumnCollapsed(columnId)) {
                iterator.remove();
            }
        }
    }

    /**
     * Creates the workbook containing the exported table data, without exporting it to the user.
     */
    @Override
    public void convertTable() {
        int startRow;
        // initial setup
        initialSheetSetup();

        // add title row
        startRow = addTitleRow();
        int row = startRow;

        // add header row
        addHeaderRow(row);
        row++;

        // add data rows
        if (isHierarchical()) {
            row = addHierarchicalDataRows(sheet, row);
        } else {
            row = addDataRows(sheet, row);
        }

        // add totals row
        if (displayTotals) {
            addTotalsRow(row, startRow);
        }

        // sheet format before export
        finalSheetFormat();
    }

    @Override
    public File writeToTempFile() {
        if (null == mimeType) {
            setMimeType(XLS_MIME_TYPE);
        }
        File tempFile = null;
        try {
          tempFile = File.createTempFile("tmp", ".xls");
          tempFile.deleteOnExit();
        } catch (IOException e) {
          LOGGER.warning("Failed to create temp file " + e);
          return null;
        }

        try (FileOutputStream fos = new FileOutputStream(tempFile)) {
          workbook.write(fos);
        } catch (IOException e) {
          LOGGER.warning("Converting to XLS failed with IOException " + e);
          return null;
        }
        return tempFile;
    }

    /**
     * Create a download resource for the export
     * 
     * @return download resource
     */
    public TemporaryFileDownloadResource getDownloadResource() {
    	return new TemporaryFileDownloadResource(getExportFileName(), getMimeType(), this::exportToTempFile, null);
    }

    /**
    /**
     * Create a download resource for the export
     * 
  	 * @param onCloseCallback callback invoked after the stream was closed (just prior to file delete)
     * @return download resource
     */
    public TemporaryFileDownloadResource getDownloadResource(Consumer<File> onCloseCallback) {
    	return new TemporaryFileDownloadResource(getExportFileName(), getMimeType(), this::exportToTempFile, onCloseCallback);
    }
    
    /**
     * Initial sheet setup. Override this method to specifically change initial, sheet-wide,
     * settings.
     */
    protected void initialSheetSetup() {
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);
        if ((isHierarchical()) && (displayTotals)) {
            hierarchicalTotalsSheet = workbook.createSheet("tempHts");
        }
    }

    /**
     * Adds the title row. Override this method to change title-related aspects of the workbook.
     * Alternately, the title Row Object is accessible via getTitleRow() after report creation. To
     * change title text use setReportTitle(). To change title CellStyle use setTitleStyle().
     *
     * @return the int
     */
    protected int addTitleRow() {
        if ((null == reportTitle) || ("".equals(reportTitle))) {
            return 0;
        }
        titleRow = sheet.createRow(0);
        titleRow.setHeightInPoints(45);
        Cell titleCell;
        CellRangeAddress cra;
        if (rowHeaders) {
            titleCell = titleRow.createCell(1);
            cra = new CellRangeAddress(0, 0, 1, getColumnIds().size() - 1);
            sheet.addMergedRegion(cra);
        } else {
            titleCell = titleRow.createCell(0);
            cra = new CellRangeAddress(0, 0, 0, getColumnIds().size() - 1);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, getColumnIds().size() - 1));
        }
        titleCell.setCellValue(reportTitle);
        titleCell.setCellStyle(titleCellStyle);
        // cell borders don't work on merged ranges so, if there are borders
        // we apply them to the merged range here.
        if (titleCellStyle.getBorderLeft() != BorderStyle.NONE) {
            RegionUtil.setBorderLeft(titleCellStyle.getBorderLeft(), cra, sheet);
        }
        if (titleCellStyle.getBorderRight() != BorderStyle.NONE) {
            RegionUtil.setBorderRight(titleCellStyle.getBorderRight(), cra, sheet);
        }
        if (titleCellStyle.getBorderTop() != BorderStyle.NONE) {
            RegionUtil.setBorderTop(titleCellStyle.getBorderTop(), cra, sheet);
        }
        if (titleCellStyle.getBorderBottom() != BorderStyle.NONE) {
            RegionUtil.setBorderBottom(titleCellStyle.getBorderBottom(), cra, sheet);
        }
        return 1;
    }

    /**
     * Adds the header row. Override this method to change header-row-related aspects of the
     * workbook. Alternately, the header Row Object is accessible via getHeaderRow() after report
     * creation. To change header CellStyle, though, use setHeaderStyle().
     *
     * @param row the row
     */
    protected void addHeaderRow(int row) {
        headerRow = sheet.createRow(row);
        Cell headerCell;
        String columnId;
        headerRow.setHeightInPoints(40);
        for (int col = 0; col < getColumnIds().size(); col++) {
            columnId = getColumnIds().get(col);
            headerCell = headerRow.createCell(col);
            headerCell.setCellValue(createHelper.createRichTextString(getTableHolder().getColumnHeader(columnId)
                    .toString()));
            headerCell.setCellStyle(getColumnHeaderStyle(row, col));

            Short poiAlignment = getTableHolder().getCellAlignment(columnId);
            CellUtil.setAlignment(headerCell, HorizontalAlignment.forInt(poiAlignment));
        }
    }

    /**
     * This method is called by addTotalsRow() to determine what CellStyle to use. By default we
     * just return totalsCellStyle which is either set to the default totals style, or can be
     * overriden by the user using setTotalsStyle(). However, if the user wants to have different
     * total items have different styles, then this method should be overriden. The parameters
     * passed in are all potentially relevant items that may be used to determine what formatting to
     * return, that are not accessible globally.
     *
     * @param row the row
     * @param col the current column
     * @return the header style
     */
    protected CellStyle getColumnHeaderStyle(int row, int col) {
        if ((rowHeaders) && (col == 0)) {
            return titleCellStyle;
        }
        return columnHeaderCellStyle;
    }

    /**
     * For Hierarchical Containers, this method recursively adds root items and child items. The
     * child items are appropriately grouped using grouping/outlining sheet functionality. Override
     * this method to make any changes. To change the CellStyle used for all Table data use
     * setDataStyle(). For different data cells to have different CellStyles, override
     * getDataStyle().
     *
     * @param row the row
     * @return the int
     */
    protected int addHierarchicalDataRows(Sheet sheetToAddTo, int row) {
        Collection<?> roots;
        int localRow = row;
        roots = getTableHolder().getRootItems();
        /*
         * For Hierarchical Containers, the outlining/grouping in the sheet is with the summary row
         * at the top and the grouped/outlined subcategories below.
         */
        sheet.setRowSumsBelow(false);
        int count = 0;
        for (Object rootId : roots) {
            count = addDataRowRecursively(sheetToAddTo, rootId, localRow);
            // for totals purposes, we just want to add rootIds which contain totals
            // so we store just the totals in a separate sheet.
            if (displayTotals) {
                addDataRow(hierarchicalTotalsSheet, rootId, localRow);
            }
            if (count > 1) {
                sheet.groupRow(localRow + 1, (localRow + count) - 1);
                if (collapseRowGroup(rootId)) {
                	sheet.setRowGroupCollapsed(localRow + 1, true);
                }
            }
            localRow = localRow + count;
        }
        return localRow;
    }

    /**
     * Determines if a group rooted in object {@code rootId} should be collapsed or not.
     * By default, all rows of the hierarchical containers are not collapsed.
     *
     * @param rootId
     * @return
     */
    protected boolean collapseRowGroup(Object rootId) {
      return true;
    }

    /**
     * this method adds row items for non-Hierarchical Containers. Override this method to make any
     * changes. To change the CellStyle used for all Table data use setDataStyle(). For different
     * data cells to have different CellStyles, override getDataStyle().
     *
     * @param row the row
     * @return the int
     */
    protected int addDataRows(Sheet sheetToAddTo, int row) {
        Collection<?> items = getTableHolder().getItems();
        int localRow = row;
        for (Object item : items) {
            addDataRow(sheetToAddTo, item, localRow);
            localRow++;
        }
        return localRow;
    }

    /**
     * Used by addHierarchicalDataRows() to implement the recursive calls.
     *
     * @param rootItem the root item
     * @param row        the row
     * @return the int
     */
    protected <X> int addDataRowRecursively(Sheet sheetToAddTo, X rootItem, int row) {
        int numberAdded = 0;
        int localRow = row;
        addDataRow(sheetToAddTo, rootItem, row);
        numberAdded++;
        for (X child : ((TableHolder<X>) getTableHolder()).getChildren(rootItem)) {
            localRow++;
            numberAdded = numberAdded + addDataRowRecursively(sheetToAddTo, child, localRow);
        }
        return numberAdded;
    }

    /**
     * This method is ultimately used by either addDataRows() or addHierarchicalDataRows() to
     * actually add the data to the Sheet.
     *
     * @param rootItem the root item id
     * @param row        the row
     */
    protected <X> void addDataRow(Sheet sheetToAddTo, X rootItem, int row) {
        Row sheetRow = sheetToAddTo.createRow(row);
        String columnId;
        Object value;
        Class<?> valueType;
        Cell sheetCell;
    	TableHolder<X> holder = (TableHolder<X>) getTableHolder();
        for (int col = 0; col < getColumnIds().size(); col++) {
            columnId = getColumnIds().get(col);
            value = holder.getColumnValue(rootItem, columnId);
            valueType = holder.getColumnType(columnId);
            sheetCell = sheetRow.createCell(col);
            setupCell(sheetCell, value, valueType, columnId, row, col);
        }
    }

    protected void setupCell(Cell sheetCell, Object value, Class<?> valueType, String columnId, int row, int col) {
        sheetCell.setCellStyle(getCellStyle(columnId, row, col, false));
        Short poiAlignment = getTableHolder().getCellAlignment(columnId);
        CellUtil.setAlignment(sheetCell, HorizontalAlignment.forInt(poiAlignment));
        setCellValue(sheetCell, value, valueType, columnId);
    }
    
	protected void setCellValue(Cell sheetCell, Object value, Class<?> valueType, String columnId) {
		if (null != value) {
		    if (!isNumeric(valueType)) {
		        if (java.util.Date.class.isAssignableFrom(valueType)) {
		            sheetCell.setCellValue((Date) value);
		        } else {
		            sheetCell.setCellValue(createHelper.createRichTextString(value.toString()));
		        }
		    } else {
		        try {
		            // parse all numbers as double, the format will determine how they appear
		            Double d = Double.parseDouble(value.toString());
		            sheetCell.setCellValue(d);
		        } catch (NumberFormatException nfe) {
		            LOGGER.warning("NumberFormatException parsing a numeric value: " + nfe);
		            sheetCell.setCellValue(createHelper.createRichTextString(value.toString()));
		        }
		    }
		}
	}

    public void setExcelFormatOfColumn(String columnId, String excelFormat) {
        if (this.columnExcelFormatMap.containsKey(columnId)) {
            this.columnExcelFormatMap.remove(columnId);
        }
        this.columnExcelFormatMap.put(columnId, excelFormat);
    }

    /**
     * This method is called by addDataRow() to determine what CellStyle to use. By default we just
     * return dataStyle which is either set to the default data style, or can be overriden by the
     * user using setDataStyle(). However, if the user wants to have different data items have
     * different styles, then this method should be overriden. The parameters passed in are all
     * potentially relevant items that may be used to determine what formatting to return, that are
     * not accessible globally.
     *
     * @param columnId     the column id
     * @param row        the row
     * @param col        the col
     * @return the data style
     */
    protected CellStyle getCellStyle(String columnId, int row, int col, boolean totalsRow) {
        // get the basic style for the type of cell (i.e. data, header, total)
        if ((rowHeaders) && (col == 0)) {
            if (null == rowHeaderCellStyle) {
                return columnHeaderCellStyle;
            }
            return rowHeaderCellStyle;
        }
        Class<?> columnType = getTableHolder().getColumnType(columnId);
        if (totalsRow) {
            if (this.columnExcelFormatMap.containsKey(columnId)) {
                short df = dataFormat.getFormat(columnExcelFormatMap.get(columnId));
                CellStyle customTotalStyle = workbook.createCellStyle();
                customTotalStyle.cloneStyleFrom(totalsDoubleCellStyle);
                customTotalStyle.setDataFormat(df);
                return customTotalStyle;
            }
            if (isIntegerLongShortOrBigDecimal(columnType)) {
                return totalsIntegerCellStyle;
            }
            return totalsDoubleCellStyle;
        }
        // Check if the user has over-ridden that data format of this property
        if (this.columnExcelFormatMap.containsKey(columnId)) {
            short df = dataFormat.getFormat(columnExcelFormatMap.get(columnId));
            if (dataFormatCellStylesMap.containsKey(df)) {
                return dataFormatCellStylesMap.get(df);
            }
            // if it hasn't already been created for re-use, we create a cell style and override the data format
            // For data cells, each data format corresponds to a single complete cell style
            CellStyle retStyle = workbook.createCellStyle();
            retStyle.cloneStyleFrom(dataFormatCellStylesMap.get(doubleDataFormat));
            retStyle.setDataFormat(df);
            dataFormatCellStylesMap.put(df, retStyle);
            return retStyle;
        }
        // if not over-ridden, use the overall setting
        if (isDoubleOrFloat(columnType)) {
            return dataFormatCellStylesMap.get(doubleDataFormat);
        } else if (isIntegerLongShortOrBigDecimal(columnType)) {
            return dataFormatCellStylesMap.get(integerDataFormat);
        } else if (java.util.Date.class.isAssignableFrom(columnType)) {
            return dataFormatCellStylesMap.get(dateDataFormat);
        }
        return dataFormatCellStylesMap.get(doubleDataFormat);
    }

    /**
     * Adds the totals row to the report. Override this method to make any changes. Alternately, the
     * totals Row Object is accessible via getTotalsRow() after report creation. To change the
     * CellStyle used for the totals row, use setFormulaStyle. For different totals cells to have
     * different CellStyles, override getTotalsStyle().
     *
     * @param currentRow the current row
     * @param startRow   the start row
     */
    protected void addTotalsRow(int currentRow, int startRow) {
        totalsRow = sheet.createRow(currentRow);
        totalsRow.setHeightInPoints(30);
        Cell cell;
        for (int col = 0; col < getColumnIds().size(); col++) {
            String columnId = getColumnIds().get(col);
            cell = totalsRow.createCell(col);
            setupTotalCell(cell, columnId, currentRow, startRow, col);
        }
    }

	protected void setupTotalCell(Cell cell, String columnId, int currentRow, int startRow, int col) {
		cell.setCellStyle(getCellStyle(columnId, startRow, col, true));
		Short poiAlignment = getTableHolder().getCellAlignment(columnId);
		CellUtil.setAlignment(cell, HorizontalAlignment.forInt(poiAlignment));
		Class<?> columnType = getTableHolder().getColumnType(columnId);
		if (isNumeric(columnType)) {
			CellRangeAddress cra = new CellRangeAddress(startRow, currentRow - 1, col, col);
		    if (isHierarchical()) {
		        // 9 & 109 are for sum. 9 means include hidden cells, 109 means exclude.
		        // this will show the wrong value if the user expands an outlined category, so
		        // we will range value it first
		        cell.setCellFormula("SUM(" + cra.formatAsString(hierarchicalTotalsSheet.getSheetName(),
		                true) + ")");
		    } else {
		        cell.setCellFormula("SUM(" + cra.formatAsString() + ")");
		    }
		} else {
		    if (0 == col) {
		        cell.setCellValue(createHelper.createRichTextString(getTotalHeader()));
		    }
		}
	}

	protected String getTotalHeader() {
		return "Total";
	}
	
    /**
     * formatting of the sheet upon completion of writing the data. For example, we can only
     * size the column widths once the data is in the report and the sheet knows how wide the data
     * is.
     */
    protected void finalSheetFormat() {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        if (isHierarchical()) {
            /*
             * evaluateInCell() is equivalent to paste special -> value. The formula refers to cells
             * in the other sheet we are going to delete. We sum in the other sheet because if we
             * summed in the main sheet, we would double count. Subtotal with hidden rows is not yet
             * implemented in POI.
             */
            for (Row r : sheet) {
                for (Cell c : r) {
                    if (c.getCellType() == CellType.FORMULA) {
                        evaluator.evaluateInCell(c);
                    }
                }
            }
            workbook.setActiveSheet(workbook.getSheetIndex(sheet));
            if (hierarchicalTotalsSheet != null) {
                workbook.removeSheetAt(workbook.getSheetIndex(hierarchicalTotalsSheet));
            }
        } else {
            evaluator.evaluateAll();
        }
        for (int col = 0; col < getColumnIds().size(); col++) {
            sheet.autoSizeColumn(col);
        }
    }

    /**
     * Returns the default title style. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultTitleCellStyle(Workbook wb) {
        CellStyle style;
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short) 18);
        titleFont.setBold(true);
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(titleFont);
        return style;
    }

    /**
     * Returns the default header style. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultHeaderCellStyle(Workbook wb) {
        CellStyle style;
        Font monthFont = wb.createFont();
        monthFont.setFontHeightInPoints((short) 11);
        monthFont.setColor(IndexedColors.WHITE.getIndex());
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFont(monthFont);
        style.setWrapText(true);
        return style;
    }

    /**
     * Returns the default data cell style. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultDataCellStyle(Workbook wb) {
        CellStyle style;
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(true);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setDataFormat(doubleDataFormat);
        return style;
    }

    /**
     * Returns the default totals row style for Double data. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultTotalsDoubleCellStyle(Workbook wb) {
        CellStyle style;
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(doubleDataFormat);
        return style;
    }

    /**
     * Returns the default totals row style for Integer data. Obtained from: http://svn.apache.org/repos/asf/poi
     * /trunk/src/examples/src/org/apache/poi/ss/examples/TimesheetDemo.java
     *
     * @param wb the wb
     * @return the cell style
     */
    protected CellStyle defaultTotalsIntegerCellStyle(Workbook wb) {
        CellStyle style;
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setDataFormat(integerDataFormat);
        return style;
    }

    protected short defaultDoubleDataFormat(Workbook wb) {
        return createHelper.createDataFormat().getFormat("0.00");
    }

    protected short defaultIntegerDataFormat(Workbook wb) {
        return createHelper.createDataFormat().getFormat("0");
    }

    protected short defaultDateDataFormat(Workbook wb) {
        return createHelper.createDataFormat().getFormat("mm/dd/yyyy");
    }

    public void setDoubleDataFormat(String excelDoubleFormat) {
        CellStyle prevDoubleDataStyle = null;
        if (dataFormatCellStylesMap.containsKey(doubleDataFormat)) {
            prevDoubleDataStyle = dataFormatCellStylesMap.get(doubleDataFormat);
            dataFormatCellStylesMap.remove(doubleDataFormat);
        }
        doubleDataFormat = createHelper.createDataFormat().getFormat(excelDoubleFormat);
        if (null != prevDoubleDataStyle) {
            doubleCellStyle = prevDoubleDataStyle;
            doubleCellStyle.setDataFormat(doubleDataFormat);
            dataFormatCellStylesMap.put(doubleDataFormat, doubleCellStyle);
        }
    }

    public void setIntegerDataFormat(String excelIntegerFormat) {
        CellStyle prevIntegerDataStyle = null;
        if (dataFormatCellStylesMap.containsKey(integerDataFormat)) {
            prevIntegerDataStyle = dataFormatCellStylesMap.get(integerDataFormat);
            dataFormatCellStylesMap.remove(integerDataFormat);
        }
        integerDataFormat = createHelper.createDataFormat().getFormat(excelIntegerFormat);
        if (null != prevIntegerDataStyle) {
            integerCellStyle = prevIntegerDataStyle;
            integerCellStyle.setDataFormat(integerDataFormat);
            dataFormatCellStylesMap.put(integerDataFormat, integerCellStyle);
        }
    }

    public void setDateDataFormat(String excelDateFormat) {
        CellStyle prevDateDataStyle = null;
        if (dataFormatCellStylesMap.containsKey(dateDataFormat)) {
            prevDateDataStyle = dataFormatCellStylesMap.get(dateDataFormat);
            dataFormatCellStylesMap.remove(dateDataFormat);
        }
        dateDataFormat = createHelper.createDataFormat().getFormat(excelDateFormat);
        if (null != prevDateDataStyle) {
            dateCellStyle = prevDateDataStyle;
            dateCellStyle.setDataFormat(dateDataFormat);
            dataFormatCellStylesMap.put(dateDataFormat, dateCellStyle);
        }
    }

    /**
     * Utility method to determine whether value being put in the Cell is numeric.
     *
     * @param type the type
     * @return true, if is numeric
     */
    public static boolean isNumeric(Class<?> type) {
        if (isIntegerLongShortOrBigDecimal(type)) {
            return true;
        }
        if (isDoubleOrFloat(type)) {
            return true;
        }
        if (Number.class.equals(type)) {
            return true;
        }
        return false;
    }

    /**
     * Utility method to determine whether value being put in the Cell is integer-like type.
     *
     * @param type the type
     * @return true, if is integer-like
     */
    public static boolean isIntegerLongShortOrBigDecimal(Class<?> type) {
        if ((Integer.class.equals(type) || (int.class.equals(type)))) {
            return true;
        }
        if ((Long.class.equals(type) || (long.class.equals(type)))) {
            return true;
        }
        if ((Short.class.equals(type)) || (short.class.equals(type))) {
            return true;
        }
        if ((BigDecimal.class.equals(type)) || (BigDecimal.class.equals(type))) {
            return true;
        }
        return false;
    }

    /**
     * Utility method to determine whether value being put in the Cell is double-like type.
     *
     * @param type the type
     * @return true, if is double-like
     */
    public static boolean isDoubleOrFloat(Class<?> type) {
        if ((Double.class.equals(type)) || (double.class.equals(type))) {
            return true;
        }
        if ((Float.class.equals(type)) || (float.class.equals(type))) {
            return true;
        }
        return false;
    }

    /**
     * Gets the workbook.
     *
     * @return the workbook
     */
    public Workbook getWorkbook() {
        return this.workbook;
    }

    /**
     * Gets the sheet name.
     *
     * @return the sheet name
     */
    public String getSheetName() {
        return this.sheetName;
    }

    /**
     * Gets the report title.
     *
     * @return the report title
     */
    public String getReportTitle() {
        return this.reportTitle;
    }

    /**
     * Gets the export file name.
     *
     * @return the export file name
     */
    public String getExportFileName() {
        return this.exportFileName;
    }

    /**
     * Gets the cell style used for report data..
     *
     * @return the cell style
     */
    public CellStyle getDoubleDataStyle() {
        return this.doubleCellStyle;
    }

    /**
     * Gets the cell style used for report data..
     *
     * @return the cell style
     */
    public CellStyle getIntegerDataStyle() {
        return this.integerCellStyle;
    }

    public CellStyle getDateDataStyle() {
        return this.dateCellStyle;
    }

    /**
     * Gets the cell style used for the report headers.
     *
     * @return the column header style
     */
    public CellStyle getColumnHeaderStyle() {
        return this.columnHeaderCellStyle;
    }

    /**
     * Gets the cell title used for the report title.
     *
     * @return the title style
     */
    public CellStyle getTitleStyle() {
        return this.titleCellStyle;
    }

    /**
     * Sets the text used for the report title.
     *
     * @param reportTitle the new report title
     */
    public void setReportTitle(String reportTitle) {
        this.reportTitle = reportTitle;
    }

    /**
     * Sets the export file name.
     *
     * @param exportFileName the new export file name
     */
    public void setExportFileName(String exportFileName) {
        this.exportFileName = exportFileName;
    }

    /**
     * Sets the cell style used for report data.
     *
     * @param doubleDataStyle the new data style
     */
    public void setDoubleDataStyle(CellStyle doubleDataStyle) {
        this.doubleCellStyle = doubleDataStyle;
    }

    /**
     * Sets the cell style used for report data.
     *
     * @param integerDataStyle the new data style
     */
    public void setIntegerDataStyle(CellStyle integerDataStyle) {
        this.integerCellStyle = integerDataStyle;
    }

    /**
     * Sets the cell style used for report data.
     *
     * @param dateDataStyle the new data style
     */
    public void setDateDataStyle(CellStyle dateDataStyle) {
        this.dateCellStyle = dateDataStyle;
    }

    /**
     * Sets the cell style used for the report headers.
     *
     * @param columnHeaderStyle CellStyle
     */
    public void setColumnHeaderStyle(CellStyle columnHeaderStyle) {
        this.columnHeaderCellStyle = columnHeaderStyle;
    }

    /**
     * Sets the cell style used for the report title.
     *
     * @param titleStyle the new title style
     */
    public void setTitleStyle(CellStyle titleStyle) {
        this.titleCellStyle = titleStyle;
    }

    /**
     * Gets the title row.
     *
     * @return the title row
     */
    public Row getTitleRow() {
        return this.titleRow;
    }

    /**
     * Gets the header row.
     *
     * @return the header row
     */
    public Row getHeaderRow() {
        return this.headerRow;
    }

    /**
     * Gets the totals row.
     *
     * @return the totals row
     */
    public Row getTotalsRow() {
        return this.totalsRow;
    }

    /**
     * Gets the cell style used for the totals row.
     *
     * @return the totals style
     */
    public CellStyle getTotalsDoubleStyle() {
        return this.totalsDoubleCellStyle;
    }

    /**
     * Sets the cell style used for the totals row.
     *
     * @param totalsDoubleStyle the new totals style
     */
    public void setTotalsDoubleStyle(CellStyle totalsDoubleStyle) {
        this.totalsDoubleCellStyle = totalsDoubleStyle;
    }

    /**
     * Gets the cell style used for the totals row.
     *
     * @return the totals style
     */
    public CellStyle getTotalsIntegerStyle() {
        return this.totalsIntegerCellStyle;
    }

    /**
     * Sets the cell style used for the totals row.
     *
     * @param totalsIntegerStyle the new totals style
     */
    public void setTotalsIntegerStyle(CellStyle totalsIntegerStyle) {
        this.totalsIntegerCellStyle = totalsIntegerStyle;
    }

    /**
     * Flag indicating whether a totals row will be added to the report or not.
     *
     * @return true, if totals row will be added
     */
    public boolean isDisplayTotals() {
        return this.displayTotals;
    }

    /**
     * Sets the flag indicating whether a totals row will be added to the report or not.
     *
     * @param displayTotals boolean
     */
    public void setDisplayTotals(boolean displayTotals) {
        this.displayTotals = displayTotals;
    }

    /**
     * See value of flag indicating whether the first column should be treated as row headers.
     *
     * @return boolean
     */
    public boolean hasRowHeaders() {
        return this.rowHeaders;
    }

    /**
     * Method getRowHeaderStyle.
     *
     * @return CellStyle
     */
    public CellStyle getRowHeaderStyle() {
        return this.rowHeaderCellStyle;
    }

    /**
     * Set value of flag indicating whether the first column should be treated as row headers.
     *
     * @param rowHeaders boolean
     */
    public void setRowHeaders(boolean rowHeaders) {
        this.rowHeaders = rowHeaders;
    }

    /**
     * Method setRowHeaderStyle.
     *
     * @param rowHeaderStyle CellStyle
     */
    public void setRowHeaderStyle(CellStyle rowHeaderStyle) {
        this.rowHeaderCellStyle = rowHeaderStyle;
    }

}
