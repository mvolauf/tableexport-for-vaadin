package com.vaadin.addon.tableexport.demo;

import com.vaadin.addon.tableexport.ExcelExport;
import com.vaadin.addon.tableexport.TableHolder;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.*;

/**
 * Example of how the ExcelExport class might be extended to implement specific formatting features
 * in the exported file.
 */
public class EnhancedFormatExcelExport extends ExcelExport {

    /**
     * The Constant serialVersionUID.
     */
    private static final long serialVersionUID = 9113961084041090666L;

    public EnhancedFormatExcelExport(final TableHolder tableHolder) {
        this(tableHolder, "Enhanced Export");
    }

    public EnhancedFormatExcelExport(final TableHolder tableHolder, final String sheetName) {
        super(tableHolder, sheetName);
        format();
    }

    private void format() {
        this.setRowHeaders(true);
        CellStyle style;
        Font f;

        style = this.getTitleStyle();
        setStyle(style, HSSFColorPredefined.DARK_BLUE.getIndex(), 18, HSSFColorPredefined.WHITE.getIndex(), true,
                HorizontalAlignment.CENTER_SELECTION);

        style = this.getColumnHeaderStyle();
        setStyle(style, HSSFColorPredefined.LIGHT_BLUE.getIndex(), 12, HSSFColorPredefined.BLACK.getIndex(), true,
                HorizontalAlignment.CENTER);

        style = this.getDateDataStyle();
        setStyle(style, HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE.getIndex(), 12, HSSFColorPredefined.BLACK.getIndex(), false,
                HorizontalAlignment.RIGHT);

        style = this.getDoubleDataStyle();
        setStyle(style, HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE.getIndex(), 12, HSSFColorPredefined.BLACK.getIndex(), false,
                HorizontalAlignment.RIGHT);
        this.setTotalsDoubleStyle(style);

        style = this.getIntegerDataStyle();
        setStyle(style, HSSFColorPredefined.LIGHT_CORNFLOWER_BLUE.getIndex(), 12, HSSFColorPredefined.BLACK.getIndex(), false,
                HorizontalAlignment.RIGHT);
        this.setTotalsIntegerStyle(style);

        // we want the rowHeader style to be like the columnHeader style, just centered differently.
        final CellStyle newStyle = workbook.createCellStyle();
        newStyle.cloneStyleFrom(style);
        newStyle.setAlignment(HorizontalAlignment.LEFT);
        this.setRowHeaderStyle(newStyle);
    }

    private void setStyle(CellStyle style, short foregroundColor, int fontHeight, short fontColor,
                          boolean isBold, HorizontalAlignment alignment) {
        style.setFillForegroundColor(foregroundColor);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font f = workbook.getFontAt(style.getFontIndex());
        f.setFontHeightInPoints((short) fontHeight);
        f.setFontName(HSSFFont.FONT_ARIAL);
        f.setColor(fontColor);
        f.setBold(isBold);
        style.setAlignment(alignment);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setRightBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setTopBorderColor(HSSFColorPredefined.BLACK.getIndex());
        style.setBottomBorderColor(HSSFColorPredefined.BLACK.getIndex());
    }

}