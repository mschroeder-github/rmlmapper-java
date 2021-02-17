package be.ugent.rml.records;

import be.ugent.rml.NAMESPACES;
import be.ugent.rml.Utils;
import be.ugent.rml.access.Access;
import be.ugent.rml.access.LocalFileAccess;
import be.ugent.rml.store.QuadStore;
import be.ugent.rml.term.NamedNode;
import be.ugent.rml.term.Term;
import java.awt.Color;
import java.awt.Dimension;
import java.io.IOException;
import java.io.InputStream;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import javax.script.Bindings;
import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.script.ScriptException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * This class is a record factory that creates Excel records.
 */
public class ExcelRecordFactory implements ReferenceFormulationRecordFactory {

    private ScriptEngineManager sem;
    private ScriptEngine jsEngine;

    private boolean useCache = true;
    private Map<String, Workbook> workbookCache;
    private Map<String, List<Record>> recordCache;

    public ExcelRecordFactory() {
        sem = new ScriptEngineManager();
        jsEngine = sem.getEngineByName("JavaScript");
        workbookCache = new HashMap<>();
        recordCache = new HashMap<>();
    }

    @Override
    public List<Record> getRecords(Access access, Term logicalSource, QuadStore rmlStore) throws IOException, SQLException, ClassNotFoundException {

        long begin = System.currentTimeMillis();
        
        //options
        boolean ignoreBlank = true;

        InputStream is = access.getInputStream();

        List<Term> sources = Utils.getObjectsFromQuads(rmlStore.getQuads(logicalSource, new NamedNode(NAMESPACES.RML + "source"), null));
        if(sources.isEmpty()) {
            throw new IOException("you have to define a rml:source which is a ss:Workbook.");
        }
        Term source = sources.get(0);
        
        List<String> sheetNames = Utils.getLiteralObjectsFromQuads(rmlStore.getQuads(source, new NamedNode(NAMESPACES.SS + "sheetName"), null));
        if (sheetNames.isEmpty()) {
            throw new IOException("you have to define a ss:sheetName to select a sheet from the excel file.");
        }
        String sheetName = sheetNames.get(0);

        List<String> ranges = Utils.getLiteralObjectsFromQuads(rmlStore.getQuads(source, new NamedNode(NAMESPACES.SS + "range"), null));
        CellRangeAddress cellRangeAddress = null;
        if (!ranges.isEmpty()) {
            cellRangeAddress = CellRangeAddress.valueOf(ranges.get(0));
        }

        String javaScriptFilter = null;
        List<String> javaScriptFilters = Utils.getLiteralObjectsFromQuads(rmlStore.getQuads(source, new NamedNode(NAMESPACES.SS + "javaScriptFilter"), null));
        if (!javaScriptFilters.isEmpty()) {
            javaScriptFilter = javaScriptFilters.get(0);
        }

        Workbook workbook = null;
        String path;

        //load workbook
        //this can take a lot of RAM if workbook has many cells
        if (access instanceof LocalFileAccess) {
            LocalFileAccess localFileAccess = (LocalFileAccess) access;

            path = localFileAccess.getPath();

            if (!(useCache && recordCache.containsKey(path + "-" + sheetName))) {
                if (localFileAccess.getPath().endsWith("xls")) {
                    try {
                        workbook = new HSSFWorkbook(is);
                    } catch (OfficeXmlFileException e) {
                        //fallback
                        workbook = new XSSFWorkbook(is);
                    }
                } else {
                    //xlsx
                    workbook = new XSSFWorkbook(is);
                }
            }

        } else {
            throw new IOException("access has to be LocalFileAccess to excel file");
        }

        List<Record> sheetRecords = null;

        String cacheId = path + "-" + sheetName;
        
        if (useCache && recordCache.containsKey(cacheId)) {
            //use the cache
            //copy the list because it will be filtered
            sheetRecords = new ArrayList<>(recordCache.get(cacheId));
            
        } else {

            //load records
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                List<Record> records = new ArrayList<>();
                
                Sheet sheet = workbook.getSheetAt(i);

                //select it if this is the current one
                if (sheet.getSheetName().equals(sheetName)) {
                    sheetRecords = records;
                }

                Dimension maxima = getMaxima(sheet);
                //int numberOfCells = maxima.width * maxima.height;

                ExcelCell[][] cellMatrix = new ExcelCell[maxima.height][maxima.width];

                int minRow = sheet.getFirstRowNum();
                int maxRow = sheet.getLastRowNum();

                for (int k = minRow; k <= maxRow; k++) {

                    Row row = sheet.getRow(k);

                    int minCol = row.getFirstCellNum();
                    int maxCol = row.getLastCellNum();

                    for (int j = minCol; j <= maxCol; j++) {
                        Cell cell = row.getCell(j);
                        if (cell == null) {
                            continue;
                        }

                        if (ignoreBlank && cell.getCellType() == CellType.BLANK) {
                            continue;
                        }

                        ExcelCell excelCell = new ExcelCell();
                        cellMatrix[cell.getAddress().getRow()][cell.getAddress().getColumn()] = excelCell;

                        int c = cell.getAddress().getColumn();
                        int r = cell.getAddress().getRow();

                        //address
                        excelCell.setRow(r);
                        excelCell.setColumn(c);
                        excelCell.setAddress(cell.getAddress().toString());

                        //cell style and content
                        CellStyle cst = cell.getCellStyle();
                        setStyle(excelCell, cell, cst);
                        setValue(excelCell, cell, cst);

                        //record
                        ExcelRecord record = new ExcelRecord();
                        record.setCellMatrix(cellMatrix);
                        record.setCell(excelCell);

                        records.add(record);

                    }// for columns
                }//for rows
                
                String innerCacheId = path + "-" + sheet.getSheetName();
                
                //save in cache
                if(useCache && !recordCache.containsKey(innerCacheId)) {
                    recordCache.put(innerCacheId, new ArrayList<>(records));
                }
                
            }//for sheets
        }

        //release memory
        if(workbook != null) {
            workbook.close();
            workbook = null;
            //because workbook RAM can be big we clean up
            System.gc();
        }
        
        if(sheetRecords == null) {
            throw new RuntimeException("sheet was not found: " + sheetName);
        }
        
        //filter the records ===================================================
        
        //only the records in the specified range are used
        CellRangeAddress cellRangeAddressFinal = cellRangeAddress;
        if (cellRangeAddressFinal != null) {
            
            sheetRecords.removeIf(record -> { 
                
                int c = ((ExcelRecord)record).getCell().getColumn();
                int r = ((ExcelRecord)record).getCell().getRow();
                
                boolean inRange = 
                        cellRangeAddressFinal.getFirstRow() <= r    && r <= cellRangeAddressFinal.getLastRow() && 
                        cellRangeAddressFinal.getFirstColumn() <= c && c <= cellRangeAddressFinal.getLastColumn();
                    
                return !inRange;
            });
        }
        
        //filter records (so that matrix is initialized)
        //check with filter
        final Bindings jsBindings = jsEngine.createBindings();
        String jsFilter = javaScriptFilter;
        if (jsFilter != null) {
            sheetRecords.removeIf(record -> {

                ExcelRecord excelRecord = (ExcelRecord) record;
                try {
                    jsBindings.put("record", record);
                    jsBindings.put("out", System.out);

                    //short cuts
                    jsBindings.put("cell", excelRecord.getCell());
                    jsBindings.put("row", excelRecord.getCell().getRow());
                    jsBindings.put("column", excelRecord.getCell().getColumn());
                    jsBindings.put("address", excelRecord.getCell().getAddress());

                    jsBindings.put("valueNumeric", excelRecord.getCell().getValueNumeric());
                    jsBindings.put("valueBoolean", excelRecord.getCell().getValueBoolean());
                    jsBindings.put("valueFormula", excelRecord.getCell().getValueFormula());
                    jsBindings.put("valueError", excelRecord.getCell().getValueError());
                    jsBindings.put("valueString", excelRecord.getCell().getValueString());
                    jsBindings.put("valueRichText", excelRecord.getCell().getValueRichText());

                    Object result = jsEngine.eval(jsFilter, jsBindings);

                    boolean filter = ((Boolean) result);

                    return !filter;

                } catch (ScriptException ex) {
                    throw new RuntimeException(ex);
                }
            });
        }

        long end = System.currentTimeMillis();
        
        long duration = end - begin;
        
        System.out.println(
                sheetRecords.size() + " records returned, " + 
                memoryStatistics() + " " + 
                sheetNames + " " + ranges + " " + duration + " ms"
        );
        
        return sheetRecords;
    }

    //if it returns 8x5 it means there are 8 columns and 5 rows filled
    private Dimension getMaxima(Sheet sheet) {

        int w = 0;
        int h = 0;

        int minRow = sheet.getFirstRowNum();
        int maxRow = sheet.getLastRowNum() + 1;
        //row number = 0-based index, that is why + 1

        h = Math.max(h, maxRow);

        for (int k = minRow; k < maxRow; k++) {
            Row row = sheet.getRow(k);
            if (row == null) {
                continue;
            }

            int maxCol = row.getLastCellNum();

            w = Math.max(w, maxCol);
        }

        return new Dimension(w, h);
    }

    private Color toAwtColor(XSSFColor color) {
        if (color == null || color.getARGB() == null) {
            return null;
        }

        return new Color(
                (int) (color.getARGB()[1] & 0xFF), //R
                (int) (color.getARGB()[2] & 0xFF), //G
                (int) (color.getARGB()[3] & 0xFF), //B

                (int) (color.getARGB()[0] & 0xFF) //A
        );
    }

    private void setValue(ExcelCell excelCell, Cell cell, CellStyle cst) {
        CellType formulaCellType = null;
        excelCell.setCellType(cell.getCellTypeEnum().toString().toLowerCase());
        switch (cell.getCellTypeEnum()) {
            case NUMERIC:
                excelCell.setValueNumeric(cell.getNumericCellValue());
                break;
            case BOOLEAN:
                excelCell.setValueBoolean(cell.getBooleanCellValue());
                break;
            case STRING:
                String plainText = cell.getStringCellValue();
                excelCell.setValueString(plainText);

                StringBuilder richTextBuilder = new StringBuilder();

                //rich text
                RichTextString richTextString = cell.getRichStringCellValue();
                if (richTextString instanceof XSSFRichTextString) {
                    XSSFRichTextString rts = (XSSFRichTextString) richTextString;

                    if (rts.hasFormatting()) {

                        //numFormattingRuns is how often the <main:r> tag is in rts.getCTRst().xmlText()
                        //String xmlText = rts.getCTRst().xmlText();
                        for (int r = 0; r < rts.numFormattingRuns(); r++) {

                            //ExcelTextStyle textStyle = new ExcelTextStyle();
                            int begin = rts.getIndexOfFormattingRun(r);
                            int length = rts.getLengthOfFormattingRun(r);
                            String subtext = plainText.substring(begin, begin + length);

                            /*
                            textStyle.setId(textStyleId++);
                            textStyle.setCellId(excelCell.getId());

                            textStyle.setBegin(begin);
                            textStyle.setEnd(begin + length);
                            textStyle.setText(subtext);
                             */
                            XSSFFont font = rts.getFontOfFormattingRun(r);

                            //if font is null use cell's font
                            if (font == null && cst instanceof XSSFCellStyle) {
                                XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
                                XSSFFont cellFont = cs.getFont();
                                font = cellFont;
                            }

                            if (font != null) {
                                /*
                                    <font face='' size='' color=''></font>
                                    <br>
                                    <i></i>
                                    <u></u>
                                    <b></b>
                                    <strike></strike>
                                 */

                                //textStyle.setFontBold(font.getBold());
                                if (font.getBold()) {
                                    richTextBuilder.append("<b>");
                                }

                                //textStyle.setFontItalic(font.getItalic());
                                if (font.getItalic()) {
                                    richTextBuilder.append("<i>");
                                }

                                //textStyle.setFontStrikeout(font.getStrikeout());
                                if (font.getStrikeout()) {
                                    richTextBuilder.append("<strike>");
                                }

                                boolean underline = font.getUnderline() != FontUnderline.NONE.getByteValue();
                                //textStyle.setFontUnderline(underline);
                                if (underline) {
                                    richTextBuilder.append("<u>");
                                }

                                //textStyle.setFontName(font.getFontName());
                                //textStyle.setFontSize(font.getFontHeightInPoints());
                                //textStyle.setFontColor(toAwtColor(font.getXSSFColor()));
                                richTextBuilder.append("<font ");
                                if (font.getFontName() != null) {
                                    richTextBuilder.append("face='" + font.getFontName() + "' ");
                                }

                                Color color = toAwtColor(font.getXSSFColor());

                                if (color != null) {
                                    String hex = String.format("#%02x%02x%02x",
                                            color.getRed(),
                                            color.getGreen(),
                                            color.getBlue()
                                    );
                                    richTextBuilder.append("color='" + hex + "' ");
                                }

                                //size is not correct when shown in JTable
                                //richTextBuilder.append("size='" + font.getFontHeightInPoints() + "' ");
                                richTextBuilder.append(">");

                                richTextBuilder.append(subtext.replace("\n", "<br>"));

                                richTextBuilder.append("</font>");

                                if (underline) {
                                    richTextBuilder.append("</u>");
                                }
                                if (font.getStrikeout()) {
                                    richTextBuilder.append("</strike>");
                                }
                                if (font.getItalic()) {
                                    richTextBuilder.append("</i>");
                                }
                                if (font.getBold()) {
                                    richTextBuilder.append("</b>");
                                }

                            } else {
                                richTextBuilder.append(subtext.replace("\n", "<br>"));
                            }

                        }
                    }
                }

                //set always if string value is available
                excelCell.setValueRichText(richTextBuilder.toString());

                //richTextString.numFormattingRuns() how many changes are made
                //richTextString.toString is unformatted
                break;

            case FORMULA: {
                try {
                    excelCell.setValueFormula(cell.getCellFormula());
                } catch (Exception e) {
                    //ignore
                    //excelCell.setValueFormula("formula error");
                }
                formulaCellType = cell.getCachedFormulaResultTypeEnum();
                break;
            }

            case ERROR:
                excelCell.setValueError(cell.getErrorCellValue());
                break;
        }

        if (formulaCellType != null) {
            excelCell.setCellTypeFormula(formulaCellType.toString().toLowerCase());

            switch (formulaCellType) {
                case NUMERIC:
                    excelCell.setValueNumeric(cell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    excelCell.setValueBoolean(cell.getBooleanCellValue());
                    break;
                case STRING:
                    excelCell.setValueString(cell.getStringCellValue());
                    break;
            }
        }
    }

    private void setStyle(ExcelCell excelCell, Cell cell, CellStyle cst) {
        if (cst instanceof XSSFCellStyle) {
            XSSFCellStyle cs = (XSSFCellStyle) cell.getCellStyle();
            XSSFFont font = cs.getFont();

            excelCell.setFontBold(font.getBold());
            excelCell.setFontItalic(font.getItalic());
            excelCell.setFontStrikeout(font.getStrikeout());
            excelCell.setFontUnderline(font.getUnderline() != FontUnderline.NONE.getByteValue());
            excelCell.setFontName(font.getFontName());
            excelCell.setFontColor(toAwtColor(font.getXSSFColor()));
            excelCell.setFontSize(font.getFontHeightInPoints());

            excelCell.setForegroundColor(toAwtColor(cs.getFillForegroundColorColor()));
            excelCell.setBackgroundColor(toAwtColor(cs.getFillBackgroundColorColor()));
            excelCell.setRotation(cs.getRotation());
            excelCell.setHorizontalAlignment(cs.getAlignmentEnum().toString().toLowerCase());
            excelCell.setVerticalAlignment(cs.getVerticalAlignmentEnum().toString().toLowerCase());
            excelCell.setBorderTop(cs.getBorderTopEnum().toString().toLowerCase());
            excelCell.setBorderLeft(cs.getBorderLeftEnum().toString().toLowerCase());
            excelCell.setBorderRight(cs.getBorderRightEnum().toString().toLowerCase());
            excelCell.setBorderBottom(cs.getBorderBottomEnum().toString().toLowerCase());
        }
    }

    //from https://stackoverflow.com/questions/3758606/how-to-convert-byte-size-into-human-readable-format-in-java
    private static String humanReadableByteCount(long bytes, boolean si) {
        int unit = si ? 1000 : 1024;
        if (bytes < unit) {
            return bytes + " B";
        }
        int exp = (int) (Math.log(bytes) / Math.log(unit));
        String pre = (si ? "kMGTPE" : "KMGTPE").charAt(exp - 1) + (si ? "" : "i");
        return String.format("%.1f %sB", bytes / Math.pow(unit, exp), pre);
    }

    private static String humanReadableByteCount(long bytes) {
        return humanReadableByteCount(bytes, true);
    }

    private static String memoryStatistics() {
        Runtime instance = Runtime.getRuntime();

        StringBuilder sb = new StringBuilder();

        long t = instance.totalMemory();
        long f = instance.freeMemory();
        long u = t - f;
        long m = instance.maxMemory();

        sb
                .append(humanReadableByteCount(u)).append(" used / ")
                .append(humanReadableByteCount(t)).append(" total / ")
                .append(humanReadableByteCount(m)).append(" max");

        return sb.toString();
    }
}
