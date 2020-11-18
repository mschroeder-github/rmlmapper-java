package be.ugent.rml.records;

import java.awt.Color;
import net.minidev.json.JSONObject;

/**
 * An Excel cell is part of a workbook.
 * We put all the data loaded with the POI library into this class so that
 * the RAM fingerprint becomes smaller.
 * @author Markus Schr&ouml;der
 */
public class ExcelCell {
    
    private int row;
    private int column;
    private String address;
    
    private boolean fontBold;
    private boolean fontItalic;
    private boolean fontStrikeout;
    private boolean fontUnderline;
    private int fontSize;
    private String fontName;
    private Color fontColor;
    
    private Color foregroundColor;
    private Color backgroundColor;
    private int rotation;
    private String horizontalAlignment;
    private String verticalAlignment;
    
    private String borderTop;
    private String borderLeft;
    private String borderRight;
    private String borderBottom;
    
    private String cellType;
    private String cellTypeFormular;
    
    private double valueNumeric;
    private boolean valueBoolean;
    private String valueFormular;
    private int valueError;
    private String valueString;
    private String valueRichText;

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

    public int getColumn() {
        return column;
    }

    public void setColumn(int column) {
        this.column = column;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public boolean isFontBold() {
        return fontBold;
    }

    public void setFontBold(boolean fontBold) {
        this.fontBold = fontBold;
    }

    public boolean isFontItalic() {
        return fontItalic;
    }

    public void setFontItalic(boolean fontItalic) {
        this.fontItalic = fontItalic;
    }

    public boolean isFontStrikeout() {
        return fontStrikeout;
    }

    public void setFontStrikeout(boolean fontStrikeout) {
        this.fontStrikeout = fontStrikeout;
    }

    public boolean isFontUnderline() {
        return fontUnderline;
    }

    public void setFontUnderline(boolean fontUnderline) {
        this.fontUnderline = fontUnderline;
    }

    public String getFontName() {
        return fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public Color getFontColor() {
        return fontColor;
    }

    public void setFontColor(Color fontColor) {
        this.fontColor = fontColor;
    }

    public int getFontSize() {
        return fontSize;
    }

    public void setFontSize(int fontSize) {
        this.fontSize = fontSize;
    }
    
    public Color getForegroundColor() {
        return foregroundColor;
    }

    public void setForegroundColor(Color foregroundColor) {
        this.foregroundColor = foregroundColor;
    }

    public Color getBackgroundColor() {
        return backgroundColor;
    }

    public void setBackgroundColor(Color backgroundColor) {
        this.backgroundColor = backgroundColor;
    }

    public int getRotation() {
        return rotation;
    }

    public void setRotation(int rotation) {
        this.rotation = rotation;
    }

    public String getHorizontalAlignment() {
        return horizontalAlignment;
    }

    public void setHorizontalAlignment(String horizontalAlignment) {
        this.horizontalAlignment = horizontalAlignment;
    }

    public String getVerticalAlignment() {
        return verticalAlignment;
    }

    public void setVerticalAlignment(String verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    public String getBorderTop() {
        return borderTop;
    }

    public void setBorderTop(String borderTop) {
        this.borderTop = borderTop;
    }

    public String getBorderLeft() {
        return borderLeft;
    }

    public void setBorderLeft(String borderLeft) {
        this.borderLeft = borderLeft;
    }

    public String getBorderRight() {
        return borderRight;
    }

    public void setBorderRight(String borderRight) {
        this.borderRight = borderRight;
    }

    public String getBorderBottom() {
        return borderBottom;
    }

    public void setBorderBottom(String borderBottom) {
        this.borderBottom = borderBottom;
    }

    public String getCellType() {
        return cellType;
    }

    public void setCellType(String cellType) {
        this.cellType = cellType;
    }

    public String getCellTypeFormular() {
        return cellTypeFormular;
    }

    public void setCellTypeFormular(String cellTypeFormular) {
        this.cellTypeFormular = cellTypeFormular;
    }

    public double getValueNumeric() {
        return valueNumeric;
    }

    public void setValueNumeric(double valueNumeric) {
        this.valueNumeric = valueNumeric;
    }

    public boolean getValueBoolean() {
        return valueBoolean;
    }

    public void setValueBoolean(boolean valueBoolean) {
        this.valueBoolean = valueBoolean;
    }

    public String getValueFormular() {
        return valueFormular;
    }

    public void setValueFormular(String valueFormular) {
        this.valueFormular = valueFormular;
    }

    public int getValueError() {
        return valueError;
    }

    public void setValueError(int valueError) {
        this.valueError = valueError;
    }

    public String getValueString() {
        return valueString;
    }

    public void setValueString(String valueString) {
        this.valueString = valueString;
    }

    public String getValueRichText() {
        return valueRichText;
    }

    public void setValueRichText(String valueRichText) {
        this.valueRichText = valueRichText;
    }
    
    public String getValue() {
        if(cellType.equals("numeric")) {
            if(getValueNumeric() % 1 == 0) {
                //integer value
                return String.valueOf((int) getValueNumeric());
            } else {
                return String.valueOf(getValueNumeric());
            }
        } else if(cellType.equals("boolean")) {
            return String.valueOf(getValueBoolean());
        } else if(cellType.equals("formular")) {
            return String.valueOf(getValueFormular());
        } else if(cellType.equals("error")) {
            return String.valueOf(getValueError());
        } else if(cellType.equals("string")) {
            return String.valueOf(getValueString());
        }
        
        return "";
    }
    
    public JSONObject toJSON() {
        JSONObject json = new JSONObject();
        
        json.put("row", row);
        json.put("column", column);
        json.put("address", address);
        
        json.put("cellType", cellType);
        json.put("cellTypeFormular", cellTypeFormular);
        
        json.put("valueNumeric", valueNumeric);
        json.put("valueBoolean", valueBoolean);
        json.put("valueFormular", valueFormular);
        json.put("valueError", valueError);
        json.put("valueString", valueString);
        json.put("valueRichText", valueRichText);
        
        return json;
    }
    
    @Override
    public String toString() {
        return "ExcelCell{ row=" + row + ", column=" + column + ", address=" + address + ", cellType=" + cellType + ", valueNumeric=" + valueNumeric + ", valueBoolean=" + valueBoolean + ", valueFormular=" + valueFormular + ", valueError=" + valueError + ", valueString=" + valueString + '}';
    }
    
}
