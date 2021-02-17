package be.ugent.rml.records;

import java.awt.Color;
import java.util.ArrayList;
import java.util.List;

/**
 * This class is a specific implementation of a record for Excel.
 *
 * One record is one cell in the
 */
public class ExcelRecord extends Record {

    //[height][width] => row * column
    private ExcelCell[][] cellMatrix;

    private ExcelCell cell;

    @Override
    public List<Object> get(String value) {
        //value is the "{value}" in rr:template or rml:reference 
        List<Object> objects = new ArrayList<>();

        ExcelCell currentCell = cell;
        
        

        //syntax proposal: "(-1,0).valueString", "valueString", "[A5].valueString"
        //[] = absolute
        //() = relative
        if (value.contains(".")) {
            String[] split = value.split("\\.");
            String left = split[0].trim();

            boolean relative = left.startsWith("(") && left.endsWith(")");
            boolean absolute = left.startsWith("[") && left.endsWith("]"); 

            //remove brackets
            left = left.substring(1, left.length() - 1);

            if (relative) {
                String[] posParts = left.split("\\,");
                
                int x = Integer.parseInt(posParts[0]);
                int y = Integer.parseInt(posParts[1]);
                
                y = currentCell.getRow() + y;
                x = currentCell.getColumn() + x;
                
                currentCell = cellMatrix[y][x];
            } else if(absolute) {
                String[] posParts = left.split("\\,");
                
                int x = currentCell.getColumn();
                int y = currentCell.getRow();
                
                //the case "[0]" is x = 0
                if(posParts.length == 1) {
                    if(!posParts[0].trim().isEmpty()) {
                        x = Integer.parseInt(posParts[0]);
                    }
                } else if(posParts.length == 2) {
                    //the case [4,5] but also [,5]
                    if(!posParts[0].trim().isEmpty()) {
                        x = Integer.parseInt(posParts[0]);
                    }
                    if(!posParts[1].trim().isEmpty()) {
                        y = Integer.parseInt(posParts[1]);
                    }
                }
                
                currentCell = cellMatrix[y][x];
            }


            value = split[1];
        }

        if (currentCell == null) {
            return objects;
        }
        
        switch (value) {
            case "row":
                objects.add(currentCell.getRow());
                break;
            case "column":
                objects.add(currentCell.getColumn());
                break;
            case "address":
                objects.add(currentCell.getAddress());
                break;
                
            case "backgroundColor":
                Color bg = currentCell.getBackgroundColor();
                if(bg != null) {
                    objects.add(Integer.toHexString(bg.getRGB()).substring(2));
                }
                break;
            case "foregroundColor":
                Color fg = currentCell.getForegroundColor();
                if(fg != null) {
                    objects.add(Integer.toHexString(fg.getRGB()).substring(2));
                }
                break;
                
            case "fontColor":
                Color fontColor = currentCell.getFontColor();
                if(fontColor != null) {
                    objects.add(Integer.toHexString(fontColor.getRGB()).substring(2));
                }
                break;
            case "fontName":
                String fontName = currentCell.getFontName();
                if(fontName != null) {
                    objects.add(fontName);
                }
                break;
            case "fontSize":
                objects.add(currentCell.getFontSize());
                break;

            case "valueNumeric":
                objects.add(currentCell.getValueNumeric());
                break;
            case "valueInt":
                objects.add((int) currentCell.getValueNumeric());
                break;
            case "valueBoolean":
                objects.add(currentCell.getValueBoolean());
                break;
            case "valueFormula":
                objects.add(currentCell.getValueFormula());
                break;
            case "valueError":
                objects.add(currentCell.getValueError());
                break;
            case "valueString":
                objects.add(currentCell.getValueString());
                break;
            case "valueRichText":
                objects.add(currentCell.getValueRichText());
                break;
            case "json":
                objects.add(currentCell.toJSON().toString());
                break;
            case "value":
                objects.add(currentCell.getValue());
                break;
        }

        return objects;
    }

    public ExcelCell[][] getCellMatrix() {
        return cellMatrix;
    }

    /*package*/ void setCellMatrix(ExcelCell[][] cellMatrix) {
        this.cellMatrix = cellMatrix;
    }

    public ExcelCell getCell() {
        return cell;
    }

    /*package*/ void setCell(ExcelCell cell) {
        this.cell = cell;
    }

}
