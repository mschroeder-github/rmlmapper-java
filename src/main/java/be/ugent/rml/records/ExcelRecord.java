package be.ugent.rml.records;

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
                
                //TODO use address here for absolute positioning
                
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

            case "valueNumeric":
                objects.add(currentCell.getValueNumeric());
                break;
            case "valueBoolean":
                objects.add(currentCell.getValueBoolean());
                break;
            case "valueFormular":
                objects.add(currentCell.getValueFormular());
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

    void setCell(ExcelCell cell) {
        this.cell = cell;
    }

}
