package parser;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.text.DecimalFormat;
import java.util.*;

public class VvpByQuarter {

    public static List<String> readDataA(Workbook workbook, List<String> tabNames){
        Row mainRow;
        List<String> rowA = new ArrayList<String>();
        mainRow = workbook.getSheet(tabNames.get(0)).getRow(0);
        rowA.add(mainRow.getCell(0).toString());
        mainRow = workbook.getSheet(tabNames.get(0)).getRow(2);
        rowA.add(mainRow.getCell(0).toString());

        for (int i = 4; i <=29 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            rowA.add(mainRow.getCell(0).toString());
        }
        mainRow = workbook.getSheet(tabNames.get(0)).getRow(31);
        rowA.add(mainRow.getCell(0).toString());

        mainRow = workbook.getSheet(tabNames.get(0)).getRow(33);
        rowA.add(mainRow.getCell(0).toString());

        for (int i = 35; i <=59 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            rowA.add(mainRow.getCell(0).toString());
        }
        return rowA;
    }

    public static List<String> readDataB(Workbook workbook, List<String> tabNames){
        Row mainRow;
        List<String> rowB = new ArrayList<String>();
        mainRow = workbook.getSheet(tabNames.get(0)).getRow(2);
        rowB.add(mainRow.getCell(1).toString());


        for (int i = 4; i <=29 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowB.add((formatter.format(mainRow.getCell(1).getNumericCellValue())));
        }

        for (int i = 35; i <=59 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowB.add((formatter.format(mainRow.getCell(1).getNumericCellValue())));
        }

        return rowB;
    }

    public static List<String> readDataC(Workbook workbook, List<String> tabNames){
        Row mainRow;
        List<String> rowC = new ArrayList<String>();
        mainRow = workbook.getSheet(tabNames.get(0)).getRow(2);
        rowC.add(mainRow.getCell(2).toString());

        mainRow = workbook.getSheet(tabNames.get(0)).getRow(3);
        rowC.add(mainRow.getCell(2).toString());

        for (int i = 4; i <=29 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowC.add((formatter.format(mainRow.getCell(2).getNumericCellValue())));
        }

        for (int i = 35; i <=59 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowC.add((formatter.format(mainRow.getCell(2).getNumericCellValue())));
        }

        return rowC;
    }

    public static List<String> readDataD(Workbook workbook, List<String> tabNames){
        Row mainRow;
        List<String> rowD = new ArrayList<String>();
        mainRow = workbook.getSheet(tabNames.get(0)).getRow(2);
        rowD.add(mainRow.getCell(3).toString());

        mainRow = workbook.getSheet(tabNames.get(0)).getRow(3);
        rowD.add(mainRow.getCell(3).toString());

        for (int i = 4; i <=29 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowD.add((formatter.format(mainRow.getCell(3).getNumericCellValue())));
        }

        for (int i = 35; i <=59 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowD.add((formatter.format(mainRow.getCell(3).getNumericCellValue())));
        }

        return rowD;
    }

    public static List<String> readDataE(Workbook workbook, List<String> tabNames){
        Row mainRow;
        List<String> rowE = new ArrayList<String>();
        mainRow = workbook.getSheet(tabNames.get(0)).getRow(2);
        rowE.add(mainRow.getCell(4).toString());

        mainRow = workbook.getSheet(tabNames.get(0)).getRow(3);
        rowE.add(mainRow.getCell(4).toString());

        for (int i = 4; i <=29 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowE.add((formatter.format(mainRow.getCell(4).getNumericCellValue())));
        }

        for (int i = 35; i <=59 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowE.add((formatter.format(mainRow.getCell(4).getNumericCellValue())));
        }

        return rowE;
    }

    public static List<String> readDataF(Workbook workbook, List<String> tabNames){
        Row mainRow;
        List<String> rowF = new ArrayList<String>();
        mainRow = workbook.getSheet(tabNames.get(0)).getRow(2);
        rowF.add(mainRow.getCell(5).toString());

        mainRow = workbook.getSheet(tabNames.get(0)).getRow(3);
        rowF.add(mainRow.getCell(5).toString());

        for (int i = 4; i <=29 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowF.add((formatter.format(mainRow.getCell(5).getNumericCellValue())));
        }

        for (int i = 35; i <=59 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowF.add((formatter.format(mainRow.getCell(5).getNumericCellValue())));
        }

        return rowF;
    }

    public static List<String> readDataH(Workbook workbook, List<String> tabNames){
        Row mainRow;
        List<String> rowH = new ArrayList<String>();
        mainRow = workbook.getSheet(tabNames.get(0)).getRow(2);
        rowH.add(mainRow.getCell(7).toString());

        mainRow = workbook.getSheet(tabNames.get(0)).getRow(3);
        rowH.add(mainRow.getCell(7).toString());

        for (int i = 4; i <=29 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowH.add((formatter.format(mainRow.getCell(7).getNumericCellValue())));
        }

        for (int i = 35; i <=59 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowH.add((formatter.format(mainRow.getCell(7).getNumericCellValue())));
        }

        return rowH;
    }

    public static List<String> readDataI(Workbook workbook, List<String> tabNames){
        Row mainRow;
        List<String> rowI = new ArrayList<String>();
        mainRow = workbook.getSheet(tabNames.get(0)).getRow(2);
        rowI.add(mainRow.getCell(8).toString());

        mainRow = workbook.getSheet(tabNames.get(0)).getRow(3);
        rowI.add(mainRow.getCell(8).toString());

        for (int i = 4; i <=29 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowI.add((formatter.format(mainRow.getCell(8).getNumericCellValue())));
        }

        for (int i = 35; i <=59 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowI.add((formatter.format(mainRow.getCell(8).getNumericCellValue())));
        }

        return rowI;
    }

    public static List<String> readDataJ(Workbook workbook, List<String> tabNames){
        Row mainRow;
        List<String> rowJ = new ArrayList<String>();
        mainRow = workbook.getSheet(tabNames.get(0)).getRow(2);
        rowJ.add(mainRow.getCell(9).toString());

        mainRow = workbook.getSheet(tabNames.get(0)).getRow(3);
        rowJ.add(mainRow.getCell(9).toString());

        for (int i = 4; i <=29 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowJ.add((formatter.format(mainRow.getCell(9).getNumericCellValue())));
        }

        for (int i = 35; i <=59 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowJ.add((formatter.format(mainRow.getCell(9).getNumericCellValue())));
        }

        return rowJ;
    }

    public static List<String> readDataK(Workbook workbook, List<String> tabNames){
        Row mainRow;
        List<String> rowK = new ArrayList<String>();
        mainRow = workbook.getSheet(tabNames.get(0)).getRow(2);
        rowK.add(mainRow.getCell(10).toString());

        mainRow = workbook.getSheet(tabNames.get(0)).getRow(3);
        rowK.add(mainRow.getCell(10).toString());

        for (int i = 4; i <=29 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowK.add((formatter.format(mainRow.getCell(10).getNumericCellValue())));
        }

        for (int i = 35; i <=59 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowK.add((formatter.format(mainRow.getCell(10).getNumericCellValue())));
        }

        return rowK;
    }

    public static List<String> readDataL(Workbook workbook, List<String> tabNames){
        Row mainRow;
        List<String> rowL = new ArrayList<String>();
        mainRow = workbook.getSheet(tabNames.get(0)).getRow(2);
        rowL.add(mainRow.getCell(11).toString());

        mainRow = workbook.getSheet(tabNames.get(0)).getRow(3);
        rowL.add(mainRow.getCell(11).toString());

        for (int i = 4; i <=29 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowL.add((formatter.format(mainRow.getCell(11).getNumericCellValue())));
        }

        for (int i = 35; i <=59 ; i++) {
            mainRow = workbook.getSheet(tabNames.get(0)).getRow(i);
            DecimalFormat formatter = new DecimalFormat("#,##0.0");
            formatter.setMaximumFractionDigits(1);
            rowL.add((formatter.format(mainRow.getCell(11).getNumericCellValue())));
        }
        return rowL;
    }
}
