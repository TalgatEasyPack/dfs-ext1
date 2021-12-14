package com.dfs.ext1.Hadlers;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;

import com.dfs.ext1.Base64Helper.RequestBodyJson;
import com.google.gson.Gson;
import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpHandler;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

enum RowRecordType {
    DOCUMENT,
    TEXT_ROW,
    TEXT_ROW_BOTTOM,
    SIGN_ROW,
    SGIN_ROW_BOTTOM
}

enum ReadPlace {
    TOP, BOTTOM
}

class RowRecord {
    public RowRecordType type;
    public String text = "";
    public Integer qr_code;

}

public class MySignListForExcel2 implements HttpHandler {
    @Override
    public void handle(HttpExchange t) throws IOException {
        try {

            // ---------------------------
            // Read content
            // ---------------------------

            Gson g = new Gson();

            RequestBodyJson requestBody = g.fromJson(new InputStreamReader(t.getRequestBody(), "UTF-8"),
                    RequestBodyJson.class);

            // ---------------------------
            // Decode content and open Stream
            // ---------------------------

            byte[] decoded_excel = Base64.getDecoder().decode(requestBody.base64_excel);

            InputStream is_excel = new ByteArrayInputStream(decoded_excel);

            byte[] decoded_word = Base64.getDecoder().decode(requestBody.base64_word);

            InputStream is_word = new ByteArrayInputStream(decoded_word);

            // ---------------------------
            // Create files
            // ---------------------------

            XSSFWorkbook workbook = new XSSFWorkbook(is_excel);

            XWPFDocument document = new XWPFDocument(is_word);

            // ---------------------------
            // Close decode stream
            // ---------------------------

            is_excel.close();
            is_word.close();

            // ---------------------------
            // Create Stream for responce
            // ---------------------------

            ByteArrayOutputStream bos = mergeExcelAndWord(workbook, document);

            // ---------------------------
            // Close files
            // ---------------------------

            document.close();
            workbook.close();

            // ---------------------------
            // Create json for responce
            // ---------------------------

            requestBody = new RequestBodyJson();

            requestBody.base64_excel = Base64.getEncoder().encodeToString(bos.toByteArray());

            // ---------------------------
            // Convert json to responce
            // ---------------------------

            // String response = "This is the response";

            String response = g.toJson(requestBody);

            // ---------------------------

            t.sendResponseHeaders(200, response.length());

            OutputStream os = t.getResponseBody();

            os.write(response.getBytes());

            os.close();

            // ---------------------------
            // End
            // ---------------------------

        } catch (Exception e) {
            System.out.println(e.getMessage());

            t.getResponseBody().close();
        }
    }

    public static ByteArrayOutputStream mergeExcelAndWord(XSSFWorkbook workbook, XWPFDocument document)
            throws IOException {

        // ---------------------------
        // Create result
        // ---------------------------

        ByteArrayOutputStream result = new ByteArrayOutputStream();

        // ---------------------------------------------------------------------------------------------------------------------------
        // Пересматриваем Word документ
        // ---------------------------------------------------------------------------------------------------------------------------
        try {

            List<IBodyElement> bodyElements = document.getBodyElements();// Список рассматриваемых элементов
            List<RowRecord> rowRecords_Top = new ArrayList<RowRecord>();// Список считываемых элементов
            List<RowRecord> rowRecords_Bottom = new ArrayList<RowRecord>();// Список считываемых элементов
            ReadPlace readPlace = ReadPlace.TOP;// Точка чтения - Выше или ниже вклеиваемого документа

            // Пересматриваем Word документ
            for (IBodyElement iBodyElement : bodyElements) {

                if (iBodyElement instanceof XWPFTable) {// Если элемент - Таблица

                    XWPFTable table = (XWPFTable) iBodyElement;

                    // Пересматриваемы строки в таблице
                    for (int i = 0; i < table.getNumberOfRows(); i++) {

                        XWPFTableRow tableRow = table.getRow(i);

                        // Получаем список ячеек в строке
                        List<XWPFTableCell> tableCells = tableRow.getTableCells();

                        // Создаем новую запись о считанном элементе
                        RowRecord rowRecord = new RowRecord();

                        // Предопределяем строку, для извлечения текста - перед склеивымым документом
                        // она слева
                        XWPFTableCell tableCell = tableCells.get(0);

                        if (tableCells.size() == 1) { // Если строка - обычный текст

                            rowRecord.type = RowRecordType.TEXT_ROW;

                        } else if (tableCells.size() == 2) { // Если строка - данные о подписи

                            rowRecord.type = RowRecordType.SIGN_ROW;

                            // Предопределяем ячейку для извлечения QR кода - перед склеивымым документом
                            // она справа
                            XWPFTableCell imageTableCell = tableCells.get(1);

                            // Пересматриваем точки ввода инфы в зависимости от точки чтения
                            if (readPlace == ReadPlace.BOTTOM) {

                                // Предопределяем ячейку для извлечения QR кода - после склеиваемого документа
                                // она справа
                                tableCell = tableCells.get(1);

                                // Предопределяем ячейку для извлечения QR кода - после склеиваемого документа
                                // она слева
                                imageTableCell = tableCells.get(0);

                            }

                            // --------------------------------------------------
                            // Извлечение QR кода
                            // --------------------------------------------------

                            List<XWPFParagraph> lParagraphs = imageTableCell.getParagraphs();

                            if (lParagraphs.size() > 0) {

                                List<XWPFRun> lRuns = lParagraphs.get(0).getRuns();

                                if (lRuns.size() > 0) {

                                    List<XWPFPicture> lPictures = lRuns.get(0).getEmbeddedPictures();

                                    if (lPictures.size() > 0) {

                                        rowRecord.qr_code = workbook.addPicture(
                                                lPictures.get(0).getPictureData().getData(), Workbook.PICTURE_TYPE_PNG);

                                    }

                                }

                            }

                            // --------------------------------------------------

                        }

                        // --------------------------------------------------
                        // Извлечение текста
                        // --------------------------------------------------

                        List<XWPFParagraph> lParagraphs = tableCell.getParagraphs();

                        List<String> text = new ArrayList<String>();

                        for (int j = 0; j < lParagraphs.size(); j++) {

                            String line = lParagraphs.get(j).getText();

                            if (!(line == null || line.isEmpty() || line.trim().isEmpty())) {

                                text.add(line);

                            }

                        }

                        for (int j = 0; j < text.size(); j++) {

                            rowRecord.text += text.get(j);

                            if (j + 1 < text.size()) {

                                rowRecord.text += "\n";

                            }

                        }

                        // --------------------------------------------------

                        // Добавляем данные об элементе
                        if (readPlace == ReadPlace.TOP) {

                            rowRecords_Top.add(rowRecord);

                        } else {

                            rowRecords_Bottom.add(rowRecord);

                        }

                    }

                } else if (iBodyElement instanceof XWPFParagraph) {// Если элемент - Текстовая строка

                    XWPFParagraph paragraph = (XWPFParagraph) iBodyElement;

                    // Если данные - поле слияния - накидываем отметки об этом
                    if (paragraph.getCTP().toString().contains("MERGEFIELD")) {

                        // Меняем отметку о мечте чтения - после вклеимового документа
                        readPlace = ReadPlace.BOTTOM;

                    }

                }

            }

            // ---------------------------------------------------------------------------------------------------------------------------
            // Подготавливаем данные для редакции Excel
            // ---------------------------------------------------------------------------------------------------------------------------

            Short height = 57 * 20;

            Short indient = 9;

            Double scale = 1.0;

            // ---------------------------

            XSSFFont font_Normal = workbook.createFont();

            String fontName = "Times New Roman";

            font_Normal.setFontName(fontName);

            XSSFFont font_Bold = workbook.createFont();

            font_Bold.setFontName(fontName);
            font_Bold.setBold(true);

            // ---------------------------

            XSSFFormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            DataFormatter formatter = new DataFormatter(true);

            // ---------------------------------------------------------------------------------------------------------------------------
            // Изменяем Excel докумет
            // ---------------------------------------------------------------------------------------------------------------------------

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {

                // ---------------------------------------------------------------------------------------------------------------------------
                // Подготавливаем данные для изменения листа документа
                // ---------------------------------------------------------------------------------------------------------------------------

                XSSFSheet sheet = workbook.getSheetAt(i);

                // ---------------------------

                XSSFCreationHelper helper = workbook.getCreationHelper();

                // ---------------------------

                scale = 1.0;

                if (sheet.getColumnWidth(0) == 2048)
                    scale = 0.9;

                // ---------------------------
                    
                int lastRowNum = determineRowCount(evaluator, formatter, sheet);

                // select last visible column in print area

                int rightColumn = 0;
                int leftColumn = sheet.getLeftCol();

                for (int index = sheet.getFirstRowNum(); index < lastRowNum; index++) {

                    XSSFRow row = sheet.getRow( index );

                    if( row != null ){

                        int colMax = row.getLastCellNum()-1;

                        if( colMax > rightColumn ){

                            rightColumn = colMax;

                        }

                    }
                    
                }
                
                // ---------------------------------------------------------------------------------------------------------------------------
                // Изменяем лист документа
                // ---------------------------------------------------------------------------------------------------------------------------

                for (int j = rowRecords_Top.size() - 1; j >= 0; j--) {

                    RowRecord rowRecord = rowRecords_Top.get(j);

                    if (rowRecord.type == RowRecordType.TEXT_ROW && rowRecord.text.isEmpty() == false) {

                        sheet.shiftRows( 0, sheet.getPhysicalNumberOfRows(), 1);

                        XSSFRow row = sheet.createRow(0);

                        XSSFCell cell = row.createCell( leftColumn );

                        cell.setCellType(CellType.STRING);

                        cell.setCellValue(rowRecord.text);

                        XSSFCellStyle cellStyle = workbook.createCellStyle();

                        cellStyle.setFont(font_Normal);

                        cellStyle.setWrapText(true);

                        cellStyle.setAlignment(HorizontalAlignment.RIGHT);

                        cell.setCellStyle(cellStyle);

                        if (rightColumn > leftColumn)
                            sheet.addMergedRegion(new CellRangeAddress(
                                    0, 0, leftColumn, rightColumn));

                        int numberOfLines = rowRecord.text.split("\n").length + 1;

                        row.setHeightInPoints(numberOfLines * sheet.getDefaultRowHeightInPoints());

                        lastRowNum += 1;

                    } else if (rowRecord.type == RowRecordType.SIGN_ROW
                            && (rowRecord.text.isEmpty() == false || rowRecord.qr_code != null)) {

                        sheet.shiftRows( 0, sheet.getPhysicalNumberOfRows(), 1);

                        XSSFRow row = sheet.createRow(0);

                        if (rightColumn > leftColumn)
                            sheet.addMergedRegion(new CellRangeAddress(
                                    0, 0, leftColumn, rightColumn));

                        row.setHeight(height);

                        if (rowRecord.text.isEmpty() == false) {

                            XSSFCell cell = row.createCell(leftColumn);

                            cell.setCellType(CellType.STRING);

                            cell.setCellValue(rowRecord.text);

                            XSSFCellStyle cellStyle = workbook.createCellStyle();

                            cellStyle.setFont(font_Normal);

                            cellStyle.setWrapText(true);

                            cellStyle.setAlignment(HorizontalAlignment.RIGHT);

                            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

                            cellStyle.setIndention(indient);

                            cell.setCellStyle(cellStyle);

                            int numberOfLines = rowRecord.text.split("\n").length + 1;

                            if (numberOfLines > 4) {

                                row.setHeightInPoints(numberOfLines * sheet.getDefaultRowHeightInPoints());

                            }

                        }

                        if (rowRecord.qr_code != null) {

                            XSSFDrawing drawing = sheet.createDrawingPatriarch();

                            XSSFClientAnchor anchor = helper.createClientAnchor();

                            anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_DO_RESIZE);

                            anchor.setCol1(rightColumn + 1);
                            anchor.setRow1(1);
                            anchor.setCol2(rightColumn + 1);
                            anchor.setRow2(1);

                            anchor.setDx1(sheet.getColumnWidth(rightColumn + 1) - Units.toEMU(57));
                            anchor.setDy1(Units.toEMU(0));

                            anchor.setDx2(sheet.getColumnWidth(rightColumn + 1));
                            anchor.setDy2(Units.toEMU(57));

                            drawing.createPicture(anchor, rowRecord.qr_code);

                        }

                        lastRowNum += 1;

                    }
                }

                for (int j = 0; j < rowRecords_Bottom.size(); j++) {

                    RowRecord rowRecord = rowRecords_Bottom.get(j);

                    if (rowRecord.type == RowRecordType.TEXT_ROW && rowRecord.text.isEmpty() == false) {

                        XSSFRow row = sheet.createRow(lastRowNum);

                        XSSFCell cell = row.createCell(leftColumn);

                        cell.setCellType(CellType.STRING);

                        cell.setCellValue(rowRecord.text);

                        XSSFCellStyle cellStyle = workbook.createCellStyle();

                        cellStyle.setFont(font_Bold);

                        cell.setCellStyle(cellStyle);

                        lastRowNum += 1;

                    } else if (rowRecord.type == RowRecordType.SIGN_ROW
                            && (rowRecord.text.isEmpty() == false || rowRecord.qr_code != null)) {

                        XSSFRow row = sheet.createRow(lastRowNum);

                        if (rightColumn > leftColumn)
                            sheet.addMergedRegion(new CellRangeAddress(
                                    lastRowNum, lastRowNum, leftColumn, rightColumn));

                        row.setHeight(height);

                        if (rowRecord.text.isEmpty() == false) {

                            XSSFCell cell = row.createCell(leftColumn);

                            cell.setCellType(CellType.STRING);

                            cell.setCellValue(rowRecord.text);

                            XSSFCellStyle cellStyle = workbook.createCellStyle();

                            cellStyle.setFont(font_Normal);

                            cellStyle.setWrapText(true);

                            cellStyle.setIndention(indient);

                            cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

                            cell.setCellStyle(cellStyle);

                            int numberOfLines = rowRecord.text.split("\n").length + 1;

                            if (numberOfLines > 5) {

                                row.setHeightInPoints(numberOfLines * sheet.getDefaultRowHeightInPoints());

                            }

                        }

                        if (rowRecord.qr_code != null) {

                            XSSFDrawing drawing = sheet.createDrawingPatriarch();

                            XSSFClientAnchor anchor = helper.createClientAnchor();

                            anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_DO_RESIZE);

                            anchor.setCol1(leftColumn);
                            anchor.setRow1(lastRowNum);
                            anchor.setCol2(leftColumn);
                            anchor.setRow2(lastRowNum);

                            anchor.setDx2(Units.toEMU(57));
                            anchor.setDy2(Units.toEMU(57));

                            XSSFPicture pict = drawing.createPicture(anchor, rowRecord.qr_code);

                            pict.resize(scale, 1.0);

                        }

                        lastRowNum += 2;

                    }

                }

                // ---------------------------------------------------------------------------------------------------------------------------

            }

        } catch (Exception e) {

            e.printStackTrace();

        }finally {

            workbook.write(result);

        }
        // ---------------------------
        // return result
        // ---------------------------

        return result;

        // ---------------------------
        // End
        // ---------------------------

    }

    private static int determineRowCount(XSSFFormulaEvaluator evaluator, DataFormatter formatter, XSSFSheet sheet) {
        int lastRowIndex = -1;

        if (sheet.getPhysicalNumberOfRows() > 0) {
            // getLastRowNum() actually returns an index, not a row number
            lastRowIndex = sheet.getLastRowNum();

            // now, start at end of spreadsheet and work our way backwards until we find a
            // row having data
            for (; lastRowIndex >= 0; lastRowIndex--) {
                XSSFRow row = sheet.getRow(lastRowIndex);
                if (!isRowEmpty(evaluator, formatter, row)) {
                    break;
                }
            }

        }

        lastRowIndex++;

        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {

            CellRangeAddress region = sheet.getMergedRegion(i);

            if( region.getFirstRow() >= lastRowIndex ){

                sheet.removeMergedRegion( i );

            }

        }

        return lastRowIndex;
    }

    /**
     * Determine whether a row is effectively completely empty - i.e. all cells
     * either contain an empty string or nothing.
     */
    private static boolean isRowEmpty(XSSFFormulaEvaluator evaluator, DataFormatter formatter, XSSFRow row) {
        if (row == null) {
            return true;
        }

        int cellCount = row.getLastCellNum() + 1;
        for (int i = 0; i < cellCount; i++) {
            String cellValue = getCellValue(evaluator, formatter, row, i);
            if (cellValue != null && cellValue.length() > 0) {
                return false;
            }
        }
        return true;
    }

    /**
     * Get the effective value of a cell, formatted according to the formatting of
     * the cell.
     * If the cell contains a formula, it is evaluated first, then the result is
     * formatted.
     * 
     * @param row         the row
     * @param columnIndex the cell's column index
     * @return the cell's value
     */
    private static String getCellValue(XSSFFormulaEvaluator evaluator, DataFormatter formatter, XSSFRow row,
            int columnIndex) {
        String cellValue;
        XSSFCell cell = row.getCell(columnIndex);
        if (cell == null) {
            // no data in this cell
            cellValue = null;
        } else {
            if (cell.getCellType() != CellType.FORMULA) {
                // cell has a value, so format it into a string
                cellValue = formatter.formatCellValue(cell);
            } else {
                // cell has a formula, so evaluate it
                cellValue = formatter.formatCellValue(cell, evaluator);
            }
        }
        return cellValue;
    }
}
