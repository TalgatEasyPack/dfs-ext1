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

import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Footer;
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
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

public class MySignListForExcel implements HttpHandler {
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
        // extract word
        // ---------------------------

        List<IBodyElement> bodyElements = document.getBodyElements();

        String fontName = "Times New Roman";

        //---------------------------------------------------------------------------------------------------------------------------
        // Таблица отв. исполнителей
        XWPFTable tableManagers;
        String tableManagers_row_0_cell_0 = null;
        List<Integer> tableManagers_row_N_cell_0  = new ArrayList<Integer>();   // Подпись отв. исполнителей
        List<String> tableManagers_row_N_cell_1   = new ArrayList<String>();    // Инициалы отв. исполнителей

        if( bodyElements.get(1) instanceof XWPFTable){

            tableManagers = (XWPFTable) bodyElements.get(1);

            // Подбираем Фонт
            fontName                 = tableManagers.getRow( 0 ).getCell( 0 ).getParagraphArray(0).getRuns().get(0).getFontName();

            // Наименование позиции отв. исполнителей
            tableManagers_row_0_cell_0 = tableManagers.getRow( 0 ).getCell( 0 ).getText();

            // Заполняем данные о подписях
            for (int i = 1; i < tableManagers.getNumberOfRows() ; i++) {

                int image = -1;

                List<XWPFRun> runs = tableManagers.getRow( i ).getCell( 0 ).getParagraphArray(0).getRuns();
                
                if( runs.size() > 0 ){

                    List<XWPFPicture> list = tableManagers.getRow( i ).getCell( 0 ).getParagraphArray( 0 ).getRuns().get(0).getEmbeddedPictures();

                    if( list.size() > 0 ){

                        image = workbook.addPicture( tableManagers.getRow( i ).getCell( 0 ).
                            getParagraphArray(0).getRuns().get(0).
                            getEmbeddedPictures().get(0).getPictureData().getData(), 
                            Workbook.PICTURE_TYPE_PNG);
        
                    }

                }

                tableManagers_row_N_cell_0.add( image );

                tableManagers_row_N_cell_1.add(
                    tableManagers.getRow( i ).getCell( 1 ).getText()
                );

            }

        }


        //---------------------------------------------------------------------------------------------------------------------------

        //---------------------------------------------------------------------------------------------------------------------------
        // Таблица отв. исполнителей

        XWPFTable tableWorkgroup;
        String tableWorkgroup_row_0_cell_0 = null;
        List<Integer> tableWorkgroup_row_N_cell_0 = new ArrayList<Integer>();   // Подпись соисполнителей
        List<String> tableWorkgroup_row_N_cell_1 = new ArrayList<String>();     // Инициалы соисполнителей

        if( bodyElements.get(1) instanceof XWPFTable){

            tableWorkgroup = (XWPFTable) bodyElements.get(3);

            // Наименование позиции соисполнителей
            tableWorkgroup_row_0_cell_0 = tableWorkgroup.getRow( 0 ).getCell( 0 ).getText();

            // Заполняем данные о подписях
            for (int i = 1; i < tableWorkgroup.getNumberOfRows() ; i++) {
                
                int image = -1;

                List<XWPFRun> runs = tableWorkgroup.getRow( i ).getCell( 0 ).getParagraphArray(0).getRuns();
                
                if( runs.size() > 0 ){

                    List<XWPFPicture> list = tableWorkgroup.getRow( i ).getCell( 0 ).getParagraphArray( 0 ).getRuns().get(0).getEmbeddedPictures();

                    if( list.size() > 0 ){

                        image = workbook.addPicture( tableWorkgroup.getRow( i ).getCell( 0 ).
                            getParagraphArray(0).getRuns().get(0).
                            getEmbeddedPictures().get(0).getPictureData().getData(), 
                            Workbook.PICTURE_TYPE_PNG);
        
                    }

                }

                tableWorkgroup_row_N_cell_0.add( image );

                tableWorkgroup_row_N_cell_1.add(
                    tableWorkgroup.getRow( i ).getCell( 1 ).getText()
                );

            }

        }
        //---------------------------------------------------------------------------------------------------------------------------

        String footerWord = document.getFooterArray(0).getText();

        // ---------------------------
        // Change Excel
        // ---------------------------

        XSSFSheet sheet = workbook.getSheetAt(0);

        XSSFCreationHelper helper = workbook.getCreationHelper();

        // size's

        Short height = 57 * 20;
        
        Short indient = 9;                
        
        Double scale = 1.0;
        if( sheet.getColumnWidth(0) == 2048 )
            scale = 0.9;

        // create fonts

        XSSFFont font_Normal = workbook.createFont();

        font_Normal.setFontName(fontName);

        XSSFFont font_Bold = workbook.createFont();

        font_Bold.setFontName(fontName);
        font_Bold.setBold( true );

        // select last visible column in print area

        int rightColumn = 0;
        int cellWidth = sheet.getColumnWidth( rightColumn );

        while( cellWidth < 10240 ){

            rightColumn++;

            cellWidth += sheet.getColumnWidth( rightColumn ) ;
            
        }

        int rightColumnHead = 0;
        cellWidth = sheet.getColumnWidth( rightColumnHead );

        while( cellWidth < 16384 ){

            rightColumnHead++;

            cellWidth += sheet.getColumnWidth( rightColumnHead ) ;
            
        }

        // Создаем строку с Наименованием позиции Отв. Исполнителей

        int lastRowNum = sheet.getLastRowNum() + 1;
        
        if( !( tableManagers_row_0_cell_0 == null || tableManagers_row_0_cell_0.isEmpty() || tableManagers_row_0_cell_0.trim().isEmpty() ) ){

            XSSFRow row  = sheet.createRow( lastRowNum );

            XSSFCell cell  = row.createCell(0);

            cell.setCellType(CellType.STRING);
    
            cell.setCellValue(tableManagers_row_0_cell_0);
    
            XSSFCellStyle cellStyle  = workbook.createCellStyle();
    
            cellStyle.setFont(font_Bold);
    
            cell.setCellStyle(cellStyle);

            lastRowNum += 2;

        }

        // Создаем строки с Подписями и Инициалами Отв. Исполнителей

        for (int i = 0; i < tableManagers_row_N_cell_1.size(); i++) {

            String value = tableManagers_row_N_cell_1.get( i );

            if( !( value == null || value.isEmpty() || value.trim().isEmpty() ) ){
            
                XSSFRow row  = sheet.createRow( lastRowNum );            
                
                if( rightColumn > 0 )
                    sheet.addMergedRegion( new CellRangeAddress(
                        lastRowNum, lastRowNum, 0, rightColumn
                ));
                
                row.setHeight( height );

                XSSFCell cell  = row.createCell(0);

                cell.setCellType(CellType.STRING);
        
                cell.setCellValue( tableManagers_row_N_cell_1.get( i ) );
        
                XSSFCellStyle cellStyle = workbook.createCellStyle();
        
                cellStyle.setFont(font_Normal);

                cellStyle.setIndention( indient );

                cellStyle.setVerticalAlignment( VerticalAlignment.CENTER );
        
                cell.setCellStyle(cellStyle);

                if( tableManagers_row_N_cell_0.size() >= i && tableManagers_row_N_cell_0.get( i ) != -1 ){

                    XSSFDrawing drawing = sheet.createDrawingPatriarch();

                    XSSFClientAnchor anchor = helper.createClientAnchor();

                    anchor.setAnchorType( ClientAnchor.AnchorType.DONT_MOVE_DO_RESIZE );

                    anchor.setCol1(0);
                    anchor.setRow1(lastRowNum);
                    anchor.setCol2(0);
                    anchor.setRow2(lastRowNum);

                    anchor.setDx2( Units.toEMU(57)  );
                    anchor.setDy2( Units.toEMU(57) );

                    XSSFPicture pict = drawing.createPicture(anchor, tableManagers_row_N_cell_0.get( i ) );

                    pict.resize(scale, 1.0);

                }

                lastRowNum += 2;

            }

        }

        // Создаем строку с Наименованием позиции Соисполнителей
        
        if( !( tableWorkgroup_row_0_cell_0 == null || tableWorkgroup_row_0_cell_0.isEmpty() || tableWorkgroup_row_0_cell_0.trim().isEmpty() ) ){

            XSSFRow row  = sheet.createRow( lastRowNum );

            XSSFCell cell  = row.createCell(0);

            cell.setCellType(CellType.STRING);
    
            cell.setCellValue(tableWorkgroup_row_0_cell_0);
    
            XSSFCellStyle cellStyle  = workbook.createCellStyle();
    
            cellStyle.setFont(font_Bold);
    
            cell.setCellStyle(cellStyle);

            lastRowNum += 2;

        }

        // Создаем строки с Подписями и Инициалами Соисполнителей

        for (int i = 0; i < tableWorkgroup_row_N_cell_1.size(); i++) {

            String value = tableWorkgroup_row_N_cell_1.get( i );

            if( !( value == null || value.isEmpty() || value.trim().isEmpty() ) ){

                XSSFRow row  = sheet.createRow( lastRowNum );            
                
                if( rightColumn > 0 )
                    sheet.addMergedRegion( new CellRangeAddress(
                        lastRowNum, lastRowNum, 0, rightColumn
                ));
                
                row.setHeight( height );

                XSSFCell cell  = row.createCell(0);

                cell.setCellType(CellType.STRING);
        
                cell.setCellValue( value );
        
                XSSFCellStyle cellStyle = workbook.createCellStyle();
        
                cellStyle.setFont(font_Normal);

                cellStyle.setIndention( indient );

                cellStyle.setVerticalAlignment( VerticalAlignment.CENTER );
        
                cell.setCellStyle(cellStyle);

                if( tableWorkgroup_row_N_cell_0.size() >= i && tableWorkgroup_row_N_cell_0.get( i ) != -1  ){

                    XSSFDrawing drawing = sheet.createDrawingPatriarch();

                    XSSFClientAnchor anchor = helper.createClientAnchor();

                    anchor.setAnchorType( ClientAnchor.AnchorType.DONT_MOVE_DO_RESIZE );

                    anchor.setCol1(0);
                    anchor.setRow1(lastRowNum);
                    anchor.setCol2(0);
                    anchor.setRow2(lastRowNum);

                    anchor.setDx2( Units.toEMU(57)  );
                    anchor.setDy2( Units.toEMU(57) );

                    XSSFPicture pict = drawing.createPicture(anchor, tableWorkgroup_row_N_cell_0.get( i ) );

                    pict.resize(scale, 1.0);

                }

                lastRowNum += 2;

            }
        }

        // create footer

        Footer footerExcel = sheet.getFooter();

        footerExcel.setLeft( HSSFFooter.fontSize( (short) 8 ) + footerWord );

        // ---------------------------
        // Create result
        // ---------------------------

        ByteArrayOutputStream result = new ByteArrayOutputStream();

        workbook.write(result);

        result.close();

        // ---------------------------
        // return result
        // ---------------------------

        return result;

        // ---------------------------
        // End
        // ---------------------------

    }

}
