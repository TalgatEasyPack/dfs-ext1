package com.dfs.ext1;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.net.InetSocketAddress;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Iterator;
import java.util.List;

import com.google.gson.Gson;
import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpHandler;
import com.sun.net.httpserver.HttpServer;

import org.apache.maven.model.Model;
import org.apache.maven.model.io.xpp3.MavenXpp3Reader;
import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Footer;
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
/**
 * Hello world!
 */

public final class App {
    private App() {
    }

    /**
     * Says hello to the world.
     * @param args The arguments of the program.
     */
    public static void main(String[] args) throws Exception {
        
        HttpServer server = HttpServer.create(new InetSocketAddress(8091), 0);
        server.createContext("/", new PingHandler());
        server.createContext("/version", new VersionHandler());
        server.createContext("/acts", new MyHandlerActs());
        server.createContext("/actsSvod", new MyHandlerActsSvod());
        server.setExecutor(null); // creates a default executor
        server.start();

    }

    static class PingHandler implements HttpHandler{
        public void handle(HttpExchange t) throws IOException {
            try {

                // ---------------------------

                String response = "Service is ready";

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
            }
        }
    }

    static class VersionHandler implements HttpHandler{
        public void handle(HttpExchange t) throws IOException {
            try {

                // ---------------------------

                MavenXpp3Reader reader = new MavenXpp3Reader();
                Model model;
                if ((new File("pom.xml")).exists())
                model = reader.read(new FileReader("pom.xml"));
                else
                model = reader.read(
                    new InputStreamReader(
                        App.class.getResourceAsStream(
                        "/META-INF/maven/de.scrum-master.stackoverflow/aspectj-introduce-method/pom.xml"
                    )
                    )
                );

                String response = model.getId();

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
            }
        }
    }

    static class MyJsonClass {
        public int id;
        public String base64_excel;
        public String base64_word;
    }

    static class MyHandlerActs implements HttpHandler {
        @Override
        public void handle(HttpExchange t) throws IOException {
            try {

                // ---------------------------
                // Read content
                // ---------------------------

                Gson g = new Gson();

                MyJsonClass myJsonClass = g.fromJson(new InputStreamReader(t.getRequestBody(), "UTF-8"),
                        MyJsonClass.class);

                // ---------------------------
                // Decode content and open Stream
                // ---------------------------

                byte[] decoded_excel = Base64.getDecoder().decode(myJsonClass.base64_excel);

                InputStream is_excel = new ByteArrayInputStream(decoded_excel);

                byte[] decoded_word = Base64.getDecoder().decode(myJsonClass.base64_word);

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

                myJsonClass = new MyJsonClass();

                myJsonClass.base64_excel = Base64.getEncoder().encodeToString(bos.toByteArray());

                // ---------------------------
                // Convert json to responce
                // ---------------------------

                // String response = "This is the response";

                String response = g.toJson(myJsonClass);

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
            }
        }

        public ByteArrayOutputStream mergeExcelAndWord(XSSFWorkbook workbook, XWPFDocument document)
                throws IOException {

            // ---------------------------
            // extract word
            // ---------------------------

            String docHead = "";
            XSSFFont fontHead = workbook.createFont();

            List<String> docText = new ArrayList<String>();
            XSSFFont fontText = workbook.createFont();

            List<Integer> wbImage = new ArrayList<Integer>();

            Iterator<IBodyElement> docElementsIterator = document.getBodyElementsIterator();

            while (docElementsIterator.hasNext()) {

                IBodyElement docElement = docElementsIterator.next();

                if ("TABLE".equalsIgnoreCase(docElement.getElementType().name())) {

                    List<XWPFTable> xwpfTableList = docElement.getBody().getTables();

                    XWPFTable xwpfTable = xwpfTableList.get(0);

                    // Table head
                    XWPFTableCell cellHead = xwpfTable.getRow(0).getCell(0);

                    for (XWPFParagraph p : cellHead.getParagraphs()) {

                        for (XWPFRun run : p.getRuns()) {

                            fontHead.setFontName( run.getFontName() );
                            fontHead.setBold( true );

                            fontText.setFontName( run.getFontName() );

                        }

                    }

                    docHead = cellHead.getText();

                    // Table body

                    for (int i = 1; i < xwpfTable.getRows().size(); i++) {

                        XWPFTableRow row = xwpfTable.getRow(i);

                        // QR-code

                        XWPFTableCell cell = row.getCell(0);

                        for (XWPFParagraph p : cell.getParagraphs()) {

                            for (XWPFRun run : p.getRuns()) {

                                for (XWPFPicture pic : run.getEmbeddedPictures()) {

                                    byte[] pictureData = pic.getPictureData().getData();

                                    wbImage.add( workbook.addPicture( pictureData , Workbook.PICTURE_TYPE_PNG) );

                                }

                            }

                        }

                        // User initals

                        String userInitials = row.getCell(1).getText();

                        docText.add(userInitials);
                        
                        //end

                    }

                }

            }
            
            String footerWord = document.getFooterArray(0).getText();

            // ---------------------------
            // Create result
            // ---------------------------
            int numberOfSheets = workbook.getNumberOfSheets();
            
            XSSFCreationHelper helper = workbook.getCreationHelper();

            Short height = 57 * 20;

            Short indient = 9;

            CellStyle cellStyleHead = workbook.createCellStyle();
            cellStyleHead.setFont(fontHead);

            CellStyle cellStyle = workbook.createCellStyle();

            cellStyle.setIndention( indient );
            cellStyle.setVerticalAlignment( VerticalAlignment.CENTER );
            cellStyle.setWrapText( true );
            cellStyle.setFont(fontText);

            for (int i = 0; i < numberOfSheets; i++) {

                // Open sheet

                XSSFSheet sheet = workbook.getSheetAt(i);

                int lastRowNum = sheet.getLastRowNum() + 1;

                // select last visible column in print area

                int rightColumn = 0;
                int cellWidth = sheet.getColumnWidth( rightColumn );

                while( cellWidth < 10240 ){

                    rightColumn++;

                    cellWidth += sheet.getColumnWidth( rightColumn ) ;
                    
                }

                // Create head row

                XSSFRow row = sheet.createRow(lastRowNum);

                XSSFCell cell = row.createCell(0);

                cell.setCellStyle( cellStyleHead );

                cell.setCellType(CellType.STRING);

                cell.setCellValue(docHead);

                // set empty row

                lastRowNum++;

                sheet.createRow(lastRowNum);

                lastRowNum++;

                // Create user's row's

                XSSFDrawing drawing = sheet.createDrawingPatriarch();

                Double scale = 1.0;
                if( sheet.getColumnWidth(0) == 2048 )
                    scale = 0.9;

                for (int j = 0; j < docText.size(); j++) {

                    int newRowNum = lastRowNum + j;
                    
                    // merge
                    if( rightColumn > 0 )
                        sheet.addMergedRegion( new CellRangeAddress(
                            newRowNum, newRowNum, 0, rightColumn
                        ));

                    // initial's 
                    row = sheet.createRow(newRowNum);

                    row.setHeight(height);

                    cell = row.createCell(0);

                    cell.setCellStyle( cellStyle );

                    cell.setCellType(CellType.STRING);

                    cell.setCellValue( docText.get(j) );

                    // qr-code's

                    XSSFClientAnchor anchor = helper.createClientAnchor();

                    anchor.setAnchorType( ClientAnchor.AnchorType.DONT_MOVE_DO_RESIZE );

                    anchor.setCol1(0);
                    anchor.setRow1(newRowNum);
                    anchor.setCol2(0);
                    anchor.setRow2(newRowNum);
                    
                    anchor.setDx2( Units.toEMU(57) );
                    anchor.setDy2( Units.toEMU(57) );

                    XSSFPicture pict = drawing.createPicture(anchor, wbImage.get(j));

                    pict.resize(scale, 1.0);
                    //pict.resize(1.0, 1.0);

                    sheet.createRow( newRowNum + 1 );

                    lastRowNum++;

                }

                //footer
    
                Footer footerExcel = sheet.getFooter();
    
                footerExcel.setLeft( HSSFFooter.fontSize( (short) 8 ) + footerWord );
    
                //end

            }

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
    
        public static ByteArrayOutputStream mergeExcelAndWord_ActsSvod(XSSFWorkbook workbook, XWPFDocument document)
        throws IOException {
    
            // ---------------------------
            // extract word
            // ---------------------------
    
            List<IBodyElement> bodyElements = document.getBodyElements();
    
            XWPFTable headTable = (XWPFTable) bodyElements.get(0);
    
            String fontName                 = headTable.getRow( 0 ).getCell( 0 ).getParagraphArray(0).getRuns().get(0).getFontName();
    
            String headTable_row_0_cell_0   = "\n";
    
            for (XWPFParagraph p : headTable.getRow( 0 ).getCell( 0 ).getParagraphs()) {
    
                for (XWPFRun run : p.getRuns()) {
    
                        headTable_row_0_cell_0 += run.getText(0);
    
                }
    
                headTable_row_0_cell_0 += "\n";
    
            }
    
            String headTable_row_1_cell_0   = "";
    
            for (XWPFParagraph p : headTable.getRow( 1 ).getCell( 0 ).getParagraphs()) {
    
                for (XWPFRun run : p.getRuns()) {
    
                    headTable_row_1_cell_0 += run.getText(0);
    
                    headTable_row_1_cell_0 += "\n";
    
                }
    
            }
    
            int    headTable_row_1_cell_1 = -1;
            headTable_row_1_cell_1 = workbook.addPicture( headTable.getRow( 1 ).getCell( 1 ).
                                                                    getParagraphArray(0).getRuns().get(0).
                                                                    getEmbeddedPictures().get(0).getPictureData().getData(), 
                                                                    Workbook.PICTURE_TYPE_PNG) ;
    
            XWPFTable footerTable = (XWPFTable) bodyElements.get(2);
    
            String footerTable_row_0_cell_0 = footerTable.getRow( 0 ).getCell( 0 ).getText();
    
            int    footerTable_row_1_cell_0 = -1;
            footerTable_row_1_cell_0 = workbook.addPicture( footerTable.getRow( 1 ).getCell( 0 ).
                                                                    getParagraphArray(0).getRuns().get(0).
                                                                    getEmbeddedPictures().get(0).getPictureData().getData(), 
                                                                    Workbook.PICTURE_TYPE_PNG) ;
            String footerTable_row_1_cell_1 = footerTable.getRow( 1 ).getCell( 1 ).getText();
    
            String footerTable_row_2_cell_0 = footerTable.getRow( 2 ).getCell( 0 ).getText();
            List<Integer> footerTable_row_N_cell_0  = new ArrayList<Integer>();
            List<String> footerTable_row_N_cell_1   = new ArrayList<String>();
    
            for (int i = 3; i < footerTable.getNumberOfRows() ; i++) {
                
                footerTable_row_N_cell_0.add(
                    workbook.addPicture( footerTable.getRow( i ).getCell( 0 ).
                    getParagraphArray(0).getRuns().get(0).
                    getEmbeddedPictures().get(0).getPictureData().getData(), 
                    Workbook.PICTURE_TYPE_PNG)
                );
    
                footerTable_row_N_cell_1.add(
                    footerTable.getRow( i ).getCell( 1 ).getText()
                );
    
            }
    
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
    
            // move down rows add 2 position
    
            sheet.shiftRows(sheet.getFirstRowNum(), sheet.getLastRowNum(), 2 );
    
            // create row #1 
    
            if ( !( headTable_row_0_cell_0 == null || headTable_row_0_cell_0.isEmpty() || headTable_row_0_cell_0.trim().isEmpty() ) ){
    
                XSSFRow row  = sheet.createRow(0);
    
                XSSFCell cell  = row.createCell(0);
    
                cell.setCellType(CellType.STRING);
        
                cell.setCellValue(headTable_row_0_cell_0);
        
                XSSFCellStyle cellStyle  = workbook.createCellStyle();
        
                cellStyle.setFont(font_Normal);
                
                cellStyle.setWrapText(true);
        
                cellStyle.setAlignment( HorizontalAlignment.RIGHT );
        
                cell.setCellStyle(cellStyle);
    
                if( rightColumn > 0 )
                    sheet.addMergedRegion( new CellRangeAddress(
                        0, 0, 0, rightColumnHead
                    ));
    
                int numberOfLines = headTable_row_0_cell_0.split("\n").length + 1;
    
                row.setHeightInPoints(numberOfLines*sheet.getDefaultRowHeightInPoints());
            }        
            
            // create row #2
    
            if ( !( headTable_row_1_cell_0 == null || headTable_row_1_cell_0.isEmpty() || headTable_row_1_cell_0.trim().isEmpty() ) ){
    
                XSSFRow row  = sheet.createRow(1);
    
                if( rightColumn > 0 )
                    sheet.addMergedRegion( new CellRangeAddress(
                        1, 1, 0, rightColumnHead
                    ));
                    
                row.setHeight( height );
    
                XSSFCell cell  = row.createCell(0);
    
                cell.setCellType(CellType.STRING);
        
                cell.setCellValue(headTable_row_1_cell_0);
        
                XSSFCellStyle cellStyle  = workbook.createCellStyle();
        
                cellStyle.setFont(font_Normal);
                
                cellStyle.setWrapText(true);
        
                cellStyle.setAlignment( HorizontalAlignment.RIGHT );
    
                cellStyle.setIndention( indient );
        
                cell.setCellStyle( cellStyle );
    
                if( headTable_row_1_cell_1 >= 0 ){
    
                    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    
                    XSSFClientAnchor anchor = helper.createClientAnchor();
    
                    anchor.setAnchorType( ClientAnchor.AnchorType.DONT_MOVE_DO_RESIZE );
    
                    anchor.setCol1( rightColumnHead+1 );
                    anchor.setRow1(1);
                    anchor.setCol2( rightColumnHead+1 );
                    anchor.setRow2(1);
    
                    anchor.setDx1( sheet.getColumnWidth( rightColumnHead+1 ) - Units.toEMU(57) );
                    anchor.setDy1( Units.toEMU( 0 ) );
    
                    anchor.setDx2( sheet.getColumnWidth( rightColumnHead+1 )  );
                    anchor.setDy2( Units.toEMU(57) );
    
                    drawing.createPicture(anchor, headTable_row_1_cell_1);
    
                }
    
            }
    
            // create row #3 + lastrow
    
            int lastRowNum = sheet.getLastRowNum() + 1;
            
            if( !( footerTable_row_0_cell_0 == null || footerTable_row_0_cell_0.isEmpty() || footerTable_row_0_cell_0.trim().isEmpty() ) ){
    
                XSSFRow row  = sheet.createRow( lastRowNum );
    
                XSSFCell cell  = row.createCell(0);
    
                cell.setCellType(CellType.STRING);
        
                cell.setCellValue(footerTable_row_0_cell_0);
        
                XSSFCellStyle cellStyle  = workbook.createCellStyle();
        
                cellStyle.setFont(font_Bold);
        
                cell.setCellStyle(cellStyle);
    
                lastRowNum += 2;
    
            }
    
            // create row #4 + lastrow
    
            if( !( footerTable_row_1_cell_1 == null || footerTable_row_1_cell_1.isEmpty() || footerTable_row_1_cell_1.trim().isEmpty() ) ){
    
                XSSFRow row  = sheet.createRow( lastRowNum );            
                
                if( rightColumn > 0 )
                    sheet.addMergedRegion( new CellRangeAddress(
                        lastRowNum, lastRowNum, 0, rightColumn
                ));
                
                row.setHeight( height );
    
                XSSFCell cell  = row.createCell(0);
    
                cell.setCellType(CellType.STRING);
        
                cell.setCellValue(footerTable_row_1_cell_1);
        
                XSSFCellStyle cellStyle = workbook.createCellStyle();
        
                cellStyle.setFont(font_Normal);
    
                cellStyle.setIndention( indient );
    
                cellStyle.setVerticalAlignment( VerticalAlignment.CENTER );
        
                cell.setCellStyle(cellStyle);
    
                if( headTable_row_1_cell_1 >= 0 ){
    
                    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    
                    XSSFClientAnchor anchor = helper.createClientAnchor();
    
                    anchor.setAnchorType( ClientAnchor.AnchorType.DONT_MOVE_DO_RESIZE );
    
                    anchor.setCol1(0);
                    anchor.setRow1(lastRowNum);
                    anchor.setCol2(0);
                    anchor.setRow2(lastRowNum);
    
                    anchor.setDx2( Units.toEMU(57)  );
                    anchor.setDy2( Units.toEMU(57) );
    
                    XSSFPicture pict = drawing.createPicture(anchor, footerTable_row_1_cell_0);
    
                    pict.resize(scale, 1.0);
    
                }
    
                lastRowNum += 2;
    
            }
            
            // create row #5 + lastrow
            
            if( !( footerTable_row_2_cell_0 == null || footerTable_row_2_cell_0.isEmpty() || footerTable_row_2_cell_0.trim().isEmpty() ) ){
    
                XSSFRow row  = sheet.createRow( lastRowNum );
    
                XSSFCell cell  = row.createCell(0);
    
                cell.setCellType(CellType.STRING);
        
                cell.setCellValue(footerTable_row_2_cell_0);
        
                XSSFCellStyle cellStyle  = workbook.createCellStyle();
        
                cellStyle.setFont(font_Bold);
        
                cell.setCellStyle(cellStyle);
    
                lastRowNum += 2;
    
            }
    
            // create row #N + lastrow
    
            for (int i = 0; i < footerTable_row_N_cell_1.size(); i++) {
                
                XSSFRow row  = sheet.createRow( lastRowNum );            
                
                if( rightColumn > 0 )
                    sheet.addMergedRegion( new CellRangeAddress(
                        lastRowNum, lastRowNum, 0, rightColumn
                ));
                
                row.setHeight( height );
    
                XSSFCell cell  = row.createCell(0);
    
                cell.setCellType(CellType.STRING);
        
                cell.setCellValue( footerTable_row_N_cell_1.get( i ) );
        
                XSSFCellStyle cellStyle = workbook.createCellStyle();
        
                cellStyle.setFont(font_Normal);
    
                cellStyle.setIndention( indient );
    
                cellStyle.setVerticalAlignment( VerticalAlignment.CENTER );
        
                cell.setCellStyle(cellStyle);
    
                if( footerTable_row_N_cell_0.size() >= i ){
    
                    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    
                    XSSFClientAnchor anchor = helper.createClientAnchor();
    
                    anchor.setAnchorType( ClientAnchor.AnchorType.DONT_MOVE_DO_RESIZE );
    
                    anchor.setCol1(0);
                    anchor.setRow1(lastRowNum);
                    anchor.setCol2(0);
                    anchor.setRow2(lastRowNum);
    
                    anchor.setDx2( Units.toEMU(57)  );
                    anchor.setDy2( Units.toEMU(57) );
    
                    XSSFPicture pict = drawing.createPicture(anchor, footerTable_row_N_cell_0.get( i ) );
    
                    pict.resize(scale, 1.0);
    
                }
    
                lastRowNum += 2;
    
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

    static class MyHandlerActsSvod implements HttpHandler {
        @Override
        public void handle(HttpExchange t) throws IOException {
            try {

                // ---------------------------
                // Read content
                // ---------------------------

                Gson g = new Gson();

                MyJsonClass myJsonClass = g.fromJson(new InputStreamReader(t.getRequestBody(), "UTF-8"),
                        MyJsonClass.class);

                // ---------------------------
                // Decode content and open Stream
                // ---------------------------

                byte[] decoded_excel = Base64.getDecoder().decode(myJsonClass.base64_excel);

                InputStream is_excel = new ByteArrayInputStream(decoded_excel);

                byte[] decoded_word = Base64.getDecoder().decode(myJsonClass.base64_word);

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

                myJsonClass = new MyJsonClass();

                myJsonClass.base64_excel = Base64.getEncoder().encodeToString(bos.toByteArray());

                // ---------------------------
                // Convert json to responce
                // ---------------------------

                // String response = "This is the response";

                String response = g.toJson(myJsonClass);

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
            }
        }
    
        public static ByteArrayOutputStream mergeExcelAndWord(XSSFWorkbook workbook, XWPFDocument document)
        throws IOException {
    
            // ---------------------------
            // extract word
            // ---------------------------
    
            List<IBodyElement> bodyElements = document.getBodyElements();
    
            XWPFTable headTable = (XWPFTable) bodyElements.get(0);
    
            String fontName                 = headTable.getRow( 0 ).getCell( 0 ).getParagraphArray(0).getRuns().get(0).getFontName();
    
            String headTable_row_0_cell_0   = "\n";
    
            for (XWPFParagraph p : headTable.getRow( 0 ).getCell( 0 ).getParagraphs()) {
    
                for (XWPFRun run : p.getRuns()) {
    
                        headTable_row_0_cell_0 += run.getText(0);
    
                }
    
                headTable_row_0_cell_0 += "\n";
    
            }
    
            String headTable_row_1_cell_0   = headTable.getRow( 1 ).getCell( 0 ).getText();
    
            int    headTable_row_1_cell_1 = -1;
            headTable_row_1_cell_1 = workbook.addPicture( headTable.getRow( 1 ).getCell( 1 ).
                                                                    getParagraphArray(0).getRuns().get(0).
                                                                    getEmbeddedPictures().get(0).getPictureData().getData(), 
                                                                    Workbook.PICTURE_TYPE_PNG) ;
    
            XWPFTable footerTable = (XWPFTable) bodyElements.get(2);
    
            String footerTable_row_0_cell_0 = footerTable.getRow( 0 ).getCell( 0 ).getText();
    
            int    footerTable_row_1_cell_0 = -1;
            footerTable_row_1_cell_0 = workbook.addPicture( footerTable.getRow( 1 ).getCell( 0 ).
                                                                    getParagraphArray(0).getRuns().get(0).
                                                                    getEmbeddedPictures().get(0).getPictureData().getData(), 
                                                                    Workbook.PICTURE_TYPE_PNG) ;
            String footerTable_row_1_cell_1 = footerTable.getRow( 1 ).getCell( 1 ).getText();
    
            String footerTable_row_2_cell_0 = footerTable.getRow( 2 ).getCell( 0 ).getText();
            List<Integer> footerTable_row_N_cell_0  = new ArrayList<Integer>();
            List<String> footerTable_row_N_cell_1   = new ArrayList<String>();
    
            for (int i = 3; i < footerTable.getNumberOfRows() ; i++) {
                
                footerTable_row_N_cell_0.add(
                    workbook.addPicture( footerTable.getRow( i ).getCell( 0 ).
                    getParagraphArray(0).getRuns().get(0).
                    getEmbeddedPictures().get(0).getPictureData().getData(), 
                    Workbook.PICTURE_TYPE_PNG)
                );
    
                footerTable_row_N_cell_1.add(
                    footerTable.getRow( i ).getCell( 1 ).getText()
                );
    
            }
    
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
    
            // move down rows add 2 position
    
            sheet.shiftRows(sheet.getFirstRowNum(), sheet.getLastRowNum(), 2 );
    
            // create row #1 
    
            if ( !( headTable_row_0_cell_0 == null || headTable_row_0_cell_0.isEmpty() || headTable_row_0_cell_0.trim().isEmpty() ) ){
    
                XSSFRow row  = sheet.createRow(0);
    
                XSSFCell cell  = row.createCell(0);
    
                cell.setCellType(CellType.STRING);
        
                cell.setCellValue(headTable_row_0_cell_0);
        
                XSSFCellStyle cellStyle  = workbook.createCellStyle();
        
                cellStyle.setFont(font_Normal);
                
                cellStyle.setWrapText(true);
        
                cellStyle.setAlignment( HorizontalAlignment.RIGHT );
        
                cell.setCellStyle(cellStyle);
    
                if( rightColumn > 0 )
                    sheet.addMergedRegion( new CellRangeAddress(
                        0, 0, 0, rightColumnHead
                    ));
    
                int numberOfLines = headTable_row_0_cell_0.split("\n").length + 1;
    
                row.setHeightInPoints(numberOfLines*sheet.getDefaultRowHeightInPoints());
            }        
            
            // create row #2
    
            if ( !( headTable_row_1_cell_0 == null || headTable_row_1_cell_0.isEmpty() || headTable_row_1_cell_0.trim().isEmpty() ) ){
    
                XSSFRow row  = sheet.createRow(1);
    
                if( rightColumn > 0 )
                    sheet.addMergedRegion( new CellRangeAddress(
                        1, 1, 0, rightColumnHead
                    ));
                    
                row.setHeight( height );
    
                XSSFCell cell  = row.createCell(0);
    
                cell.setCellType(CellType.STRING);
        
                cell.setCellValue(headTable_row_1_cell_0);
        
                XSSFCellStyle cellStyle  = workbook.createCellStyle();
        
                cellStyle.setFont(font_Normal);
                
                cellStyle.setWrapText(true);
        
                cellStyle.setAlignment( HorizontalAlignment.RIGHT );
                
                cellStyle.setVerticalAlignment( VerticalAlignment.CENTER );
    
                cellStyle.setIndention( indient );
        
                cell.setCellStyle( cellStyle );
    
                if( headTable_row_1_cell_1 >= 0 ){
    
                    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    
                    XSSFClientAnchor anchor = helper.createClientAnchor();
    
                    anchor.setAnchorType( ClientAnchor.AnchorType.DONT_MOVE_DO_RESIZE );
    
                    anchor.setCol1( rightColumnHead+1 );
                    anchor.setRow1(1);
                    anchor.setCol2( rightColumnHead+1 );
                    anchor.setRow2(1);
    
                    anchor.setDx1( sheet.getColumnWidth( rightColumnHead+1 ) - Units.toEMU(57) );
                    anchor.setDy1( Units.toEMU( 0 ) );
    
                    anchor.setDx2( sheet.getColumnWidth( rightColumnHead+1 )  );
                    anchor.setDy2( Units.toEMU(57) );
    
                    drawing.createPicture(anchor, headTable_row_1_cell_1);
    
                }
    
            }
    
            // create row #3 + lastrow
    
            int lastRowNum = sheet.getLastRowNum() + 1;
            
            if( !( footerTable_row_0_cell_0 == null || footerTable_row_0_cell_0.isEmpty() || footerTable_row_0_cell_0.trim().isEmpty() ) ){
    
                XSSFRow row  = sheet.createRow( lastRowNum );
    
                XSSFCell cell  = row.createCell(0);
    
                cell.setCellType(CellType.STRING);
        
                cell.setCellValue(footerTable_row_0_cell_0);
        
                XSSFCellStyle cellStyle  = workbook.createCellStyle();
        
                cellStyle.setFont(font_Bold);
        
                cell.setCellStyle(cellStyle);
    
                lastRowNum += 2;
    
            }
    
            // create row #4 + lastrow
    
            if( !( footerTable_row_1_cell_1 == null || footerTable_row_1_cell_1.isEmpty() || footerTable_row_1_cell_1.trim().isEmpty() ) ){
    
                XSSFRow row  = sheet.createRow( lastRowNum );            
                
                if( rightColumn > 0 )
                    sheet.addMergedRegion( new CellRangeAddress(
                        lastRowNum, lastRowNum, 0, rightColumn
                ));
                
                row.setHeight( height );
    
                XSSFCell cell  = row.createCell(0);
    
                cell.setCellType(CellType.STRING);
        
                cell.setCellValue(footerTable_row_1_cell_1);
        
                XSSFCellStyle cellStyle = workbook.createCellStyle();
        
                cellStyle.setFont(font_Normal);
    
                cellStyle.setIndention( indient );
    
                cellStyle.setVerticalAlignment( VerticalAlignment.CENTER );
        
                cell.setCellStyle(cellStyle);
    
                if( headTable_row_1_cell_1 >= 0 ){
    
                    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    
                    XSSFClientAnchor anchor = helper.createClientAnchor();
    
                    anchor.setAnchorType( ClientAnchor.AnchorType.DONT_MOVE_DO_RESIZE );
    
                    anchor.setCol1(0);
                    anchor.setRow1(lastRowNum);
                    anchor.setCol2(0);
                    anchor.setRow2(lastRowNum);
    
                    anchor.setDx2( Units.toEMU(57)  );
                    anchor.setDy2( Units.toEMU(57) );
    
                    XSSFPicture pict = drawing.createPicture(anchor, footerTable_row_1_cell_0);
    
                    pict.resize(scale, 1.0);
    
                }
    
                lastRowNum += 2;
    
            }
            
            // create row #5 + lastrow
            
            if( !( footerTable_row_2_cell_0 == null || footerTable_row_2_cell_0.isEmpty() || footerTable_row_2_cell_0.trim().isEmpty() ) ){
    
                XSSFRow row  = sheet.createRow( lastRowNum );
    
                XSSFCell cell  = row.createCell(0);
    
                cell.setCellType(CellType.STRING);
        
                cell.setCellValue(footerTable_row_2_cell_0);
        
                XSSFCellStyle cellStyle  = workbook.createCellStyle();
        
                cellStyle.setFont(font_Bold);
        
                cell.setCellStyle(cellStyle);
    
                lastRowNum += 2;
    
            }
    
            // create row #N + lastrow
    
            for (int i = 0; i < footerTable_row_N_cell_1.size(); i++) {
                
                XSSFRow row  = sheet.createRow( lastRowNum );            
                
                if( rightColumn > 0 )
                    sheet.addMergedRegion( new CellRangeAddress(
                        lastRowNum, lastRowNum, 0, rightColumn
                ));
                
                row.setHeight( height );
    
                XSSFCell cell  = row.createCell(0);
    
                cell.setCellType(CellType.STRING);
        
                cell.setCellValue( footerTable_row_N_cell_1.get( i ) );
        
                XSSFCellStyle cellStyle = workbook.createCellStyle();
        
                cellStyle.setFont(font_Normal);
    
                cellStyle.setIndention( indient );
    
                cellStyle.setVerticalAlignment( VerticalAlignment.CENTER );
        
                cell.setCellStyle(cellStyle);
    
                if( footerTable_row_N_cell_0.size() >= i ){
    
                    XSSFDrawing drawing = sheet.createDrawingPatriarch();
    
                    XSSFClientAnchor anchor = helper.createClientAnchor();
    
                    anchor.setAnchorType( ClientAnchor.AnchorType.DONT_MOVE_DO_RESIZE );
    
                    anchor.setCol1(0);
                    anchor.setRow1(lastRowNum);
                    anchor.setCol2(0);
                    anchor.setRow2(lastRowNum);
    
                    anchor.setDx2( Units.toEMU(57)  );
                    anchor.setDy2( Units.toEMU(57) );
    
                    XSSFPicture pict = drawing.createPicture(anchor, footerTable_row_N_cell_0.get( i ) );
    
                    pict.resize(scale, 1.0);
    
                }
    
                lastRowNum += 2;
    
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

}

