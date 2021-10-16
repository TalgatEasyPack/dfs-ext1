package com.dfs.ext1.Hadlers;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Iterator;
import java.util.List;

import com.dfs.ext1.Base64Helper.RequestBodyJson;
import com.google.gson.Gson;
import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpHandler;

import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFCell;
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

public class MyHandlerActs implements HttpHandler {
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

}
