package com.dfs.ext1;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.net.InetSocketAddress;

// import com.aspose.cells.Workbook;
// import com.aspose.pdf.SaveFormat;
import com.dfs.ext1.Hadlers.MyHandlerActs;
import com.dfs.ext1.Hadlers.MyHandlerActsSvod;
import com.dfs.ext1.Hadlers.MySignListForExcel;
import com.dfs.ext1.Hadlers.MySignListForExcel2;
import com.dfs.ext1.Hadlers.PingHandler;
import com.dfs.ext1.Hadlers.VersionHandler;
import com.sun.net.httpserver.HttpServer;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/**
 * Hello world!
 */

public final class App {

    private static String version = "1.0.006";

    private App() {
    }

    /**
     * Says hello to the world.
     * @param args The arguments of the program.
     */
    public static void main(String[] args) throws Exception {
        
        HttpServer server = HttpServer.create(new InetSocketAddress(8091), 0);
        server.createContext("/", new PingHandler());
        server.createContext("/version", new VersionHandler( version ));
        server.createContext("/acts", new MyHandlerActs());
        server.createContext("/actsSvod", new MyHandlerActsSvod());
        server.createContext("/signListForExcel", new MySignListForExcel2());
        server.setExecutor(null); // creates a default executor
        server.start();

        // LoadTestFile();

        // System.out.println("Выполнено слияние MySignListForExcel2");

    }

    public static void LoadTestFile() throws Exception{

        // ---------------------------

        FileInputStream is_excel = new FileInputStream(new File( "D:\\temp\\1.xlsx"));

        FileInputStream is_word = new FileInputStream(new File( "D:\\temp\\2.docx"));
            
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

        ByteArrayOutputStream bos = MySignListForExcel2.mergeExcelAndWord(workbook, document);

        FileOutputStream out_excel = new FileOutputStream( new File( "D:\\temp\\3.xlsx") );

        out_excel.write( bos.toByteArray() );

        out_excel.close();
        
        // ---------------------------

        //ByteArrayInputStream ios = new ByteArrayInputStream( bos.toByteArray() );

        // Workbook workbook2 = new Workbook( ios );

        // workbook2.save( "D:\\temp\\4.pdf", SaveFormat.Pdf );
        
        // ---------------------------

        bos.close();

        // ---------------------------
        // Close files
        // ---------------------------

        document.close();
        workbook.close();

        // ---------------------------

    }

}

