package com.dfs.ext1.Hadlers;

import java.io.IOException;
import java.io.OutputStream;

import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpHandler;

public class VersionHandler implements HttpHandler{

    private String version;

    public VersionHandler( String version ){

        this.version = version;

    }

    public void handle(HttpExchange t) throws IOException {
        try {

            // ---------------------------

            String response = version;

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
}
