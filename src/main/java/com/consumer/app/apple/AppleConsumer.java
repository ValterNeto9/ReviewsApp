package com.consumer.app.apple;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;


public class AppleConsumer {

    private static Workbook workbook;
    private static int rowNum;

    private final static int AUTHOR_NAME = 0;
    private final static int RATING = 1;
    private final static int TITLE = 2;
    private final static int CONTENT = 3;
    private final static int VERSION = 4;
    private final static int UPDATED = 5;

    private static Elements allEntries = new Elements();
    private static Elements otherEntries = new Elements();
    
    public static void main(String[] args) throws Exception {
        getAndReadXml();
    }

    /**
     * Utilizando a biblioteca JSoup para acessar a URL e manipular os elementos do xml
     * @throws Exception
     */
    private static void getAndReadXml() throws Exception {
        
    	System.out.println("getAndReadXml");
        
        Document doc = Jsoup.connect("https://itunes.apple.com/br/rss/customerreviews/page=1/id=680819774/sortby=mostrecent/xml").get();
        
        Integer totalPages = getTotalPages(doc);
        
        //Adiciona as tags <entry> da primeira página na allEntries.
        unionEntries(doc.select("entry"));
        
        //Adiciona as tags <entry> das páginas seguintes na allEntries.
        for(int i = 2; i<= totalPages.intValue(); i++){  
        	doc = Jsoup.connect("https://itunes.apple.com/br/rss/customerreviews/page="+ i +"/id=680819774/sortby=mostrecent/xml").get();
        	unionEntries(doc.select("entry"));
        }                                                	
      
        //Adiciona as tags <entry> da primeira página na allEntries, considerando o país "US".
        doc = Jsoup.connect("https://itunes.apple.com/us/rss/customerreviews/page=1/id=680819774/sortby=mostrecent/xml").get();
        unionEntries(doc.select("entry"));
        
        initXls();

        Sheet sheet = workbook.getSheetAt(0);
        workbook.setSheetName(workbook.getSheetIndex(sheet), "Apple Store");
        
        for (Element entry : allEntries) {
			
        	String author = entry.getElementsByTag("author").select("name").text();
    		String rating = entry.getElementsByTag("im:rating").get(0).ownText();
    		String title = entry.getElementsByTag("title").get(0).ownText();
    		String content = entry.getElementsByTag("content").get(0).ownText();
    		String version = entry.getElementsByTag("im:version").get(0).ownText();
    		String updated = entry.getElementsByTag("updated").get(0).ownText();
    		
    		Row row = sheet.createRow(rowNum++);
    		Cell cell = row.createCell(AUTHOR_NAME);
    		cell.setCellValue(author);
    		
    		cell = row.createCell(RATING);
    		cell.setCellValue(rating);
    		
    		cell = row.createCell(TITLE);
    		cell.setCellValue(title);
    		
    		cell = row.createCell(CONTENT);
    		cell.setCellValue(content);
    		
    		cell = row.createCell(VERSION);
    		cell.setCellValue(version);
    		
    		cell = row.createCell(UPDATED);
    		cell.setCellValue(updated);
        	
		}
        
        FileOutputStream fileOut = new FileOutputStream("C:/TEMP/appreviews.xlsx");
        workbook.write(fileOut);
        workbook.close();
        fileOut.close();
        
        System.out.println("O Arquivo appreviews.xlsx foi criado em C:\\TEMP.");

    }

    /**
     * Adiciona as entries na allEntries  
     * @param entries
     */
	private static void unionEntries(Elements entries) {
		
        for (Element entry : entries) {
			if(!entry.select("author").isEmpty()){
				//System.out.println(entry.toString());
				allEntries.add(entry);
			}else{
				otherEntries.add(entry);
			}
		}
        
	}

	/**
	 * Verifica o total de páginas de resultados 
	 * @param doc
	 * @return
	 */
	private static Integer getTotalPages(Document doc) {
		
		Elements link = doc.getElementsByAttributeValueContaining("rel", "last");
        String url = link.get(0).attr("href");
         
		return Integer.valueOf(url.split("/id")[0].substring(48).substring(5));
	}
	
	 /**
     * Initializes the POI workbook and writes the header row
     */
    private static void initXls() {
        workbook = new XSSFWorkbook();
        
        CellStyle style = workbook.createCellStyle();
        Font boldFont = workbook.createFont();
        boldFont.setBold(true);
        style.setFont(boldFont);
        style.setAlignment(CellStyle.ALIGN_CENTER);

        Sheet sheet = workbook.createSheet();
        rowNum = 0;
        Row row = sheet.createRow(rowNum++);
        Cell cell = row.createCell(AUTHOR_NAME);
        cell.setCellValue("Author Name");
        cell.setCellStyle(style);

        cell = row.createCell(RATING);
        cell.setCellValue("Rating");
        cell.setCellStyle(style);

        cell = row.createCell(TITLE);
        cell.setCellValue("Title");
        cell.setCellStyle(style);

        cell = row.createCell(CONTENT);
        cell.setCellValue("Content");
        cell.setCellStyle(style);

        cell = row.createCell(VERSION);
        cell.setCellValue("Version");
        cell.setCellStyle(style);

        cell = row.createCell(UPDATED);
        cell.setCellValue("Updated at");
        cell.setCellStyle(style);

    }
	
}
