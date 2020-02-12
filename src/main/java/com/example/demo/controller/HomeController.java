package com.example.demo.controller;

import com.example.demo.entity.Article;
import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.repository.ClientRepository;
import com.example.demo.repository.FactureRepository;
import com.example.demo.repository.LigneFactureRepository;
import com.example.demo.service.impl.ClientServiceImpl;
import com.example.demo.service.ArticleService;
import com.example.demo.service.FactureService;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.*;
import java.time.LocalDate;
import java.time.Period;
import java.util.List;

/**
 * Controller principale pour affichage des clients / factures sur la page d'acceuil.
 */
@Controller
public class HomeController {

    @Autowired
    private ClientRepository clientRepository;

    @Autowired
    private FactureRepository factureRepository;

    @Autowired
    private LigneFactureRepository ligneFactureRepository;

    private ArticleService articleService;
    private ClientServiceImpl clientServiceImpl;
    private FactureService factureService;

    public HomeController(ArticleService articleService, ClientServiceImpl clientService, FactureService factureService) {
        this.articleService = articleService;
        this.clientServiceImpl = clientService;
        this.factureService = factureService;
    }

    @GetMapping("/")
    public ModelAndView home() {
        ModelAndView modelAndView = new ModelAndView("home");

        List<Article> articles = articleService.findAll();
        modelAndView.addObject("articles", articles);

        List<Client> toto = clientServiceImpl.findAllClients();
        modelAndView.addObject("clients", toto);

        List<Facture> factures = factureService.findAllFactures();
        modelAndView.addObject("factures", factures);

        return modelAndView;
    }


    // 1° export articles.csv avec la colonne description
    @GetMapping("/articles/csv")
    public void articlesCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachement; filename=\"export-articles.csv\"");
        PrintWriter writer = response.getWriter();

        List<Article> articles = articleService.findAll();

        writer.println("Libelle;Prix;Description");
        for (Article article : articles) {
            String line = article.getLibelle() + ";" + article.getPrix() + ";" + article.getDescription() + ";";
            writer.println(line);
        }
    }

    // Export articles au format xlsx
    private static String[] columns = { "Libelle", "Prix", "Description"};
    @GetMapping("/articles/xlsx")
    public void articlesXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachement; filename=\"articles-excel.xlsx\"");

        List<Article> articles = articleService.findAll();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Articles");

        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 13);
        headerFont.setColor(IndexedColors.BLUE.getIndex());

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setBorderBottom(BorderStyle.MEDIUM);
        headerCellStyle.setBottomBorderColor(IndexedColors.BLUE.getIndex());
        headerCellStyle.setBorderLeft(BorderStyle.MEDIUM);
        headerCellStyle.setLeftBorderColor(IndexedColors.BLUE.getIndex());
        headerCellStyle.setBorderRight(BorderStyle.MEDIUM);
        headerCellStyle.setRightBorderColor(IndexedColors.BLUE.getIndex());
        headerCellStyle.setBorderTop(BorderStyle.MEDIUM);
        headerCellStyle.setTopBorderColor(IndexedColors.BLUE.getIndex());
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        int rowNum = 1;

        for (Article article : articles) {

            Row row = sheet.createRow(rowNum++);

            Font FontData = workbook.createFont();
            FontData.setBold(false);
            FontData.setFontHeightInPoints((short) 12);
            FontData.setColor(IndexedColors.BLACK.getIndex());

            CellStyle CellStyleData = workbook.createCellStyle();
            CellStyleData.setBorderBottom(BorderStyle.DOUBLE);
            CellStyleData.setBottomBorderColor(IndexedColors.BLUE.getIndex());
            CellStyleData.setBorderLeft(BorderStyle.DOUBLE);
            CellStyleData.setLeftBorderColor(IndexedColors.BLUE.getIndex());
            CellStyleData.setBorderRight(BorderStyle.DOUBLE);
            CellStyleData.setRightBorderColor(IndexedColors.BLUE.getIndex());
            CellStyleData.setBorderTop(BorderStyle.DOUBLE);
            CellStyleData.setTopBorderColor(IndexedColors.BLUE.getIndex());
            CellStyleData.setFont(FontData);

            Cell cell_0 = row.createCell(0);
            Cell cell_1 = row.createCell(1);
            Cell cell_2 = row.createCell(2);
            row.createCell(0).setCellValue(article.getLibelle());
            row.createCell(1).setCellValue(article.getPrix());
            row.createCell(2).setCellValue(article.getDescription());
            cell_0.setCellStyle(CellStyleData);
            cell_1.setCellStyle(CellStyleData);
            cell_2.setCellStyle(CellStyleData);
        }

        // Resize all columns to fit the content size
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        ServletOutputStream out = response.getOutputStream();
        workbook.write(out);
        workbook.close();
        System.out.println("Le fichier articles-excel.xlsx a été correctement enregistré sur le disque !");

    }

    // 5° export PDF des articles
    @GetMapping("/articles/pdf")
    public void articlesPDF(HttpServletRequest request, HttpServletResponse response) throws IOException, DocumentException {
        response.setContentType("Content-Type: text/html; charset=UTF-8");
        response.setHeader("Content-Disposition", "attachement; filename=\"articles.pdf\"");

        List<Article> articles = articleService.findAll();

        // 1. Create document
        Document document = new Document(PageSize.A4, 50, 50, 50, 50);

        // 2. Create PdfWriter
        PdfWriter.getInstance(document, response.getOutputStream());

        // 3. Open document
        document.open();

        PdfPTable table = new PdfPTable(3); // 3 columns.
        table.setWidthPercentage(100); //Width 100%
        table.setSpacingBefore(10f); //Space before table
        table.setSpacingAfter(10f); //Space after table

        //Set Column widths
        float[] columnWidths = {2f, 1f, 2f};
        table.setWidths(columnWidths);

        // 4. Add content
        document.add(new Paragraph("Liste des articles"));

        PdfPCell cell1 = new PdfPCell(new Paragraph("Libellé"));
        cell1.setBorderColor(BaseColor.BLUE);
        cell1.setPaddingLeft(10);
        cell1.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell1.setVerticalAlignment(Element.ALIGN_MIDDLE);

        PdfPCell cell2 = new PdfPCell(new Paragraph("Prix"));
        cell2.setBorderColor(BaseColor.GREEN);
        cell2.setPaddingLeft(10);
        cell2.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell2.setVerticalAlignment(Element.ALIGN_MIDDLE);

        PdfPCell cell3 = new PdfPCell(new Paragraph("Description"));
        cell3.setBorderColor(BaseColor.RED);
        cell3.setPaddingLeft(10);
        cell3.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell3.setVerticalAlignment(Element.ALIGN_MIDDLE);

        table.addCell(cell1);
        table.addCell(cell2);
        table.addCell(cell3);

        for (Article article : articles) {
            table.addCell(article.getLibelle());
            table.addCell(String.valueOf(article.getPrix()));
            table.addCell(article.getDescription());
        }

        document.add(table);

        // 5. Close document
        document.close();
    }

    // 2° export client.csv => rajouter l'age
    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachement; filename=\"export-clients.csv\"");
        PrintWriter writer = response.getWriter();

        List<Client> clients = clientServiceImpl.findAllClients();

        writer.println("Nom;Prénom;Age");
        for (Client client : clients) {
            writer.write(client.getNom() + ";" + client.getPrenom() + ";" + client.getAge() + " ans" + "\n");
        }
    }


    // 3° export client.xlsx avec la mise en forme demandée
    private static String[] columnsClient = { "Nom", "Prénom", "Age"};
    @GetMapping("/clients/xlsx")
    public void clientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachement; filename=\"clients-excel.xlsx\"");
        //PrintWriter writer = response.getWriter();

        List<Client> clients = clientServiceImpl.findAllClients();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");


        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.PINK.getIndex());

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setBorderBottom(BorderStyle.THICK);
        headerCellStyle.setBottomBorderColor(IndexedColors.BLUE.getIndex());
        headerCellStyle.setBorderLeft(BorderStyle.THICK);
        headerCellStyle.setLeftBorderColor(IndexedColors.BLUE.getIndex());
        headerCellStyle.setBorderRight(BorderStyle.THICK);
        headerCellStyle.setRightBorderColor(IndexedColors.BLUE.getIndex());
        headerCellStyle.setBorderTop(BorderStyle.THICK);
        headerCellStyle.setTopBorderColor(IndexedColors.BLUE.getIndex());
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        for (int i = 0; i < columnsClient.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columnsClient[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Create Other rows and cells with contacts data
        int rowNum = 1;

        for (Client client : clients) {
            // Calcul de l'âge du client
            LocalDate birthday = client.getDateNaissance();
            LocalDate now = LocalDate.now();
            Period period = Period.between(birthday, now);

            Row row = sheet.createRow(rowNum++);
            Font FontData = workbook.createFont();
            FontData.setBold(false);
            FontData.setFontHeightInPoints((short) 12);
            FontData.setColor(IndexedColors.BLACK.getIndex());

            CellStyle CellStyleData = workbook.createCellStyle();
            CellStyleData.setBorderBottom(BorderStyle.THICK);
            CellStyleData.setBottomBorderColor(IndexedColors.BLUE.getIndex());
            CellStyleData.setBorderLeft(BorderStyle.THICK);
            CellStyleData.setLeftBorderColor(IndexedColors.BLUE.getIndex());
            CellStyleData.setBorderRight(BorderStyle.THICK);
            CellStyleData.setRightBorderColor(IndexedColors.BLUE.getIndex());
            CellStyleData.setBorderTop(BorderStyle.THICK);
            CellStyleData.setTopBorderColor(IndexedColors.BLUE.getIndex());
            CellStyleData.setFont(FontData);

            Cell cell_0 = row.createCell(0);
            Cell cell_1 = row.createCell(1);
            Cell cell_2 = row.createCell(2);
            row.createCell(0).setCellValue(client.getNom());
            row.createCell(1).setCellValue(client.getPrenom());
            row.createCell(2).setCellValue(period.getYears());
            cell_0.setCellStyle(CellStyleData);
            cell_1.setCellStyle(CellStyleData);
            cell_2.setCellStyle(CellStyleData);
        }

        // Resize all columns to fit the content size
        for (int i = 0; i < columnsClient.length; i++) {
            sheet.autoSizeColumn(i);
        }

        ServletOutputStream out = response.getOutputStream();
        workbook.write(out);

        workbook.close();
        System.out.println("Le fichier clients-excel.xlsx a été correctement enregistré sur le disque !");
    }


    // 5° export PDF Clients
    @GetMapping("/clients/pdf")
    public void clientsPDF(HttpServletRequest request, HttpServletResponse response) throws IOException, DocumentException {
            response.setContentType("Content-Type: text/html; charset=UTF-8");
            response.setHeader("Content-Disposition", "attachement; filename=\"clients.pdf\"");

            List<Client> clients = clientServiceImpl.findAllClients();

            // 1. Create document
            Document document = new Document(PageSize.A4, 50, 50, 50, 50);

            // 2. Create PdfWriter
            PdfWriter.getInstance(document, response.getOutputStream());

            // 3. Open document
            document.open();

            PdfPTable table = new PdfPTable(3); // 3 columns.
            table.setWidthPercentage(100); //Width 100%
            table.setSpacingBefore(10f); //Space before table
            table.setSpacingAfter(10f); //Space after table

            //Set Column widths
            float[] columnWidths = {1f, 1f, 1f};
            table.setWidths(columnWidths);

            // 4. Add content
            document.add(new Paragraph("Liste des clients"));

            PdfPCell cell1 = new PdfPCell(new Paragraph("Nom"));
            cell1.setBorderColor(BaseColor.BLUE);
            cell1.setPaddingLeft(10);
            cell1.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell1.setVerticalAlignment(Element.ALIGN_MIDDLE);

            PdfPCell cell2 = new PdfPCell(new Paragraph("Prénom"));
            cell2.setBorderColor(BaseColor.GREEN);
            cell2.setPaddingLeft(10);
            cell2.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell2.setVerticalAlignment(Element.ALIGN_MIDDLE);

            PdfPCell cell3 = new PdfPCell(new Paragraph("Age"));
            cell3.setBorderColor(BaseColor.RED);
            cell3.setPaddingLeft(10);
            cell3.setHorizontalAlignment(Element.ALIGN_CENTER);
            cell3.setVerticalAlignment(Element.ALIGN_MIDDLE);

            table.addCell(cell1);
            table.addCell(cell2);
            table.addCell(cell3);

            for (Client client : clients) {
                // Calcul de l'âge du client
                LocalDate birthday = client.getDateNaissance();
                LocalDate now = LocalDate.now();
                Period period = Period.between(birthday, now);

                table.addCell(client.getNom());
                table.addCell(client.getPrenom());
                table.addCell(String.valueOf(period.getYears()));
            }

            document.add(table);

            // 5. Close document
            document.close();
    }

    // 4° export total factures
    private static String[] columnsFacture = { "Nom", "Prénom", "Année de naissance", "Numéro Facture"};
    @GetMapping("/factures/xlsx")
    public void facturesXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachement; filename=\"factures-excel.xlsx\"");
        //PrintWriter writer = response.getWriter();

        //List<Facture> factures = factureService.findAllFactures();
        List<Facture> factures = factureService.findFacturesNom();



        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Factures");


        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.DARK_BLUE.getIndex());

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setBorderBottom(BorderStyle.DOTTED);
        headerCellStyle.setBottomBorderColor(IndexedColors.GREEN.getIndex());
        headerCellStyle.setBorderLeft(BorderStyle.DOTTED);
        headerCellStyle.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        headerCellStyle.setBorderRight(BorderStyle.DOTTED);
        headerCellStyle.setRightBorderColor(IndexedColors.GREEN.getIndex());
        headerCellStyle.setBorderTop(BorderStyle.DOTTED);
        headerCellStyle.setTopBorderColor(IndexedColors.GREEN.getIndex());
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        for (int i = 0; i < columnsFacture.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columnsFacture[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Create Other rows and cells with contacts data
        int rowNum = 1;

        for (Facture facture : factures) {

            Row row = sheet.createRow(rowNum++);
            Font FontData = workbook.createFont();
            FontData.setBold(false);
            FontData.setFontHeightInPoints((short) 12);
            FontData.setColor(IndexedColors.BLACK.getIndex());

            CellStyle CellStyleData = workbook.createCellStyle();
            CellStyle CellStyleDate = workbook.createCellStyle();

            CellStyleData.setFont(FontData);
            CellStyleDate.setFont(FontData);

            /*
            for (int i = 0; i < 3; i++) {
                Cell cell = row.createCell(i);
                cell.setCellStyle(headerCellStyle);
            }
            */


            CreationHelper createHelper = workbook.getCreationHelper();
            CellStyleDate.setDataFormat(
                    createHelper.createDataFormat().getFormat("m/d/yy"));

            Cell cell_0 = row.createCell(0);
            Cell cell_1 = row.createCell(1);
            Cell cell_2 = row.createCell(2);
            Cell cell_3 = row.createCell(3);
            row.createCell(0).setCellValue(facture.getClient().getNom());
            row.createCell(1).setCellValue(facture.getClient().getPrenom());
            // setCellValue ne fonctionne pas avec localdate
            java.util.Date date = java.sql.Date.valueOf(facture.getClient().getDateNaissance());
            row.createCell(2).setCellValue(date);
            row.createCell(3).setCellValue(facture.getId());
            cell_0.setCellStyle(CellStyleData);
            cell_1.setCellStyle(CellStyleData);
            cell_2.setCellStyle(CellStyleDate);
            cell_3.setCellStyle(CellStyleData);
        }

        // Resize all columns to fit the content size
        for (int i = 0; i < columnsFacture.length; i++) {
            sheet.autoSizeColumn(i);
        }

        ServletOutputStream out = response.getOutputStream();
        workbook.write(out);

        //FileOutputStream outputStream = new FileOutputStream("clients-excel.xlsx");
        //workbook.write(outputStream);
        workbook.close();
        System.out.println("Le fichier factures-excel.xlsx a été correctement enregistré sur le disque !");

    }

    // 5° Factures Clients PDF
    @GetMapping("/factures/{id}/pdf")
    public void factureClientPDF(@PathVariable("id") Long id, HttpServletRequest request, HttpServletResponse response) throws IOException, DocumentException
    {
        response.setContentType("Content-Type: text/html; charset=UTF-8");
        response.setHeader("Content-Disposition", "attachement; filename=\"facture-"+id+".pdf\"");


        //List<Facture> factures = factureService.findFacturesNom();
        List<Article> articles = articleService.findAll();
        Facture factures = factureRepository.findById(id).get();
        //Client clients = clientRepository.findById(id).get();
        //List<LigneFacture> ligneFacture = factures.getLigneFactures();


        // 1. Create document
        Document document = new Document(PageSize.A4, 50, 50, 50, 50);

        // 2. Create PdfWriter
        PdfWriter.getInstance(document, response.getOutputStream());

        // 3. Open document
        document.open();

        PdfPTable table = new PdfPTable(3); // 3 columns.
        table.setWidthPercentage(100); //Width 100%
        table.setSpacingBefore(10f); //Space before table
        table.setSpacingAfter(10f); //Space after table

        //Set Column widths
        float[] columnWidths = {1f, 1f, 1f};
        table.setWidths(columnWidths);

        // 4. Add content
        document.add(new Paragraph("Facture numéro " + id + " de " + factures.getClient().getNom() + " " + factures.getClient().getPrenom()));

        PdfPCell cell1 = new PdfPCell(new Paragraph("Désignation"));
        cell1.setBorderColor(BaseColor.BLUE);
        cell1.setPaddingLeft(10);
        cell1.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell1.setVerticalAlignment(Element.ALIGN_MIDDLE);

        PdfPCell cell2 = new PdfPCell(new Paragraph("Quantité"));
        cell2.setBorderColor(BaseColor.GREEN);
        cell2.setPaddingLeft(10);
        cell2.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell2.setVerticalAlignment(Element.ALIGN_MIDDLE);

        PdfPCell cell3 = new PdfPCell(new Paragraph("Prix unitaire"));
        cell3.setBorderColor(BaseColor.RED);
        cell3.setPaddingLeft(10);
        cell3.setHorizontalAlignment(Element.ALIGN_CENTER);
        cell3.setVerticalAlignment(Element.ALIGN_MIDDLE);

        table.addCell(cell1);
        table.addCell(cell2);
        table.addCell(cell3);

        /*
        for (Facture facture : factures) {
            table.addCell(factures.getClient().getClass().);
            table.addCell("");
            table.addCell("zz");
        }

         */

        document.add(table);

        // 5. Close document
        document.close();
    }

    // Factures Client XLSX
    @GetMapping("/clients/{id}/factures/xlsx")
    public void facturesClientXLSX(@PathVariable("id") Long id, HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.ms-excel");
        response.setHeader("Content-Disposition", "attachement; filename=\"factures-client-"+id+"-excel.xlsx\"");
        //PrintWriter writer = response.getWriter();

        List<Facture> factures = factureService.findFacturesNom();
        Client clients = clientRepository.findById(id).get();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(clients.getNom()+" "+clients.getPrenom());

        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);

        CellStyle factureCellStyle = workbook.createCellStyle();
        CellStyle CellStyleDate = workbook.createCellStyle();
        factureCellStyle.setFont(headerFont);

        CreationHelper createHelper = workbook.getCreationHelper();
        CellStyleDate.setDataFormat(
                createHelper.createDataFormat().getFormat("yyyy"));

        Row headerRow = sheet.createRow(0);

        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
        }

        // ROW Nom
        Row row = sheet.createRow(0);
        Cell cell0 = row.createCell(0);
        cell0.setCellValue("Nom : ");
        Cell cell1 = row.createCell(1);
        cell1.setCellValue(clients.getNom());

        // ROW Prénom
        Row row1 = sheet.createRow(1);
        Cell cell3 = row1.createCell(0);
        cell3.setCellValue("Prénom : ");
        Cell cell4 = row1.createCell(1);
        cell4.setCellValue(clients.getPrenom());

        // ROW Année de naissance
        Row row2 = sheet.createRow(2);
        Cell cell5 = row2.createCell(0);
        cell5.setCellValue("Année de naissance : ");
        Cell cell6 = row2.createCell(1);
        // setCellValue ne fonctionne pas avec localdate
        java.util.Date date = java.sql.Date.valueOf(clients.getDateNaissance());
        cell6.setCellValue(date);
        cell6.setCellStyle(CellStyleDate);

        // ROW Facturation
        Row row3 = sheet.createRow(3);
        for (Facture facture : factures) {
            Cell cell7 = row3.createCell(0);
            cell7.setCellValue("2 facture(s) : ");
            cell7.setCellStyle(factureCellStyle);
        }
        Cell cell8 = row3.createCell(1);
        cell8.setCellValue("");
        cell8.setCellStyle(CellStyleDate);

        // Resize all columns to fit the content size
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        ServletOutputStream out = response.getOutputStream();
        workbook.write(out);

        workbook.close();
    }

}
