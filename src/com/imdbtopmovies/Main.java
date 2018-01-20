package com.imdbtopmovies;

import java.net.*;
import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.net.MalformedURLException;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

    private static String ExcelName = "IMDB Top 250.xlsx";
    private static String ExcelPath = "";
    private static String SheetName = "Top 250";
    private static String ImdbURL = "http://www.imdb.com/chart/top?ref_=nv_mv_250_6";

    public static void main(String[] args) {

        ExcelPath = System.getProperty("user.dir") + "\\";

        String ImdbHtml = GetHtmlFromUrl(ImdbURL);

        if (!ImdbHtml.isEmpty()) {

            String tbodyHtml = GetMovieListHtml(ImdbHtml);
            List<Movie> currentMovieList = GetMovieList(tbodyHtml);
            List<Movie> oldMovieList = new ArrayList<>();

            boolean isFileExist = CheckFileExists();
            if (isFileExist) {
                oldMovieList = ReadExcelFile();
            }

            if (oldMovieList.size() > 0 && currentMovieList.size() > 0) {
                List<Movie> newMovieList = CombineMovieLists(oldMovieList, currentMovieList);
                WriteExcelFile(newMovieList, isFileExist);
            } else {
                WriteExcelFile(currentMovieList, isFileExist);
            }
        }
    }

    private static String GetHtmlFromUrl(String url) {
        StringBuilder htmlText = new StringBuilder("");

        try {
            URL ImdbUrl = new URL(url);
            URLConnection urlCon = ImdbUrl.openConnection();
            urlCon.setRequestProperty("Accept-Language","en-US");

            BufferedReader in = new BufferedReader(new InputStreamReader(urlCon.getInputStream()));

            String html;
            while ((html = in.readLine()) != null) {
                htmlText.append(html);
            }
            in.close();

        } catch (MalformedURLException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return htmlText.toString();
    }

    private static String GetMovieListHtml(String imdbHtml) {
        String movieListPatternString = "(<tbody class=\"lister-list\">.*?</tbody)";
        Pattern movieListPattern = Pattern.compile(movieListPatternString);
        Matcher movieListMatcher = movieListPattern.matcher(imdbHtml);

        String tbodyHtml = "";
        while (movieListMatcher.find()) {
            tbodyHtml = tbodyHtml + movieListMatcher.group(1);
        }
        return tbodyHtml;
    }

    private static List<Movie> GetMovieList(String tbodyHtml) {
        String movieNamePatternString = "(<td class=\"titleColumn\">.*?<a href.*?>)(.*?)(</a>.*?<span class=\"secondaryInfo\">)(.*?)(</span>)";
        Pattern movieNamePattern = Pattern.compile(movieNamePatternString);
        Matcher movieNameMatcher = movieNamePattern.matcher(tbodyHtml);

        Movie[] movieList = new Movie[250];
        int index = 0;
        while (movieNameMatcher.find()) {
            movieList[index] = new Movie();
            movieList[index].Rank = Integer.toString(index + 1);
            movieList[index].Name = movieNameMatcher.group(2);
            movieList[index].Year = movieNameMatcher.group(4).replace("(", "").replace(")", "");
            index++;
        }
        return Arrays.asList(movieList);
    }

    private static boolean CheckFileExists() {
        File excelfile = new File(ExcelPath + ExcelName);
        if (excelfile.exists() && !excelfile.isDirectory()) {
            return true;
        }
        return false;
    }

    private static List<Movie> ReadExcelFile() {
        List<Movie> list = new ArrayList<>();

        try {

            FileInputStream excelFile = new FileInputStream(new File(ExcelPath + ExcelName));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                if (currentRow.getRowNum() == 0 || currentRow.getRowNum() == 1)
                    continue;

                Iterator<Cell> cellIterator = currentRow.iterator();
                Movie movie = new Movie("", "", "", "", "");
                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();

                    //if (currentCell.getCellTypeEnum() != CellType.BLANK && currentCell.getCellTypeEnum() != CellType._NONE) {
                    currentCell.setCellType(CellType.STRING);
                    int columnNumber = currentCell.getColumnIndex();
                    switch (columnNumber) {
                        case 0:
                            movie.Rank = currentCell.getStringCellValue().trim();
                            break;
                        case 1:
                            movie.OldRank = currentCell.getStringCellValue().trim();
                            break;
                        case 2:
                            movie.Name = currentCell.getStringCellValue().trim();
                            break;
                        case 3:
                            movie.Year = currentCell.getStringCellValue().trim();
                            break;
                        case 4:
                            movie.Status = currentCell.getStringCellValue().trim();
                            break;
                    }
                    //}
                }

                if (movie != null)
                    list.add(movie);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return list;
    }

    private static List<Movie> CombineMovieLists(List<Movie> oldMovieList, List<Movie> currentMovieList) {

        List<Movie> newlist = new LinkedList<>();

        for (Movie movie : currentMovieList) {

            Optional<Movie> filtered = oldMovieList.stream().filter(mov -> mov.Name.toLowerCase().replaceAll("\\h*$","").equals(movie.Name.toLowerCase()) && mov.Year.equals(movie.Year)).findFirst();

            if (filtered.isPresent()) {
                if (!filtered.get().Rank.equals(movie.Rank)) {
                    movie.OldRank = filtered.get().Rank;
                }
                else
                {
                    movie.OldRank = filtered.get().OldRank;
                }
                movie.Status = filtered.get().Status;
                newlist.add(movie);
                oldMovieList.removeIf(mov -> mov.Name.toLowerCase().replaceAll("\\h*$","").equals(movie.Name.toLowerCase()) && mov.Year.equals(movie.Year));
            } else {
                newlist.add(movie);
            }
        }

        for (Movie oldMovie : oldMovieList) {

            oldMovie.OldRank = oldMovie.Rank;
            oldMovie.Rank = "";
            newlist.add(oldMovie);

        }

        return newlist;
    }

    private static void WriteExcelFile(List<Movie> movieList, boolean isFileExist) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet(SheetName);

        // Top 250 Title Style
        XSSFCellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(new XSSFColor(new java.awt.Color(255, 192, 0)));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);

        Font font = workbook.createFont();
        font.setColor(IndexedColors.WHITE.getIndex());
        font.setFontHeightInPoints((short) 16);
        font.setBold(true);
        style.setFont(font);

        // Top 250 Title Row
        Row titleRow = sheet.createRow((short) 0);
        Cell cell = titleRow.createCell((short) 0);
        cell.setCellValue("Top 250");
        cell.setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));

        // Title Style
        XSSFCellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFillForegroundColor(new XSSFColor(new java.awt.Color(255, 192, 0)));
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleStyle.setAlignment(HorizontalAlignment.LEFT);

        Font titleFont = workbook.createFont();
        titleFont.setColor(IndexedColors.WHITE.getIndex());
        titleFont.setFontHeightInPoints((short) 14);
        titleStyle.setFont(titleFont);

        // Title Row
        titleRow = sheet.createRow((short) 1);
        CreateExcelColumn("Rank", titleRow, 0, titleStyle);
        CreateExcelColumn("OldRank", titleRow, 1, titleStyle);
        CreateExcelColumn("Name", titleRow, 2, titleStyle);
        CreateExcelColumn("Year", titleRow, 3, titleStyle);
        CreateExcelColumn("Status", titleRow, 4, titleStyle);
        sheet.setAutoFilter(CellRangeAddress.valueOf("A2:E2"));

        // Data Style
        XSSFCellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setAlignment(HorizontalAlignment.RIGHT);

        Row row;
        int rowNum = 2;
        for (Movie movie : movieList) {
            row = sheet.createRow(rowNum++);
            int colNum = 0;
            CreateExcelColumn(movie.Rank, row, colNum, dataStyle);
            CreateExcelColumn(movie.OldRank, row, ++colNum);
            CreateExcelColumn(movie.Name, row, ++colNum);
            CreateExcelColumn(movie.Year, row, ++colNum);
            CreateExcelColumn(movie.Status, row, ++colNum);
        }

        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);
        sheet.autoSizeColumn(4);

        try {
            FileOutputStream outputStream = new FileOutputStream(ExcelPath + ExcelName);
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    private static void CreateExcelColumn(Object field, Row row, int colNum) {
        Cell cell = row.createCell(colNum);
        if (field != null) {
            if (field instanceof String)
                cell.setCellValue((String) field);
            if (field instanceof Integer)
                cell.setCellValue((Integer) field);
        }
    }

    private static void CreateExcelColumn(Object field, Row row, int colNum, CellStyle style) {
        Cell cell = row.createCell(colNum);
        cell.setCellStyle(style);
        if (field != null) {
            if (field instanceof String)
                cell.setCellValue((String) field);
            if (field instanceof Integer)
                cell.setCellValue((Integer) field);
        }
    }

}
