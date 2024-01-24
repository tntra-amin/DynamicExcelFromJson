package com.example.DynamicExcel.DynamicExcel.POC;

import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class SongDataExcelGenerator {

    static String[] genres = {"Country","Electronic music","Jazz","Pop Music","Rock","Classical","Metal"};

    static String[] artists = {"Anu Malik", "Atif Aslam", "Shreya Ghoshal", "Arijit Singh", "Sonu Nigam", "Lata Mangeshkar", "Neha Kakkar", "KK", "Armaan Malik", "B Praak"};

    static String playCountFormula = "ROUND(RAND()*(10-1)+1,0)";

    static String timeStampFormula = "TEXT(RAND(), \"HH:MM\")";

    static Map<String, List<String>> songs = new HashMap<>();

    public static void initializer(){

        songs.put("Anu Malik", List.of("Ye Kaali Kaali Aankhen", "Baazigar O Baazigar", "Chura Ke Dil Mera", "Sandese Aate Hai", "Chori Chori Dil Tera", "Dil Kehta Hai", "Main Hoon Na", "Tumse Milke Dil Ka", "Neend Churayee Meri", "Aa Gale Lag Jaa"));

        songs.put("Atif Aslam", List.of("Dil Meri Na Sune", "Piya O Re Piya", "Paniyon Sa", "Musafir", "O Saathi", "Dekhte Dekhte", "Younhi", "O Saathi", "Jeena Jeena", "Tera Hua"));

        songs.put("Shreya Ghoshal", List.of("Agar Tum Mil Jao", "Mere Dholna", "Saibo", "Manwa Laage", "Deewani Mastani", "Pal Pal Har Pal", "Hasi", "Piyu Bole", "Jaadu Hai Nasha", "Saans"));

        songs.put("Arijit Singh", List.of("Phir Bhi Tumko Chaahunga", "Har Kisi Ko", "Tum Hi Ho", "Phir Mohabbat", "Hum Mar Jayenge", "Kabira", "Tose Naina", "Milne Hai Mujhse Aayi", "Sanam Re", "Agar Tum Saath Ho"));

        songs.put("Sonu Nigam", List.of("Achchha Sila Diya Toone Mere Pyar Ka", "Deewane Hoke Hum", "Jeena", "Do Pyaar", "Mera Rang De Basanti.", "Sandese Aate Hain", "Tu Fiza Hai", "Main Ki Karaan", "Papa mere papa", "Bole Chudhiyan"));

        songs.put("Lata Mangeshkar", List.of("Mujhse Juda Hokar", "Ek Pyar Ka Nagma Hai", "Main Teri Dushman", "Mera Dil Ye Pukare", "Der ho na jaye kahin", "Mera Dil Yeh Pukare Aaja", "Gori Kab Se Huyee Jawan", "Dil Deewana", "Robb Ne Banaya Tujhe Mere Liye", "Bangle Ke Peechhe"));

        songs.put("Neha Kakkar", List.of("Mile Ho Tum", "Baarish Mein Tum", "Kala Chashma", "Gaadi Kaali", "Narazgi","Aao Raja", "Mehbooba","Suroor", "La La La", "Zindagi Mil Jayegi"));

        songs.put("KK", List.of("Tu Hi Meri Shab Hai", "Pyaar Ke Pal","Ankhon Mein Teri", "Tadap Tadap", "Dil Ibaadat", "Tu Jo Mila", "Yaaron", "Beete Lamhein", "Kya Mujhe Pyaar Hai", "O Meri Jaan"));

        songs.put("Armaan Malik", List.of("JAB TAK", "Dil Mein Ho Tum", "Bol Do Na Zara", "Hua Hain Aaj Pehli Baar", "Tumhe Apna Banane Ka", "Sab Tera", "Mujhko Barsaat Bana Lo", "Buddhu Sa Mann", "Tere Mere", "Dil Mein Chhupa Loonga"));

        songs.put("B Praak", List.of("Kya loge tum", "Filhaal 2", "Filhaal", "Baarish ki jaaye", "Roohedariyaan", "Bijlee Bijlee", "Ittar", "Ik mili mainu apsara", "Kudiyan lahore diyan", "Jaani ve Jaani"));

    }

    public static int getRandomNumber(int max) {
        return (int) ((Math.random() * max));
    }

    public static List<String> getRow(){

        List<String> res = new ArrayList<>();

        res.add(timeStampFormula);

        String artist = artists[getRandomNumber(artists.length)];
        res.add(artist);

        String genre = genres[getRandomNumber(genres.length)];
        res.add(genre);

        List<String> songTitles = songs.get(artist);
        String songTitle = songTitles.get(getRandomNumber(songTitles.size()));
        res.add(songTitle);

        res.add(playCountFormula);

        return res;

    }

    public static void main(String[] args) throws IOException {

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("song_data");

        XSSFFont boldFont = workbook.createFont();
        boldFont.setFontName("Calibri");
        boldFont.setFontHeightInPoints((short) 11);
        boldFont.setBold(true);

        XSSFCellStyle boldCellStyle = workbook.createCellStyle();
        boldCellStyle.setFont(boldFont);

        XSSFRow row;
        XSSFCell cell;
        List<String> data;

        initializer();

        row = sheet.createRow(0);

        cell = row.createCell(0);
        cell.setCellValue("TIMESTAMP");
        cell.setCellStyle(boldCellStyle);

        cell = row.createCell(1);
        cell.setCellValue("ARTIST");
        cell.setCellStyle(boldCellStyle);

        cell = row.createCell(2);
        cell.setCellValue("GENRE");
        cell.setCellStyle(boldCellStyle);

        cell = row.createCell(3);
        cell.setCellValue("SONG TITLE");
        cell.setCellStyle(boldCellStyle);

        cell = row.createCell(4);
        cell.setCellValue("PLAY COUNT");
        cell.setCellStyle(boldCellStyle);

        for(int i=1 ;i<100000;i++){
            row = sheet.createRow(i);

            data = getRow();

            cell = row.createCell(0);
            cell.setCellFormula(data.get(0));

            cell = row.createCell(1);
            cell.setCellValue(data.get(1));

            cell = row.createCell(2);
            cell.setCellValue(data.get(2));

            cell = row.createCell(3);
            cell.setCellValue(data.get(3));

            cell = row.createCell(4);
            cell.setCellFormula(data.get(4));

        }

        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);
        sheet.autoSizeColumn(4);

        FileOutputStream out = new FileOutputStream(new File("User_song_data.xlsx"));
        workbook.write(out);
        workbook.close();
        out.close();


    }
}
