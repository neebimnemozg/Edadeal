/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package edadeal;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.Set;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

/**
 *
 * @author ksmirnov
 */
public class Edadeal {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws FileNotFoundException, IOException, ClassNotFoundException, SQLException {

        File dataFile = new File("D://parse_5.xls");
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(dataFile));
        HSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
        Iterator<Row> rowIterator1 = sheet.iterator();
        Iterator<Row> rowIterator2 = sheet.iterator();
        Iterator<Row> rowIterator3 = sheet.iterator();
        Iterator<Row> rowIterator4 = sheet.iterator();
        Row row;
        //Set<String> links = new LinkedHashSet<String>();
        //Set<String> flinks = new LinkedHashSet<String>();
        //Set<String> shops = new LinkedHashSet<String>();
        Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
        Connection con = DriverManager.getConnection("jdbc:sqlserver://srv-mbc-sql-12;databaseName=ANALYTICS;integratedSecurity=true;user=KSmirnov");
        Statement st = con.createStatement();
        String s;
        String s1;
        String shop;
        String city = null;
        String dat = null;
        String product = null;
        String price = null;
        int counter = 1;
        java.sql.Date date = new java.sql.Date(Calendar.getInstance().getTime().getTime());

        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            //links.add(row.getCell(0).toString());
            for (int i = 3; i < 33; i++) {
                if (row.getCell(i) != null) {
                    //flinks.add(row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).toString());
                    s = row.getCell(i, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).toString();
                    String[] parts = s.split("from");
                    s1 = parts[0];
                    if (null == row.getCell(1).toString()) {
                        city = "-";
                    } else {
                        switch (row.getCell(1).toString()) {
                            case "Скидки в категории «Пиво и сидр» в Москве — Едадил":
                                city = "Москва";
                                break;
                            case "Скидки в категории «Пиво и сидр» в Санкт-Петербурге — Едадил":
                                city = "Санкт-Петербург";
                                break;
                            case "Скидки в категории «Пиво и сидр» в Казани — Едадил":
                                city = "Казань";
                                break;
                            case "Скидки в категории «Пиво и сидр» в Воронеже — Едадил":
                                city = "Воронеж";
                                break;
                            case "Скидки в категории «Пиво и сидр» в Екатеринбурге — Едадил":
                                city = "Екатеринбург";
                                break;
                            case "Скидки в категории «Пиво и сидр» в Омске — Едадил":
                                city = "Омск";
                                break;
                            case "Скидки в категории «Пиво и сидр» в Волгограде — Едадил":
                                city = "Волгоград";
                                break;
                            case "Скидки в категории «Пиво и сидр» в Уфе — Едадил":
                                city = "Уфа";
                                break;
                            case "Скидки в категории «Пиво и сидр» в Перми — Едадил":
                                city = "Пермь";
                                break;
                            case "Скидки в категории «Пиво и сидр» в Новосибирске — Едадил":
                                city = "Новосибирск";
                                break;
                            default:
                                city = "-";
                                break;
                        }
                    }
                    st.executeUpdate("INSERT INTO [ANALYTICS].[dbo].[links] (Sku, DT, DailyID, City) VALUES ('" + s1 + "', '" + date + "', '" + counter + "', '" + city + "')");
                    counter = counter + 1;
                }
            }
        }
        counter = 1;
        while (rowIterator1.hasNext()) {
            row = rowIterator1.next();
            for (int j = 99; j < 131; j++) {
                if (row.getCell(j) != null && !row.getCell(j).toString().contains("Едадил")) {
                    //shops.add(row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).toString());
                    price = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).toString();
                    st.executeUpdate("UPDATE [ANALYTICS].[dbo].[links] SET Price = ('" + price.replaceAll("От ", "").replaceAll(" руб", "") + "') WHERE DailyID = '" + counter + "'");
                    counter = counter + 1;
                }
            }
        }
        counter = 1;
        while (rowIterator2.hasNext()) {
            row = rowIterator2.next();
            for (int j = 34; j < 66; j++) {
                if (row.getCell(j) != null && !row.getCell(j).toString().contains("Едадил")) {
                    //shops.add(row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).toString());
                    dat = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).toString();
                    st.executeUpdate("UPDATE [ANALYTICS].[dbo].[links] SET DAT = ('" + dat.replaceAll("июля", "июль") + "') WHERE DailyID = '" + counter + "'");
                    counter = counter + 1;
                }
            }
        }
        counter = 1;
        while (rowIterator3.hasNext()) {
            row = rowIterator3.next();
            for (int j = 66; j < 99; j++) {
                if (row.getCell(j) != null && !row.getCell(j).toString().contains("Едадил")) {
                    //shops.add(row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).toString());
                    product = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).toString();
                    st.executeUpdate("UPDATE [ANALYTICS].[dbo].[links] SET Prouct = ('" + product.replaceAll("'", "") + "') WHERE DailyID = '" + counter + "'");
                    counter = counter + 1;
                }
            }
        }
        counter = 1;
        while (rowIterator4.hasNext()) {
            row = rowIterator4.next();
            for (int j = 163; j < 264; j++) {
                if (row.getCell(j) != null && !row.getCell(j).toString().contains("Едадил")) {
                    //shops.add(row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).toString());
                    shop = row.getCell(j, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL).toString().replaceAll("'", "\\'");
                    st.executeUpdate("UPDATE [ANALYTICS].[dbo].[links] SET Shop = ('" + shop + "') WHERE DailyID = '" + counter + "'");
                    counter = counter + 1;
                }
            }
        }
        workbook.close();
    }
}
