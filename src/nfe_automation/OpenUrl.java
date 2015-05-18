/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package nfe_automation;

import jxl.read.biff.BiffException;
import org.openqa.selenium.WebDriver;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.By;

public class OpenUrl {

    OpenDataBase opendatabase = new OpenDataBase();
    FileOutputStream fOut = null;

    public void OpenURL(WebDriver driver) throws BiffException, Exception {
        String tablename = ("tb");
        opendatabase.Open(tablename);
        File file = new File("BD.xls");
        int result = 0;
        String STR = "CRI -";
        int ARQ = 1;
        int WARNING = 2;
        int CRITICAL = 3;

        Workbook wb = new HSSFWorkbook(new FileInputStream(file.getCanonicalPath()));

        for (int i = 0; i < opendatabase.sheet.getRows(); i++) {

            String url = opendatabase.sheet.getCell(0, i).getContents();

            double start = System.currentTimeMillis();
            driver.manage().window().maximize();
            driver.get(url);
            double finish = System.currentTimeMillis();
            double totalTime = finish - start;
            double totalTimeSec = totalTime / 1000;

            try {
                driver.findElement(By.xpath("/html/body/div/span[2]/ul/li/a")).click();
            } catch (Exception e) {
                driver.findElement(By.xpath("/html/body/div/span/ul/li/a")).click();
            }

            String titlepage = driver.getTitle().replace(" ", "");

            FileWriter fw = new FileWriter(titlepage + ".txt");
            int statusCode = Statuscode(driver.getCurrentUrl(), driver);
            String line = url + "," + totalTimeSec + "," + statusCode;

            fw.write(line);
            fw.close();


        }


    }

    public int Statuscode(String URLName, WebDriver driver) {
        try {
            int statuscode = Integer.parseInt(driver.findElement(By.xpath("/html/body/div/span/span/pre[2]")).getText().substring(9, 12));
            return statuscode;
        } catch (Exception e) {
            e.printStackTrace();
            return 1;
        }
    }
}
