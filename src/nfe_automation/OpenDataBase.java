
package nfe_automation;

import java.io.File;
import java.io.IOException;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;

public class OpenDataBase{

    Sheet sheet;
    Workbook bdurl;
      
    public Sheet Open(String tablename) throws BiffException, IOException, Exception { 

        File file = new File("BD.xls");
            WorkbookSettings ws = new WorkbookSettings();
            ws.setEncoding("Cp1252");
            
            bdurl = Workbook.getWorkbook(new File(file.getCanonicalPath()), ws);             
            
            sheet = bdurl.getSheet(tablename);
            return sheet;
        }

    public void Close(Workbook bdurl) throws BiffException, IOException, Exception {

        bdurl.close();
    }
}
