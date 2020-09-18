import Model.PersonBase;
import Model.PersonReport;
import Reading.Reading;
import Write.WriteAndReport;
import org.apache.log4j.Logger;
import java.io.File;
import java.io.IOException;
import java.util.List;


public class Run {
    private static Logger log = Logger.getLogger (Run.class);
    public static void main(String[] args) {
        Reading reading = new Reading ();
        File fileExcel1 = new File ("C:\\Users\\cbura\\Desktop\\OptimRpa\\src\\main\\resources\\InputData\\Excel1.xlsx");
        File fileExcel2 = new File ("C:\\Users\\cbura\\Desktop\\OptimRpa\\src\\main\\resources\\InputData\\Excel2.xlsx");
        WriteAndReport writeAndReport = new WriteAndReport ();
        List<PersonBase> excel1List =null;
        List<PersonBase> excel2List =null;
        try {
            excel1List = reading.readingProcess (fileExcel1);
            excel2List = reading.readingProcess (fileExcel2);
        } catch (Exception e) {
            log.error ("Excelden veri okunurken bir hata olustu. Hata mesaji: "+e.getMessage ());
            throw new RuntimeException (); //Excelden veri okunurken bir hata ile karsilasildiginda kod akisi sonlandrip ilgili hata mesaji verildi.
        }
        try {

            List<PersonReport> reportPersonList = writeAndReport.getReportPersonList (excel1List, excel2List);
            writeAndReport.reportBuild (reportPersonList);

        } catch (IOException e) {
            log.error ("Excele veri yazilirken bir hata olustu. Hata mesaji: "+e.getMessage ());
            throw new RuntimeException ();
        }
    }
}
