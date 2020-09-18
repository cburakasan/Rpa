package Write;


import Model.PersonBase;
import Model.PersonExcel1;
import Model.PersonExcel2;
import Model.PersonReport;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;
import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class WriteAndReport {
    private static Logger log = Logger.getLogger (WriteAndReport.class);

    //Bu metodta yazma islemi sonucunda elde edilen listeler verilen karsilastirma kriterlerine uygun olarak karsilastirilip
    // raporlanacak yeni liste elde edilmektedir.
    public List<PersonReport> getReportPersonList(List<PersonBase> excel1List, List<PersonBase> excel2List) {
        List<PersonReport> reportList = new ArrayList<PersonReport> ();
        for (PersonBase excel1 : excel1List) {
            String ad = excel1.getAd ();
            String soyad = excel1.getSoyad ();
            String telefon = excel1.getTelefon ();
            String mail = excel1.getMail ();
            String dogumTarihi = excel1.getDogumTarihi ();
            PersonExcel1 excel1New = (PersonExcel1) excel1;
            String durum = excel1New.getDurum ();
            for (PersonBase excel2 : excel2List) {
                String ad2 = excel2.getAd ();
                String soyad2 = excel2.getSoyad ();
                String mail2 = excel2.getMail ();
                String telefon2 = excel2.getTelefon ();
                PersonExcel2 excel2New = (PersonExcel2) excel2;
                String calismaDurumu = excel2New.getCalismaDurumu ();
                String dogumYeri = excel2New.getDogumYeri ();
                String universite = excel2New.getUniversite ();
                if (ad.equals (ad2) && soyad.equals (soyad2)) {
                    if (telefon == null || mail == null) {
                        log.error ("telefon veya mail adreslerinden biri dolu olmalidir, excel1 de olmayan veri " + excel1.getAd () + " " + excel1.getSoyad ());
                        break;

                    } else if (telefon2 == null || mail2 == null) {
                        log.error ("telefon veya mail adreslerinden biri dolu olmalidir, excel2 de olmayan veri " + excel2.getAd () + " " + excel2.getSoyad ());
                        break;
                    }
                    if (telefon.equals (telefon2) || mail.equals (mail2)) {
                        PersonReport personReport = new PersonReport ();
                        personReport.setAd (ad);
                        personReport.setSoyad (soyad);
                        personReport.setMail (mail);
                        personReport.setTelefon (telefon);
                        personReport.setDurum (durum);
                        personReport.setCalismaDurumu (calismaDurumu);
                        personReport.setDogumYeri (dogumYeri);
                        personReport.setDogumTarihi (dogumTarihi);
                        personReport.setUniversite (universite);
                        reportList.add (personReport);
                        break;

                    } else {
                        log.error ("telefon veya mail adreslerinden biri ayni olmalidir " + excel1.getAd () + " " + excel1.getSoyad ());  //
                        break;
                    }
                }
            }
        }
        return reportList;
    }

    public void reportBuild(List<PersonReport> reportPersonList) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook ();
        XSSFSheet sheet = workbook.createSheet();

        int rowCount = 0;
        for (PersonReport personReport : reportPersonList) {
            Row row = sheet.createRow(rowCount);
            writePerson(personReport, row);
            rowCount++;
        }

        String path = "C:\\Users\\cbura\\Desktop\\OptimRpa\\src\\main\\resources\\Report\\Report.xlsx";
        try  {
            FileOutputStream outputStream = new FileOutputStream(path);
            workbook.write(outputStream);
        }
        catch (Exception e){
            e.getStackTrace ();
        }

    }

    private void writePerson(PersonReport personReport, Row row) {
        Cell cellAd = row.createCell(0);
        cellAd.setCellValue(personReport.getAd ());

        Cell cellSoyad = row.createCell(1);
        cellSoyad.setCellValue(personReport.getSoyad ());

        Cell cellDogumTarihi = row.createCell(2);
        cellDogumTarihi.setCellValue(personReport.getDogumTarihi ());

        Cell cellDogumYeri = row.createCell(3);
        cellDogumYeri.setCellValue(personReport.getDogumYeri ());

        Cell cellMail = row.createCell(4);
        cellMail.setCellValue(personReport.getMail ());

        Cell cellDurum = row.createCell(5);
        cellDurum.setCellValue(personReport.getDurum ());

        Cell cellCalismaDurumu = row.createCell(6);
        cellCalismaDurumu.setCellValue(personReport.getCalismaDurumu ());

        Cell cellUniversite = row.createCell(7);
        cellUniversite.setCellValue(personReport.getUniversite ());

        Cell cellTelefon = row.createCell(8);
        cellTelefon.setCellValue(personReport.getTelefon ());


    }
}

