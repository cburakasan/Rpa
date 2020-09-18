package Reading;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import Factory.PersonFactory;
import Model.PersonBase;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Reading {
// Bu metodta excel dosyalarini okuyup verilerin person listesinde tutulmasini saglamaktadir.
    public List<PersonBase> readingProcess(File file) throws IOException {
        FileInputStream fis = new FileInputStream (file);
        String fileName = file.getName ();
        XSSFWorkbook workbook = new XSSFWorkbook (fis);
        XSSFSheet sheet = workbook.getSheetAt (0);
        Iterator<Row> rowIt = sheet.iterator ();
        PersonFactory personFactory = new PersonFactory ();
        List<PersonBase> personBaseList = new ArrayList<PersonBase> ();
        while (rowIt.hasNext ()) {
            Row row = rowIt.next ();
            Iterator<Cell> cellIterator = row.cellIterator ();
            PersonBase personBase = personFactory.buildPerson (cellIterator, fileName);
            personBaseList.add (personBase);
        }
        workbook.close ();
        fis.close ();
        return personBaseList;
    }
}
