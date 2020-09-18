package Factory;

import Model.PersonBase;
import Model.PersonExcel1;
import Model.PersonExcel2;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;

//Bu metodta FactoryDesign Pattern kullanarak nesne uretme islemi yapilmaktadir.

public class PersonFactory {

    public PersonBase buildPerson(Iterator<Cell> cellIterator, String fileName) {
        PersonBase personBase = null;
        Integer column=1;

        if (fileName.equals ("Excel1.xlsx")) {
            personBase = new PersonExcel1 ();
            PersonExcel1 personExcel1 = (PersonExcel1) personBase;

            while (cellIterator.hasNext ()) {
                Cell cell = cellIterator.next ();
                String cellValue = cell.toString ();
                if (column==1){
                    personBase.setAd (cellValue);
                }
                else if (column==2){
                    personBase.setSoyad (cellValue);
                }
                else if (column==3){
                    personBase.setDogumTarihi (cellValue);
                }
                else if (column == 4) {
                    personBase.setMail (cellValue);
                } else if (column == 5) {
                    personBase.setTelefon (cellValue);

                } else if (column == 6) {
                    personExcel1.setDurum (cellValue);
                }
                column = column + 1;

            }
        }
        else if (fileName.equals ("Excel2.xlsx")){
            personBase= new PersonExcel2 ();
            PersonExcel2 personBase2 = (PersonExcel2) personBase;
            while (cellIterator.hasNext ()){
                Cell cell = cellIterator.next ();
                String cellValue = cell.toString ();
                if (column==1){
                    personBase.setAd (cellValue);
                }
                else if (column==2){
                    personBase.setSoyad (cellValue);
                }
                else if (column==3){
                    personBase.setDogumTarihi (cellValue);
                }
                else if (column==4){
                    personBase2.setDogumYeri (cellValue);
                }
                else if (column==5){
                    personBase2.setMail (cellValue);
                }
                else if (column==6){
                    personBase2.setTelefon (cellValue);
                }
                else if (column==7){
                    personBase2.setCalismaDurumu (cellValue);
                }
                else if (column==8){
                    personBase2.setUniversite (cellValue);
                }
                column++;

            }
        }

        return personBase;
    }

}
