package Model;

import lombok.Data;


@Data
public abstract class PersonBase {
    private String ad;
    private String soyad;
    private String dogumTarihi;
    private String mail;
    private String telefon;

}
