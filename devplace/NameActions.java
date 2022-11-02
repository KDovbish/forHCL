import lotus.domino.*;

public class NameActions {


    /*
     * Получить каноническое имя
     *
     * Ivan I Ivanov/KIB		->	CN=Ivan I Ivanov/O=KIB
     * CN=Ivan I Ivanov/O=KIB	->	CN=Ivan I Ivanov/O=KIB
     * Ivan I Ivanov			->	Ivan I Ivanov
     */
    static String GetCanonicalName(String cName){
        try{
            Session hSession=NotesFactory.createSession();
            Name hN=hSession.createName(cName);
            return hN.getCanonical();
        } catch (Exception e) {
            return "";
        }
    }


    //  Получить компонент CN заданного имени
    static String GetCommonName(String cFullName){
        try{
            Session hSession;
            Name hN;

            hSession=NotesFactory.createSession();
            hN=hSession.createName(cFullName);
            return hN.getCommon();

        }catch (Exception e){
            return "";
        }
    }


    //  Получать Lastname заданного имени
    static String GetLastName(String cFullName){
        try{
            String cCN;
            String[] aNamePart;

            cCN=GetCommonName(cFullName);
            aNamePart=cCN.split(" ");
            if (aNamePart.length>0) return aNamePart[aNamePart.length-1];
            else return "";

        }catch (Exception e){
            return "";
        }
    }


    //  Получить сокращенное имя
    static String GetAbbreviatedName(String cName){
        try{
            Session hSession;
            Name hN;

            hSession=NotesFactory.createSession();
            hN=hSession.createName(cName);
            return hN.getAbbreviated();
        }catch (Exception e){
            return "";
        }
    }

}