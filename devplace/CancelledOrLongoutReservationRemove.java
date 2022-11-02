import java.util.*;

import lotus.domino.*;

/**
 * Подчистка заказанных ресурсов за уволенными и длительно-отсутствующими
 */
public class CancelledOrLongoutReservationRemove extends AgentBase {

    //  Имя представления, документы которого анализируются агентом
    private final String RESERVATION_VIEW="ResByDate";
    //  Местоположение Списка сотрудников
    private String STAFF_SERVER="APP-001/KIB";
    private String STAFF_DB="staff.nsf";
    //  Задержка в днях от даты увольнения/выхода в длит.отсутствие, после которой возможно удаление резервирований
    private int DAYS_DELAY=7;
    //  Местоположение базы, в которой будут сохраняться документы перед удалением из продбазы
    private String cBackupServer="EMA/KIB";
    private String cBackupDb="temp\\rrbackup.nsf";
    private Vector aRecipient;


    public void NotesMain() {

        try {

            //  Приостановка для отладки
            //Thread.sleep(20000);


            Session session = getSession();
            AgentContext agentContext = session.getAgentContext();

            Database hDb;
            Document hDoc;
            Database hNames, hStaff;
            String cReservedFor;
            DateTime dt=session.createDateTime(new Date());
            DateTime dtStartDateTime;
            ArrayList<Document> aReservDoc=new ArrayList<Document>();
            View hView;
            String cRQStatus;
            Document hPDoc;


            //  Получаем ссылку на текущую базу
            //hDb=agentContext.getCurrentDatabase();
            //hDb=session.getDatabase("Parsek-010/KIB", "Resource Reservations.nsf");
            hDb=session.getDatabase("EMA/KIB", "temp\\rrtest.nsf");

            //  Читаем профильный документ с настройками
            hPDoc=hDb.getProfileDocument("CancelledOrLongOutReservationRemove", "");

            //  Считываем параметры конфигурационного документа
            STAFF_SERVER=hPDoc.getItemValueString("StaffServer");
            STAFF_DB=hPDoc.getItemValueString("StaffDb");
            DAYS_DELAY=hPDoc.getItemValueInteger("DaysDelay");
            cBackupServer=hPDoc.getItemValueString("BackupServer");
            cBackupDb=hPDoc.getItemValueString("BackupDb");
            aRecipient=hPDoc.getItemValue("RecipientList");

            //  Выбираем все документы резервирования помещений в статусе "Заказ принят"(RQStatus="A") для всех дат начиная от сегодня.
            //  Заказы в состоянии "Ожидает утверждения"("T") не беру, т.к. данное время свободно для любого заказа другим сотрудником.
            //cQuery="SELECT Form=\"Reservation\" & @IsAvailable(ReservedFor) & RQStatus=\"A\" & @Date(StartDateTime)>=[" + session.createDateTime("Today").getDateOnly() + "]";
            //hDColl=hDb.search(cQuery);

            //  Открывается представление, документы которого будут анализироваться
            hView=hDb.getView(RESERVATION_VIEW);
            if (hView!=null) {
                if (hView.getEntryCount()>0) {
                    //  Открываем АК
                    hNames=session.getDatabase(hDb.getServer(), "names.nsf");
                    if (hNames.isOpen()){
                        //  Открываем Список сотрудников
                        hStaff=session.getDatabase(STAFF_SERVER, STAFF_DB);
                        if (hStaff.isOpen()) {

                            //  Перебираем все выбранные из текущей базы записи...
                            hDoc=hView.getFirstDocument();
                            while (hDoc!=null) {

                                cReservedFor=hDoc.getItemValueString("ReservedFor");
                                dtStartDateTime=(DateTime) hDoc.getItemValueDateTimeArray("StartDateTime").elementAt(0);
                                cRQStatus=hDoc.getItemValueString("RQStatus");

                                //  Если дата резервирования помещения не является уже прошедшей датой и резервирование в статусе "Заказ принят"...
                                if (dtStartDateTime.timeDifference(session.createDateTime("Today"))>=0 && cRQStatus.equals("A")) {
                                    //  Существует ли Mail-In база с именем для которого зарезервировано помещение...
                                    if (  Engine.IsDatabase(hNames, cReservedFor)==false ) {
                                        //  Существует ли учетная запись пользователя с именем для которого зарезервировано помещение...
                                        if ( Engine.IsPerson(hNames, cReservedFor)==false ) {
                                            //  Анализируем список сотрудников на предмет уволенности или длительного отсутствия учетной записи пользователя...
                                            if (  Engine.IsCancelledOrLongOut(hStaff, cReservedFor, dt) ) {

                                                //  В адресной книге учетной записи сотрудника, который зарезервировал помещение, нет.
                                                //  В Списке сотрудников такой сотрудник является либо уволенным либо длительно отсутствующим.

                                                //  Если дата увольнения лежит в прошлом(или это сегодня) и с даты увольнения прошло как минимум DAY_DELAY дней...
                                                if ((dt.timeDifference(session.createDateTime("Today"))<=0) && (Engine.DayDiffrence(dt, session.createDateTime("Today"))>=DAYS_DELAY)) {
                                                    aReservDoc.add(hDoc);
                                                }

                       						  /*
                       						  //  Если дата увольнения/выхода в длит.отсут. лежит в прошлом, по отношению к дате резервирования ресурса и
                       						  //  между этими датами прошло как минимум 30 дней...
                       						  if ((dt.timeDifference(dtStartDateTime)<=0) && (Engine.DayDiffrence(dt, dtStartDateTime)>=DAYS_DELAY)) {
                       							  aReservDoc.add(hDoc);
                       						  }
                       						  */
                                            }
                                            else {
                                                //  В адресной книге учетной записи сотрудника, который зарезервировал помещение, нет.
                                                //  В Списке сотрудников такой сотрудника нет ни среди уволенных ни среди длительно отсутствующих.
                                                //  Ключевым моментом для данной ветки алгоритма является то, что сотрудника уже нет в адресной книге!
                                                aReservDoc.add(hDoc);
                                            }

                                        }
                                    }
                                }

                                hDoc=hView.getNextDocument(hDoc);
                            }

                            if (aReservDoc.isEmpty()==false) {
                                //  Делаем рассылку только в том случае, если список получателей не пуст.
                                //  (приходиться делать рассылку перед удалением, поскольку после удаления уже неоткуда будет брать информацию)
                                if (aRecipient.isEmpty()==false) SendMessage(aReservDoc, aRecipient);
                                //  Предварительно копируем документы которые будут удаляться в другую базу
                                if (cBackupDb.equals("")==false) copyDocumentArrayToOtherDb(aReservDoc, cBackupServer, cBackupDb);
                                //  Удаляем документы
                                removeDocumentArray(aReservDoc);
                            }

                        }
                    }
                }
            }



        } catch(Exception e) {
            e.printStackTrace();
        }
    }


    /*
     * Удалаяет массив докуметов
     * Входные параметры:
     * 		aDoc	массив документов, подлежащих удалению
     */
    private void removeDocumentArray(ArrayList<Document> aDoc) {
        try {
            int i;

            for (i=0; i<=aDoc.size()-1; i++) {
                aDoc.get(i).removePermanently(false);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    /*
     * Сохранение списка документов в заданной базе
     *
     * Входные параметры:
     * 	aDoc		массив объектов типа Document
     * 	cServer		имя сервера
     * 	cDb			имя базы
     */
    private void copyDocumentArrayToOtherDb(ArrayList<Document> aDoc, String cServer, String cDb) {

        try {
            Session session = getSession();
            Database hDb;
            int i;
            Document hDoc;


            //  Пытаемся открыть базу. Если база не существует(возврат null), то создаем ее
            hDb=session.getDatabase(cServer, cDb, false);
            if (hDb==null) {
                hDb=session.getCurrentDatabase().createCopy(cServer, cDb);
            }

            if (hDb!=null) {
                //  Копируем все документы из массива в базу
                for (i=0; i<=aDoc.size()-1; i++) {
                    aDoc.get(i).copyToDatabase(hDb);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

    }



    /*
     * Рассылка сообщений о документах резервирования
     * Используется для рассылки уведомлений о документах резервирования, которые были удалены данным агентом
     * Входные параметры:
     * 	aReservDoc		массив документов резервирования
     * 	aRecipients		получатели сообщения
     */
    private void SendMessage(ArrayList<Document> aReservDoc, Vector aRecipients) {

        try {
            Database hDb=getSession().getCurrentDatabase();
            Document hDoc;
            RichTextItem hBody;
            int i;
            Document hReservDoc;
            DateTime dStartDateTime;
            RichTextStyle hRTextStyle;


            hDoc=hDb.createDocument();
            hDoc.replaceItemValue("Form","Memo");
            //hDoc.replaceItemValue("SendTo", "CN=Konstantin G Dovbish/O=KIB");
            hDoc.replaceItemValue("SendTo", aRecipients);
            hDoc.replaceItemValue("Principal", "CancelledOrLongoutReservationRemove Agent");
            hDoc.replaceItemValue("Subject", "Удалены резервирования уволенных/длит.отсутствующих");

            hBody=hDoc.createRichTextItem("Body");
            hRTextStyle=getSession().createRichTextStyle();

            hRTextStyle.setFontSize(12);
            hRTextStyle.setBold(RichTextStyle.YES);
            hBody.appendStyle(hRTextStyle);
            hBody.appendText("Список удаленных документов резервирования, которые изначально были созданы уволенными или длит.отсутствующими");
            hBody.addNewLine();
            hRTextStyle.setFontSize(10);
            hBody.appendStyle(hRTextStyle);
            hBody.appendText("Со дня увольнения/выхода в длит.отсутствие прошло(дней): " + DAYS_DELAY);
            hBody.addNewLine(2);

            hRTextStyle.setFontSize(10);
            hRTextStyle.setBold(RichTextStyle.NO);
            hBody.appendStyle(hRTextStyle);
            for (i=0; i<=aReservDoc.size()-1; i++) {

                hReservDoc=aReservDoc.get(i);
                dStartDateTime=(DateTime)hReservDoc.getItemValueDateTimeArray("StartDateTime").elementAt(0);
                hBody.appendText( dStartDateTime.getDateOnly()+" "+dStartDateTime.getTimeOnly() + "\t" +
                        String.format("%-50s", NameActions.GetAbbreviatedName(hReservDoc.getItemValueString("ReservedFor"))) + "\t" +
                        NameActions.GetAbbreviatedName(hReservDoc.getItemValueString("ResourceName")) );
                hBody.addNewLine();
            }

            hDoc.send();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}