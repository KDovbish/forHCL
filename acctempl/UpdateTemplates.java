import lotus.domino.*;
import java.util.Date;
import java.util.Vector;

/**
 * Обновить шаблоны
 */
public class UpdateTemplates extends AgentBase {

    public void NotesMain() {

        try {
            Session session = getSession();
            AgentContext agentContext = session.getAgentContext();
            String cCommand;


            System.out.println("UpdateTemplates  НАЧАЛО");

            //  Получаем ссылку на конфигурационный документ агента
            Document hConf=getAgentConfig();

            System.out.println("UpdateTemplates  Блокировка основного кода: " + hConf.getItemValueString("DisableMainCode") );
            System.out.println("UpdateTemplates  Сервер: " + getAbbreviatedName( hConf.getItemValueString("Server") ));
            System.out.println("UpdateTemplates  Маска: " + hConf.getItemValueString("DbMask") );
            System.out.println("UpdateTemplates  Искомый шаблон на базах: " + hConf.getItemValueString("TemplateName") );
            System.out.println("UpdateTemplates  Менять на шаблон: " + hConf.getItemValueString("NewTemplateFileName") );
            System.out.println("UpdateTemplates  Максимальное количество баз для обработки: " + hConf.getItemValueInteger("DbMax") );

            //  Если выполнение основного когда агента не блокировано...
            if ( hConf.getItemValueString("DisableMainCode").equals("1")==false ) {

                //  Создается заготовка сообщения, в котором будут все базы, которые подверглись обновлению шаблона
                Document hMsg=session.getCurrentDatabase().createDocument();
                RichTextItem hMsgBody=hMsg.createRichTextItem("Body");
                fillingMessage(hConf.getItemValue("Recipients"),
                        "Лог обновления шаблонов",
                        getAbbreviatedName( hConf.getItemValueString("Server") ) + " " + hConf.getItemValueString("DbMask") + "   " + hConf.getItemValueString("TemplateName") + " -> " + hConf.getItemValueString("NewTemplateFileName"),
                        hMsg,
                        hMsgBody);


                //  Счетчик обработанных баз
                int nDbProcessed=0;

                //  Перебираем базы на заданном сервере...
                DbDirectory hDbDir=session.getDbDirectory( hConf.getItemValueString("Server") );
                Database hDb=hDbDir.getFirstDatabase(DbDirectory.DATABASE);
                while (hDb!=null) {

                    //  Если имя найденно йбазы подпадает под искомый шаблон...
                    if ( satisfyFilePath(hConf.getItemValueString("DbMask"), hDb.getFilePath()) ) {
                        //  Если шаблон базы соответствует шаблону, заданному в конфигурационном документе, т.е.
                        //  требуется обновление шаблона данной базы...
                        if ( hDb.getDesignTemplateName().equals(hConf.getItemValueString("TemplateName") ) ) {

                            //  Отсылаем на консоль заданного сервера команду замены шаблона для заданной базы
                            cCommand="load convert -u " + hDb.getFilePath() + " " + hDb.getDesignTemplateName() + " " + hConf.getItemValueString("NewTemplateFileName");
                            System.out.println("UpdateTemplates  " + cCommand);
                            session.sendConsoleCommand(hConf.getItemValueString("Server") , cCommand);

                            //  Наполнение тела информационного сообщения именами баз, подвергшися обновлению шаблона
                            hMsgBody.appendText( String.format("%-40s", hDb.getFilePath()) + hDb.getTitle()); hMsgBody.addNewLine();

                            //  Фиксируем факт обновления шаблона на базе
                            addTemplateUpdateLog( hDb.getServer(), hDb.getFilePath(), hDb.getTitle(), hDb.getDesignTemplateName(), hConf.getItemValueString("NewTemplateFileName")  );

                            //  Если берем в расчет параметр "Максимальное количество баз для обработки"...
                            if ( hConf.getItemValueInteger("DbMax")>0 ) {
                                //  Если достигнуто максимальное число обработанных баз, то прерываем дальнейшую обработку
                                if ( (++nDbProcessed)>=hConf.getItemValueInteger("DbMax") ) break;
                            }

                            //  Если берем в расчет параметр "Пауза, между запусками команды convert, в секундах"...
                            if ( hConf.getItemValueInteger("BetweenPause")>0 ) {
                                Thread.sleep(hConf.getItemValueInteger("BetweenPause") * 1000);
                            }


                        }
                    }

                    hDb=hDbDir.getNextDatabase();
                }

                //  Отсылка информационного сообщения...
                if ( hConf.getItemValue("Recipients").size()>0 ) hMsg.send();

            }

            System.out.println("UpdateTemplates  КОНЕЦ");


        } catch(Exception e) {
            e.printStackTrace();
        }
    }


    /*
     * Получить ссылку на конфигурационный документ агента
     *
     * Конфигурационный документ агента: Form="UpdateTemplateSettings" и Type="Configuration"
     * Представление Lookup\Configuration отбирает все документы с Type="Configuration". Индекс построен по полю Form.
     */
    Document getAgentConfig() {
        try {
            View hV=getSession().getCurrentDatabase().getView("Lookup\\Configuration");
            return hV.getDocumentByKey("UpdateTemplatesSettings");
        } catch (Exception e) {
            //return null;
            e.printStackTrace();
            return null;
        }
    }


    /*
     * Добавить документ логирования обновления шаблона базы
     *
     * Входные параметры:
     * 	String	Местоположение базы. Имя сервера
     * 	String	Местоположение базы. Имя файла
     * 	String	Описание(Title) базы
     * 	String	Текущее имя шаблона базы
     * 	String	Имя файла шаблона, который был использован для обновления шаблона базы
     */
    void addTemplateUpdateLog(String cServer, String cFilepath, String cTitle, String cTemplateName, String cNewTemplateFileName) {
        try {
            Document hDoc=getSession().getCurrentDatabase().createDocument();
            hDoc.appendItemValue("Form", "TemplateUpdateLog");
            hDoc.appendItemValue("Time", getSession().createDateTime(new Date()));
            hDoc.appendItemValue("Server", cServer).setNames(true); ;
            hDoc.appendItemValue("Filepath", cFilepath);
            hDoc.appendItemValue("Title", cTitle);
            hDoc.appendItemValue("TemplateName", cTemplateName);
            hDoc.appendItemValue("NewTemplateFileName", cNewTemplateFileName);
            hDoc.save();
        } catch (Exception e) {e.printStackTrace();}
    }


    //  Получить сокращенное имя
    String getAbbreviatedName(String cName){
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


    /*
     * Проверка имени файла(базы) на предмет принадлежности к заданной папке
     *
     * Вход:
     * 		String		имя папки; либо пустая строка(означает root, точнее корень notesdata) либо должна обязательно завершаться обратным слешем
     * 		String		имя файла(базы)
     *
     * Выход:
     * 		true/false
     *
     * Примеры:
     * 		"MailIn\\"	"MailIn\\a_ivanov.nsf"			true
     * 		"MailIn\\"	"MailIn\\logs\\_ivanov.nsf"		false
     * 		"mail\\"	"MailIn\\logs\\_ivanov.nsf"		false
     * 		""			"MailIn\\logs\\_ivanov.nsf"		false
     * 		""			a_ivanov.nsf					true
     *
     * Тестирование выполнено 05-07-2019
     *
     */
    boolean satisfyFilePath(String cFilePathTempl, String cFilePath){
        if ( cFilePath.toLowerCase().startsWith(cFilePathTempl.toLowerCase()) ) {
            if ( cFilePath.indexOf('\\', cFilePathTempl.length())==-1 ) {
                return true;
            }
        }
        return false;
    }


    /*
     *  Подготовить документ для информационного сообщения
     *
     *  Вход:
     *  	String[]		массив получателей сообщения
     *  	String			тема письма
     *  	String			заголовок письма(первая строка тела письма, выделанная жирным и отделенная пустой строкой от основного тела)
     *  	Document		ссылка на документ, на базе которого будет сформировано сообщение
     *  	RichTextItem	ссылка на поле Body документа
     *
     */
    void fillingMessage(Vector<String> aRecipient, String cSubject, String cTitle, Document hMsg, RichTextItem hBody) {

        try {
            Session hS=getSession();

            hMsg.setSaveMessageOnSend(false);
            hMsg.appendItemValue("Form", "Memo");
            hMsg.appendItemValue("SendTo", aRecipient).setNames(true);
            hMsg.appendItemValue("Principal", "UpdateTemplates agent");
            hMsg.appendItemValue("Subject", cSubject);

            //  Определям стиль для заголовка и добавляем заголовок в тело письма
            RichTextStyle hTextStyle=hS.createRichTextStyle();
            hTextStyle.setFontSize(12); hTextStyle.setBold(RichTextStyle.YES); hBody.appendStyle(hTextStyle);
            hBody.appendText(cTitle); hBody.addNewLine(); hBody.addNewLine();

            //  Определеяем стиль для всех остальных(информационных) строк письма
            hTextStyle.setFont(RichTextStyle.FONT_COURIER); hTextStyle.setFontSize(10); hTextStyle.setBold(RichTextStyle.NO); hBody.appendStyle(hTextStyle);

        } catch(Exception e) {
            e.printStackTrace();
        }

    }


}