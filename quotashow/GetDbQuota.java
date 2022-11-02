/*
 * Входные параметры агента:
 * 		String		PersonUNID		UNID документа Person из адресной книги, квоту по которому необходимо узнать
 *
 * Входные параметры агента:
 * 		Fullname		нотес-имя пользователя
 * 		MailServer		почтовый сервер(текстовый список)
 * 		MailFile		почтовый файл(текстовый список)
 * 		Size			размер почтового файла(текстовый список)
 * 		Quota			квота почтового файла(текстовый список)
 * 		Treshold		порог квоты почтового файла(текстовый список)
 *
 * Примечание:
 * Все текствые списки синхронны. Т.е. если у пользователя две реплики, все текствоые списки
 * будут содержать по два синхронных элемента
 *
 */
import java.util.Vector;

import lotus.domino.*;

public class GetDbQuota extends AgentBase {


    static final String TEMP_AGENT_PARAMETERS_DB="temp/getquotaparams.nsf";


    //  Ссылка на документ учетной записи Person, по которому нужно вывести всю информацию о квоте его базы
    Document hPerson;

    public void NotesMain() {

        try {

            // *** DEBUG ***
            //Thread.sleep(20000);

            Session session = getSession();
            AgentContext agentContext = session.getAgentContext();
            String cParamNoteID;      //  NoteID, переданное в агент в качестве параметра RunOnServer()
            String cAgentServer;		//  сервер, на котором запускается данный агент
            Database hTempDb;			//  ссылка на временную служебную базу для работы с параметрами агента
            Document hParamDoc;		//  ссылка на документ во временной базе, через который передаются параметры в агент и результаты из агента
            Database hNames;			//  ссылка на адресную книгу
            Document hPerson;			//  ссылка на документ Person из адресной книги
            Database hPersonDb;		//  ссылка на почтовую базу прописанную в документе Person
            String cPersonDbReplicaID;
            Vector<String> aClusterNeighbours;		//  соседи по кластеру того сервера, который прописан в учетной записи пользователя
            String cMailServer, cMailFile;
            int i;
            Database hPersonDbReplica;



            //  Получили переданный в агент NoteID
            cParamNoteID=agentContext.getCurrentAgent().getParameterDocID();		//  *** Продуктив ****
            //cParamNoteID="FFFF0001";   											//  *** DEBUG Код для создания временной базы параметров ***
            //cParamNoteID="000008F6";


            //  *** DEBUG ***
            System.out.println("cParamNoteID: " + cParamNoteID);   //  *** DEBUG ***



            //  Имя сервера на котором выполняется агент
            cAgentServer=agentContext.getCurrentDatabase().getServer();


            //  Через параметр передана команда агенту создать временную базу для параметров:
            if ( cParamNoteID.equals("FFFF0001") ) {
                if ( isDb(cAgentServer, TEMP_AGENT_PARAMETERS_DB)==false ) {
                    hTempDb=session.getDbDirectory(cAgentServer).createDatabase(TEMP_AGENT_PARAMETERS_DB, true);
                    if (hTempDb.isOpen()) {
                        hTempDb.grantAccess("LocalDomainServers", ACL.LEVEL_MANAGER);
                        hTempDb.grantAccess("LocalDomainAdmins", ACL.LEVEL_MANAGER);
                        hTempDb.grantAccess("-Default-", ACL.LEVEL_EDITOR);
                        println("Создана временная база для передачи параметров: " + TEMP_AGENT_PARAMETERS_DB);
                    }
                }
                return;
            }


            //  Открываем временную базу через которую осуществляется передача параметров агенту
            hTempDb=session.getDatabase(cAgentServer, TEMP_AGENT_PARAMETERS_DB);
            if (hTempDb.isOpen()==false) {
                println("Проблемы при открытии временной базы " + TEMP_AGENT_PARAMETERS_DB);
                return;
            }

            //  Получаем по NoteID документ,через который передаются параметры для агента
            hParamDoc=hTempDb.getDocumentByID(cParamNoteID);

            //  Открываем адресную книгу и ищем документ Person, UNID которого передан в качестве параметра агента
            hNames=session.getDatabase(cAgentServer, "names.nsf");
            //  Если по каким-то невероятным причинам поиск будет неуспешен, произойдет сработка исключения
            hPerson=hNames.getDocumentByUNID(hParamDoc.getItemValueString("PersonUNID"));

            //  Почтовый сервер и почтовая база, прописанные в документе Person
            cMailServer=hPerson.getItemValueString("MailServer");
            cMailFile=hPerson.getItemValueString("MailFile");
            //  Только если поля Сервера и Почтовой базы заполенены в документе Person, алгоритм пойдет дальше
            if (!(cMailServer.trim().isEmpty()==false && cMailFile.trim().isEmpty()==false )) return;

            //  Сохраняем значение имени учетной записи пользователя в документе-параметре(уже на выход!)
            hParamDoc.appendItemValue("Fullname", hPerson.getItemValueString("Fullname") ).setNames(true);;

            //  Открывам почтовую базу прописанную в учетной записи Person
            hPersonDb=session.getDatabase(cMailServer, cMailFile, false);
            if ( hPersonDb==null ) {
                println("Проблемы при открытии базы, прописанной в документе Person");
                return;
            }

            //  Добавляем информацию о базе, которая прописана в документе Person в
            //  документ-параметр. Уже в качестве исходящих(результирующих) данных, которые поставляет агент
            appendDbInfo(hParamDoc, hPersonDb);


            //  Получаем соседей по кластеру почтового сервера, который прописан в документе Person
            aClusterNeighbours=getClusterNeighbours(hNames, cMailServer);

            //  Если эти соседи существуют, то на каждом сервере-соседе ищем реплику и сохраняем
            //  информацию о ней в документе-параметре
            if ( aClusterNeighbours.size()>0 ) {

                cPersonDbReplicaID=hPersonDb.getReplicaID();
                for (i=0; i<=aClusterNeighbours.size()-1; i++) {

                    hPersonDbReplica=session.getDatabase(null, null);
                    if ( hPersonDbReplica.openByReplicaID(aClusterNeighbours.get(i), cPersonDbReplicaID) ) {

                        appendDbInfo(hParamDoc, hPersonDbReplica);

                    }


                }

            }

            hParamDoc.save();


        } catch(Exception e) {
            e.printStackTrace();
        }
    }


    /*
     * Cуществует ли база?
     *
     * Вход:
     * 		String		имя сервера
     * 		String		имя базы
     * Выход:
     * 		true/false
     *
     * Тестирование проведено
     */
    private boolean isDb(String cServerName, String cDbName) throws Exception {
        try {
            getSession().getDbDirectory( cServerName ).openDatabase(cDbName, false);
            return true;
        } catch (NotesException e) {
            if ( e.id==NotesError.NOTES_ERR_SYS_FILE_NOT_FOUND ) return false; else throw e;
        } catch (Exception e) {
            throw e;
        }
    }



    /*
     * Вывести строку в лог
     */
    private void println(String cS) {
        try {
            System.out.println(getSession().getAgentContext().getCurrentAgent().getName() +": " + cS  );
        } catch (Exception e) {

        }
    }


    /*
     * Получить имена всех серверов кластера, к которому принадлежит и заданный сервер
     *
     * Вход:
     * 		Database	ссылка на адресную книгу
     * 		String		имя известного сервера
     * Выход:
     * 		массив Vector
     *
     * Тестирование проведено 04-06-2019
     */
    private Vector<String> getClusterServerNames(Database hNames, String cServer) {
        try {

            View hServersLookupVw, hClustersVw;
            Document hServer;
            String cCluster;
            DocumentCollection hClusterServersColl;
            int i;

            //  На выход по умолчанию...
            Vector<String> aClusterServerName=new Vector<String>();

            //  Ищем документ сервера, с тем чтобы узнать к какому кластеру принадлежит этот сервер
            hServersLookupVw=hNames.getView("$ServersLookup");
            hServer=hServersLookupVw.getDocumentByKey(cServer, true);
            if (hServer==null) return aClusterServerName;

            //  Получаем имя кластера из серверного документа...
            cCluster=hServer.getItemValueString("ClusterName");

            //  Получаем коллекцию серверных документов заданного кластера...
            hClustersVw=hNames.getView("$Clusters");
            hClusterServersColl=hClustersVw.getAllDocumentsByKey(cCluster, true);

            //  Перебираем найденную коллекцию и сохраняем имена серверов в массиве
            for (i=1; i<=hClusterServersColl.getCount(); i++) {
                aClusterServerName.add( hClusterServersColl.getNthDocument(i).getItemValueString("MailServer") );
            }

            return aClusterServerName;
        } catch (Exception e) {
            return null;
        }
    }



    /*
     * Получить для заданного сервера все его сервера-соседи по кластеру
     *
     * Вход:
     * 		Database	адресная книга
     * 		String		имя сервера, соседи которого ищутся
     * Выход:
     * 		массив Vector соседей по кластеру или null
     */
    private Vector<String> getClusterNeighbours(Database hNames, String cServer) {
        try {

            View hServersLookupVw, hClustersVw;
            Document hServer;
            String cCluster;
            DocumentCollection hClusterServersColl;
            int i;

            //  На выход по умолчанию...
            Vector<String> aClusterServerName=new Vector<String>();

            //  Ищем документ сервера, с тем чтобы узнать к какому кластеру принадлежит этот сервер
            hServersLookupVw=hNames.getView("$ServersLookup");
            hServer=hServersLookupVw.getDocumentByKey(cServer, true);
            if (hServer==null) return aClusterServerName;

            //  Получаем имя кластера из серверного документа...
            cCluster=hServer.getItemValueString("ClusterName");

            //  Получаем коллекцию серверных документов заданного кластера...
            hClustersVw=hNames.getView("$Clusters");
            hClusterServersColl=hClustersVw.getAllDocumentsByKey(cCluster, true);

            //  Перебираем найденную коллекцию и сохраняем имена серверов в массиве. Всех серверов, кроме того
            //  который передан в функцию.
            for (i=1; i<=hClusterServersColl.getCount(); i++) {

                if ( hClusterServersColl.getNthDocument(i).getItemValueString("MailServer").equalsIgnoreCase(cServer) ) continue;
                else aClusterServerName.add( hClusterServersColl.getNthDocument(i).getItemValueString("MailServer") );
            }

            return aClusterServerName;
        } catch (Exception e) {
            return null;
        }
    }



    /*
     * Добавить в документ информацию по базе
     *
     * Вход:
     * 		Document		документ, в который будет добавлена информация по базе
     * 		Database		база, информация по которой будет добавляться в документ
     *
     * Примечание:
     * Следующая информация будет добавляться в документ:
     * 		Сервер базы						поле MailServer
     * 		Имя файла базы					поле MailFile
     * 		Размер базы						поле Size
     * 		Квота на базе					поле Quota
     * 		Порог предупреждения на базе	поле Threshold
     * Все поля текстовые
     */
    private void appendDbInfo(Document hDoc, Database hDb) {

        try {


            Item hItem;

            //  Сервер
            if ( hDoc.hasItem("MailServer")==false ) {
                hDoc.appendItemValue("MailServer", getCanonicalName(hDb.getServer())).setNames(true);

            } else {

                hItem=hDoc.getFirstItem("MailServer");
                hItem.appendToTextList(getCanonicalName(hDb.getServer()));
            }

            //  Имя файла базы
            if ( hDoc.hasItem("MailFile")==false ) {
                hDoc.appendItemValue("MailFile", hDb.getFilePath() );

            } else {

                hItem=hDoc.getFirstItem("MailFile");
                hItem.appendToTextList(hDb.getFilePath());
            }


            //  Размер базы
            //  Метод Database.getSize() отдает байты в переменной типа double
            if ( hDoc.hasItem("Size")==false ) {
                hDoc.appendItemValue( "Size", hDb.getSize() );
            } else {
                Vector<Double> aSize=hDoc.getItemValue("Size");
                aSize.add(hDb.getSize());
                hDoc.replaceItemValue("Size", aSize);
            }

            //  Квота
            //  Метод Database.getSizeQuota() отдает килобайты
            if ( hDoc.hasItem("Quota")==false ) {
                hDoc.appendItemValue( "Quota", hDb.getSizeQuota() );
            } else {
                Vector<Double> aQuota=hDoc.getItemValue("Quota");
                aQuota.add( (new Integer(hDb.getSizeQuota())).doubleValue() );
                hDoc.replaceItemValue("Quota", aQuota);
            }

            //  Порог предупреждения
            //  Метод Database.getSizeWarning() отдает килобайты
            if ( hDoc.hasItem("Threshold")==false ) {
                hDoc.appendItemValue( "Threshold", hDb.getSizeWarning() );
            } else {
                Vector<Double> aThreshold=hDoc.getItemValue("Threshold");
                aThreshold.add( (new Long(hDb.getSizeWarning())).doubleValue() );
                hDoc.replaceItemValue("Threshold", aThreshold);
            }





    		/*
    		//  Размер базы
    		//  Метод Database.getSize() отдает байты в переменной типа double
    		if ( hDoc.hasItem("Size")==false ) {
    			hDoc.appendItemValue( "Size", Long.toString(new Double(hDb.getSize()).longValue()) );

    		} else {

    			hItem=hDoc.getFirstItem("Size");
    			hItem.appendToTextList( Long.toString(new Double(hDb.getSize()).longValue()) );
    		}

    		//  Квота
    		//  Метод Database.getSizeQuota() отдает килобайты
    		if ( hDoc.hasItem("Quota")==false ) {
    			hDoc.appendItemValue( "Quota", Integer.toString(hDb.getSizeQuota()) );

    		} else {

    			hItem=hDoc.getFirstItem("Quota");
    			hItem.appendToTextList( Integer.toString(hDb.getSizeQuota()) );
    		}

    		//  Порог предупреждения
    		//  Метод Database.getSizeWarning() отдает килобайты
    		if ( hDoc.hasItem("Threshold")==false ) {
    			hDoc.appendItemValue( "Threshold", Long.toString(hDb.getSizeWarning()) );

    		} else {

    			hItem=hDoc.getFirstItem("Threshold");
    			hItem.appendToTextList( Long.toString(hDb.getSizeWarning()) );
    		}
    		*/




        } catch (Exception e) {

        }
    }


    /*
     * Получить каноническое имя
     *
     * Ivan I Ivanov/KIB		->	CN=Ivan I Ivanov/O=KIB
     * CN=Ivan I Ivanov/O=KIB	->	CN=Ivan I Ivanov/O=KIB
     * Ivan I Ivanov			->	Ivan I Ivanov
     */
    private String getCanonicalName(String cName){
        try{
            return getSession().createName(cName).getCanonical();
        } catch (Exception e) {
            return "";
        }
    }


}