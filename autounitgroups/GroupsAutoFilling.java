import java.util.Vector;
import java.io.FileWriter;
import java.util.HashSet;
import lotus.domino.*;

/**
 * Автомат наполнения групп по подразделениям
 */
public class GroupsAutoFilling extends AgentBase {

    //  Ссылка на адресную книгу
    private Database hNames;
    //  Ссылка на Список сотрудников
    private Database hStaff;


    //  Набор переменных, значения которых будут заполняться после верификации/парсинга очередной конфигурационной строки(фунция parsingGroupConfig())
    //  и использоваться далее в функционале наполнения группы(функция fillGroup())
    private String cGCUnitTypeItem;
    private String cGCUnitName;
    private String cGCPersonsType;
    private Document hGCGroup;


    public void NotesMain() {

        try {
            Session session = getSession();
            AgentContext agentContext = session.getAgentContext();

            Document hPDoc;
            DocumentCollection hPDocColl;
            Vector<String> aGroupConfig;
            boolean bParsingErr=false;
            boolean bFillErr=false;


            System.out.println("GroupAutoFilling  НАЧАЛО");


            //  Читаем профильный документ агента. Если документа нет, то на выход.
            hPDocColl=session.getCurrentDatabase().getProfileDocCollection("GroupsAutoFilling");
            if ( hPDocColl.getCount()>0 ) hPDoc=hPDocColl.getFirstDocument(); else return;

            aGroupConfig=hPDoc.getItemValue("GroupsConfig");
            if (aGroupConfig.size()>0) {


                //  Создаем документ, на базе которого будет проходить отправка информационного сообщения.
                //  Сразу же создается поле Body в этом документе, которое в процессе отработки основного функционала агента
                //  будет по надобности наполняться.
                Document hMsg=session.getCurrentDatabase().createDocument();
                RichTextItem hMsgBody=hMsg.createRichTextItem("Body");


                //  *** ПРОДУКТИВ ***
                //hNames=session.getDatabase("", "names.nsf");
                //  *** DEBUG ***
                hNames=session.getDatabase("CN=EMA/O=KIB", "names.nsf");

                //  Открываем Список сотрудников
                hStaff=session.getDatabase(hPDoc.getItemValueString("StaffServer"), hPDoc.getItemValueString("StaffDb"));



                //  Перебираем конфигурационные строки для автозаполнения групп
                for (int i=0; i<=aGroupConfig.size()-1; i++ ) {

                    System.out.println("GroupAutoFilling  Конфигурационная строка: " + aGroupConfig.elementAt(i));

                    //  Проверяем корректность очередной конфигурационной строки
                    if (parsingGroupConfig( aGroupConfig.elementAt(i) )){
                        //  Если проверка прошла нормально, но в переменных GC*(преффикс сокращенно GroupConfig) будут входные параметрыв для функции заливки группы
                        if (fillGroup()==false) {
                            System.out.println("GroupAutoFilling  Ошибка при наполнении группы:  " + hGCGroup.getItemValueString("Listname"));
                            hMsgBody.appendText("Ошибка при наполнении группы:  " + hGCGroup.getItemValueString("Listname")); hMsgBody.addNewLine();
                            if (bFillErr==false) bFillErr=true;
                        }

                    } else {
                        System.out.println("GroupAutoFilling  Ошибка парсинга/верификации конфигурационной строки:  " + aGroupConfig.elementAt(i));
                        hMsgBody.appendText("Ошибка парсинга/верификации конфигурационной строки:  " + aGroupConfig.elementAt(i)); hMsgBody.addNewLine();
                        if (bParsingErr==false) bParsingErr=true;
                    }

                }

                //  Отправка информационного сообщения об ошибках
                if (bParsingErr || bFillErr) sendMessage(hPDoc.getItemValue("Recipients"), "GroupsAutoFilling agent", "Ошибки", hMsg);

            }


            System.out.println("GroupAutoFilling  КОНЕЦ");


            //  *** DEBUG ***
            //System.out.println();







        } catch(Exception e) {
            e.printStackTrace();
        }
    }


    /*
     * Разбор конфигурационной строки для наполнения группы
     *
     * Вход
     * 		String		конфигурационная строка
     * Выход
     * 		Заполнение глобальных переменных:
     * 		cGCUnitTypeItem		имя поля в карточке Списка сотрудников, в котором храниться соответствующий тип подразделения
     * 		cGCUnitName			имя подразделения
     * 		cGCPersonsType		фиксированные значения: штатные/внештатные/штатные+внештатные
     * 		hGCGroup			ссылка на группу
     *
     * Конфигурационная строка представляет собой четыре элемента разделенные двоеточием:
     * 	Фиксированные значения: Блок/Департамент/Управление/Отдел/Группа
     * 	Фиксированные значения: Штатные/Внештатные/Штатные+Внештатные
     * 	Название подразделения
     * 	Имя существующей группы
     */
    boolean parsingGroupConfig(String cGroupConfig) {

        try {
            //  Бьем конфигурационную строку на части.
            //  Частей должно быть только четыре: Тип подразделения, Название подразделение, Кого брать(тип сотрудников), Название группы
            String[] aGroupConfigPart=cGroupConfig.split(":");
            if (aGroupConfigPart.length==4) {

                //  Вычисляем каким должно быть название поля из карточки БД Список сотрудников, в котором хранится имя
                //  соответствующего подразделения. В случае, если пользователь сделал ошибку в названии типа подразделения,
                //  то парсинг считается неуспешным.
                switch (aGroupConfigPart[0].trim().toLowerCase()) {
                    case "блок": cGCUnitTypeItem="BlockName"; break;
                    case "департамент": cGCUnitTypeItem="Officedir"; break;
                    case "управление": cGCUnitTypeItem="signOffice"; break;
                    case "отдел": cGCUnitTypeItem="OtdelName"; break;
                    case "группа": cGCUnitTypeItem="GroupName"; break;
                    default: return false;
                }

                //  Верификация типа пользователей, которые нужно включать в запрос.
                //  Аналогично, если пользователь сделал ошибку в названии, парсинг считается неуспешным
                switch (aGroupConfigPart[2].trim().toLowerCase()) {
                    case "штатные": cGCPersonsType="штатные"; break;
                    case "внештатные": cGCPersonsType="внештатные"; break;
                    case "штатные+внештатные": cGCPersonsType="штатные+внештатные"; break;
                    default: return false;
                }

                //  Проверяем наличие группы в адресной книге.
                //  В случае, если группа не существует, парсинг(верификация) считается неуспешной
                if ((hGCGroup=isGroup(hNames, aGroupConfigPart[3].trim()))==null) return false;

                //  Запоминаем имя структурного подразделения.
                cGCUnitName=aGroupConfigPart[1].trim();
                return true;
            }
            else return false;

        } catch (Exception e) {
            return false;
        }

    }




    /*
 		Функция заливки группы

 		Входными параметрами для данной функции, являются внешние переменные:
 			cGCUnitTypeItem		Имя поля из карточки БД Список сотрудников, в котором хранится название структурного подразделения
 			cGCUnitName			Имя структурного подразделения
 			cGCPersonsType		Тип пользователей, которые включить в выборку: "штатные"/"внештатные"/"штатные+внештатные"
 			hGCGroup			Ссылка на группу, которую нужно заливать

 		Выход:
 			false будет возвращен, если тип пользователя не соответствует ни одному из константных значений
 			false будет возвращен, если запись в группу физические невозможно (проблема 32К)
 			false будет возвращен, при сработке любого исключения

     */
    boolean fillGroup(){

        try{

            //  *** DEBUG ***
            //System.out.println(cGCUnitTypeItem + "  " + cGCUnitName + "  " + cGCPersonsType + "  " + hGCGroup.getItemValueString("Listname"));

            String cQuery;

            //  Формируем запрос, в зависимости от типа пользователей, которые нужно выбрать
            switch( cGCPersonsType  ) {
                case "штатные":
                    cQuery="SELECT Form=\"Sign\" & !(signCancelled=\"1\" | longOut=\"1\" | outofstaff=\"1\") & " + cGCUnitTypeItem + "=\"" + cGCUnitName + "\"";
                    break;
                case "внештатные":
                    cQuery="SELECT Form=\"Sign\" & outofstaff=\"1\" & !(signCancelled=\"1\" | longOut=\"1\") & " + cGCUnitTypeItem + "=\"" + cGCUnitName + "\"";
                    break;
                case "штатные+внештатные":
                    cQuery="SELECT Form=\"Sign\" & !(signCancelled=\"1\" | longOut=\"1\") & " + cGCUnitTypeItem + "=\"" + cGCUnitName + "\"";
                    break;
                default:
                    return false;
            }

            //  *** DEBUG ***
            //System.out.println(cQuery);


            //  Получаем коллекцию документов из Списка сотрудников
            DocumentCollection hStaffCardColl=hStaff.search(cQuery);
            //  *** DEBUG ***
            //System.out.println("Количество документов в коллекции из Списка сотрудников: " + hStaffCardColl.getCount());

            //  Массив, который будет наполнен только реально существующими в АК учетными записями. Имена учетных записей берутся из
            //  карточек Списка сотрудников. Последней фазой данного алгоритма может быть(а может и не быть) запись данного массива
            //  в поле Members заданной группы.
            Vector<String> aStaffAccount=new Vector<String>();
            //  Множество, в котором будут те-же учетки, что и в массиве aStaffAccount, но в нижнем регистре.
            //  Предназначено для сравнения с множеством, созданным на базе членов заданной группы
            HashSet<String> setStaffAccount=new HashSet<String>();


            //  *** DEBUG ***
            //FileWriter hDebugLog=new FileWriter("d:\\space\\out\\GroupsAutoFilling.txt", false);



            //  Перебираем документы коллекции, полученной из Списка сотрудников. Заполняем вектор и множество.
            //  Вектор нужен, чтобы произвести запись учетных записей в группу, в том виде, в котором они храняться в Списке сотрудников.
            //  Множество нужно для сравнения со множеством полученным из группы.
            Document hStaffCard=hStaffCardColl.getFirstDocument();
            while (hStaffCard!=null) {
                if (isRoutedAccount(hNames, hStaffCard.getItemValueString("signPersonRLat"))!=null) {
                    aStaffAccount.add( hStaffCard.getItemValueString("signPersonRLat") );
                    setStaffAccount.add( hStaffCard.getItemValueString("signPersonRLat").toLowerCase() );
                } else {

                    //  *** DEBUG ***
                    //hDebugLog.write(hStaffCard.getItemValueString("signPersonRLat") + "\n");

                }
                hStaffCard=hStaffCardColl.getNextDocument();
            }



            //  *** DEBUG ***
            //hDebugLog.close();



            //  Вычитываем членов группы и размещаем их в множестве. Предварительно преобразовываем все имена в нижний регистр.
            HashSet<String> setMember=new HashSet<String>();
            addLowerCaseVectorToSet(hGCGroup.getItemValue("Members"), setMember);


            //  *** DEBUG ***
            //System.out.println("Количество элементов в множестве Списка сотрудников: " + setStaffAccount.size());
            //System.out.println("Количество элементов в множестве Группы: " + setMember.size());


            //  Сравниваем два множества: полученное из Списка сотрудников и полученное из группы.
            //  Если множества не эквивалентны, то делаем перезапись содержимого группы, массивом из Списка сотрудников.
            if ( setMember.equals( setStaffAccount )==false ) {

                //  *** DEBUG ***
                //System.out.println("Множество из группы НЕ ЭКВИВАЛЕНТНО множеству из Списка сотрдников!");


                //  Перезапись содержимого группы
                hGCGroup.replaceItemValue("Members", aStaffAccount).setNames(true);
                hGCGroup.save(true);

                //  Контрольная вычитка содержимого группы. Данная вычитка необходима вследствии того, что java-агенты
                //  не реагируют никаким образом на проблему 32K. Загоняем содержимое групппы в нижнем регистре в множество.
                addLowerCaseVectorToSet(hGCGroup.getItemValue("Members"), setMember);

                //  *** DEBUG ***
                //System.out.println("Повторная вычитка группы  Количество элементов в множестве: " + setMember.size());

                //  Сравниваем текущее содержимое группы(уже после записи!) и то, что действительно нужно было записать - множество из Списка сотрудников
                if ( setMember.equals( setStaffAccount )==false ) {
                    //  *** DEBUG ***
                    //System.out.println("Повторная вычитка группы  ТЕКУЩЕЕ множество из группы не эквивалентно множеству из Списка сотрудников!");
                    return false;
                }

            } else {

                //  *** DEBUG ***
                //System.out.println("Множества группы и Списка сотрудников эквивалентны. Перезапись содержимого группы не нужна!");

            }

            return true;



        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }


    }


    /*
     * На основе элементов вектора String сформировать идентичное множество, но уже в нижнем регистре
     *
     * Вход:
     * 		Vector<String>		вектор
     * 		HashSet<String>		множество, которое сначала очищается, а потом наполняется значениями из вектора
     */
    private void addLowerCaseVectorToSet(Vector<String> v, HashSet<String> s) {
        s.clear();
        for (int i=0; i<=v.size()-1; i++) s.add( v.elementAt(i).toLowerCase());
    }




    /*
     * Конвертировать название типа структурного подразделения в соответствующее название поля
     * из карточки Списка сотрудников
     *
     * Вход:
     * 	String	Тип структурного подразделения
     * Выход:
     * 	String	Название поля либо пустая строка
     */
    private String convertUnitTypeToItem(String cUnitType){
        try{
            switch (cUnitType.toLowerCase()) {
                case "блок": return "BlockName";
                case "департамент": return "Officedir";
                case "управление": return "signOffice";
                case "отдел": return "OtdelName";
                case "группа": return "GroupName";
                default: return "";
            }
        }catch(Exception e){
            return "";
        }
    }


    /*
     * Существует ли группа?
     * Входные параметры:
     * 		Database		ссылка на адресную книгу
     * 		String			имя искомой группы
     * Выход:
     * 		сслылка на найденную группу или null
     */
    private Document isGroup(Database hNames, String cGroup) {
        try {
            View hGroupVw=hNames.getView("$RegisterGroups");
            if (hGroupVw==null) return null; else return hGroupVw.getDocumentByKey( getAbbreviatedName(cGroup) );
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }


    //  Получить сокращенное имя
    private String getAbbreviatedName(String cName){
        try{
            Session hSession=getSession();
            Name hN=hSession.createName(cName);
            return hN.getAbbreviated();
        }catch (Exception e){
            return "";
        }
    }


    /*
     * Существует ли учетная запись, на которую можно маршрутизировать почту, с зданным именем?
     * Входные параметры:
     * 		hNames			ссылка на адресную книгу
     * 		cAccount		имя искомой учетной записи
     * Выход:
     * 		ссылка на найденную учетную запись либо null
     *
     * Тестирование проведено
     */
    private Document isRoutedAccount(Database hNames, String cAccount) {
        try {
            View hUsersVw=hNames.getView("$Users");
            if (hUsersVw!=null)	return hUsersVw.getDocumentByKey(cAccount, true); else return null;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }



    /*
     *  Отправить информационного сообщение
     *
     *  Вход:
     *  	String[]		массив получателей сообщения
     *  	String			тема письма
     *  	Document		ссылка на документ, на базе которого будет сформировано и отправлено сообщение
     *
     *  Отличительной особенностью данной функции является то, что в функцию передается документ, на базе которого идет отсылка.
     *  Т.е. поле B
     *
     */
    private void sendMessage(Vector<String> aRecipient, String cPrincipal, String cSubject, Document hMsg) {

        try {
            if (aRecipient.size()>0) {
                hMsg.setSaveMessageOnSend(false);
                hMsg.appendItemValue("Form", "Memo");
                hMsg.appendItemValue("SendTo", aRecipient).setNames(true);
                hMsg.appendItemValue("Principal", cPrincipal);
                hMsg.appendItemValue("Subject", cSubject);
                hMsg.send();
            }
        } catch(Exception e) {
            e.printStackTrace();
        }

    }

}