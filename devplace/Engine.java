import java.util.*;
import lotus.domino.*;

public class Engine {


    /*
     *   Существует ли имя в адресной книге?
     *   Входные параметры:
     *   	hNames	ссылка на АК
     *   	cName	имя, поиск которого будет осуществляться в АК
     *   Выход:
     *   	true/false
     */
    static boolean IsPerson(Database hNames, String cName ){
        try {
            View hPeople;
            Vector aKey;
            Document hDoc;
            boolean bResult;


            bResult=false;
            if (hNames.isOpen()) {
                if (cName.length()>0){

                    //  Получаем ссылку на представление $People
                    hPeople=hNames.getView("$People");

                    //  Формируем массив из двух ключей для поиска в представлении $People
                    aKey=new Vector(2);
                    aKey.add(0, NameActions.GetLastName(cName).substring(0,1).toUpperCase() );		//  первая буква Lastname в верхнем регистре
                    aKey.add(1, NameActions.GetCanonicalName(cName));																//  Fullname

                    //  Поиск с точным совпадением
                    hDoc=hPeople.getDocumentByKey(aKey, true);
                    if (hDoc==null) bResult=false; else bResult=true;
                }
            }
            return bResult;

        } catch(Exception e){
            return false;
        }
    }


    /*
     * Взять из поля типа DateTime значение и отдать только дату в виде строки.
     * Входные параметры:
     * 		hDoc		ссылка на документ
     * 		cFieldName	имя поля типа DateTime
     * Выход:
     * 		строка с датой либо пустая строка
     */
    static String GetDateAsString(Document hDoc, String cFieldName){
        try{
            DateTime d;
            d=(DateTime)hDoc.getItemValue(cFieldName).elementAt(0);
            return d.getDateOnly();
        } catch(Exception e){
            return "";
        }
    }


    /*
     * Существует ли общий почтовый ящик с заданным именем?
     * Входные параметры:
     * 		hNmaes		ссылка на адресную книгу
     * 		cName		имя искомого общего почтового ящика
     * Выход:
     * 		true/false	существует/не существует
     */
    static boolean IsDatabase(Database hNames, String cName){
        try{
            View hMailInView;
            Vector aKey;
            boolean bResult;


            bResult=false;
            if (hNames.isOpen()){
                if (cName.length()>0){

                    hMailInView=hNames.getView("Mail-In Databases");
                    if (hMailInView!=null){
                        //  Формируем массив из двух ключей для поиска в представлении "Mail-In Databases"
                        aKey=new Vector(2);
                        aKey.add(0, "Databases");
                        aKey.add(1, NameActions.GetAbbreviatedName(cName));

                        //  Поиск с точным совпадением...
                        if (hMailInView.getDocumentByKey(aKey, true)!=null) bResult=true;
                    }

                }
            }
            return bResult;

        } catch(Exception e){
            return false;
        }

    }


    /*
     * Является ли учетная запись типа Person уволенной или длительно-отсутствующей?
     * Входные параметры:
     * 		in	hStaff			ссылка на Список сотрудников
     * 		in	cPerson			имя учетной записи
     * 		out	dt				дата увольнения/дата выхода в длит.отсутствие
     * Выход:
     * 		true/false			является/не является
     */
    static boolean IsCancelledOrLongOut(Database hStaff, String cPerson, DateTime dt){
        try{

            String cQuery;
            DocumentCollection hDColl;


            if (hStaff.isOpen()) {
                // Выбираем все записи, касающиеся данного сотрудника из Списка сотрудников
                cQuery="SELECT Form=\"Sign\":\"Sign_cnd\" & @LowerCase(@Name([ABBREVIATE]; signPersonRLat))=\""+NameActions.GetAbbreviatedName(cPerson).toLowerCase()+"\"";

                hDColl=hStaff.search(cQuery);
                if (hDColl.getCount()>0) {
                    return IsCancelledOrLongOut_Collection(hDColl, dt);
                }
            }

            return false;
        } catch (Exception e){
            return false;
        }
    }


    /*
     * Есть коллекция документов из Списка сотрудников. Предполагается, что все документы этой коллекции относяться к одному
     * сотруднику. Является ли данный сотрудник уволенным или длительно-отсутствующим?
     *
     * Входные параметры:
     * 	in	hDColl		ссылка на коллекцию документов
     * 	out	dt			если функция возращает true, то переменная будет хранить наиболее актуальную дату либо увольнения либо выхода в длит.отсутствие
     *
     * Выход:
     * 	true/false		является/не является
     */
    static boolean IsCancelledOrLongOut_Collection(DocumentCollection hDColl, DateTime dt) {
        try{

            Document hDoc, hTempDoc;
            DateTime dtDoc, dtTempDoc;
            boolean bDocCancelled, bDocLongout, bTempDocCancelled;


            //  Дань законам Java - любая переменная перед ее использованием должна быть инициализирована
            hTempDoc=null;
            bTempDocCancelled=false;

            hDoc=hDColl.getFirstDocument();
            while (hDoc!=null) {

                //  Признак уволенности для текущего документа
                bDocCancelled=hDoc.getItemValueString("signCancelled").equals("1");
                //  Признак длит.отсут. для текущего документа
                bDocLongout=hDoc.getItemValueString("longOut").equals("1");


                if (bDocCancelled==false && bDocLongout==false) return false;
                else {

                    if (hTempDoc==null) {
                        //  Итак, первый раз встречен документ, который является либо уволенным либо длите.отсут.
                        //  Запоминаем ссылку на этот документ...
                        hTempDoc=hDoc;
                        bTempDocCancelled=bDocCancelled;
                    } else {
                        //  Документ, который является или уволенным или длит.отсут. встречен уже не в первый раз.
                        //  Необходимо выбрать, какой документ является "более уволенным или длит.отстут", т.е. у которого дата увольнения или выхода
                        //  в длит.отсутствие более близка к сегодняшнему дню.

                        //  Получение либо даты увольнения либо даты выхода в длит.отсут. для текущего документа
                        if (bDocCancelled) dtDoc=(DateTime) hDoc.getItemValueDateTimeArray("signDateClose").elementAt(0);
                        else dtDoc=(DateTime) hDoc.getItemValueDateTimeArray("longoutDateFrom").elementAt(0);

                        //  Получение либо даты увольнения либо даты выхода в длит.отсут. для временного документа
                        if (bTempDocCancelled) dtTempDoc=(DateTime) hTempDoc.getItemValueDateTimeArray("signDateClose").elementAt(0);
                        else dtTempDoc=(DateTime) hTempDoc.getItemValueDateTimeArray("longoutDateFrom").elementAt(0);

                        //  Если дата текущего документа больше даты временного документа...
                        if (dtDoc.timeDifference(dtTempDoc)>0) {
                            hTempDoc=hDoc;
                            bTempDocCancelled=bDocCancelled;
                        }
                    }

                }

                hDoc=hDColl.getNextDocument();
            }

            /*
             * В этой точке алгоритма, могут быть только два варианта:
             * 		1. в коллекции не было ни одного документа и, соответственно, переменная hTempDoc осталась равной null
             * 		2. вся коллекция состоит только из документов уволенных и длит.отсут. и в переменной hTempDoc храниться документ, наиболее близкий к сегодня
             */
            if (hTempDoc!=null) {
                //  Прописываем дату непосредственно в объекте, ссылка на который передана в функцию, чтобы дата была передана в вызывающую процедуру
                if (bTempDocCancelled) dt.setLocalTime( ((DateTime)hTempDoc.getItemValueDateTimeArray("signDateClose").elementAt(0)).toJavaDate() );
                else dt.setLocalTime( ((DateTime)hTempDoc.getItemValueDateTimeArray("longoutDateFrom").elementAt(0)).toJavaDate() );
                return true;
            }

            return false;
        } catch (Exception e){
            e.printStackTrace();
            return false;
        }
    }


    /*
     * Получить разницу между двумя датами в днях(неважно какая дата больше или меньше, нужна только разница в днях)
     * Входные параметры:
     * 	dt1		первая дата
     * 	dt2		вторая дата
     * Выход:
     * 	Разница в днях между первой и второй датами
     */
    static int DayDiffrence(DateTime dt1, DateTime dt2) {
        try {
            return (int)(Math.abs( (float)dt1.timeDifference(dt2) / (float)86400  ));
        } catch (Exception e) {
            e.printStackTrace();
            return 0;
        }
    }




    /*
     * Является ли сотрудник длительно отсутствующим? (анализ коллекции документов из Списка сотрудников)
     *
     * Входные параметры:
     * 		in	hDocCollection	коллекция документов относящаяяся к заданному сотруднику ииз Списка сотрудников
     * 		out	dtLongOutFrom	если функция возвращает true, то дата выхода в длительное отсутствие
     * Выход:
     * 		true/false			является/не является
     *
     * ТРЕБУЕТСЯ ТЕСТИРОВНИЕ
     */
	/*
	static boolean IsLongOut_Collection(DocumentCollection hDocCollection, DateTime dtLongOutFrom){
		try{
			Document hDoc;
			Document hTempDoc;
			boolean bLongOut, bOtherPerson;
			DateTime dtTempDoc, dtDoc;
			boolean bFnRes;

			// --- DEBUG ---
			Debug dbg=new Debug("c:\\space\\out\\Test_IsLongOut_Collection.txt");


			//  Возврат функции по умолчанию...
			bFnRes=false;

			//  Признак того, что среди документов коллекции найден хотя-бы один длительно-отсутствующий
			bLongOut=false;
			//  Признак того, что найден хотя бы один не длительно-отсутствующий и не уволенный, т.е. теоретически "живой" сотрудник
			bOtherPerson=false;

			hDoc=hDocCollection.getFirstDocument();
			//  Java требует явной инициализации переменной
			hTempDoc=hDoc;

			while (hDoc!=null){

				// --- DEBUG ---
				dbg.writeln("hasItem(longOut): " + hDoc.hasItem("longOut"));
				dbg.writeln("getItemValueString(longOut): "+hDoc.getItemValueString("longOut"));
				dbg.writeln("getItemValueString(longOut)==\"1\": " + (hDoc.getItemValueString("longOut")=="1"));
				dbg.writeln("getItemValueString(longOut).equals(\"1\"): " + hDoc.getItemValueString("longOut").equals("1"));

				if (hDoc.hasItem("longOut") && hDoc.getItemValueString("longOut").equals("1")){

					if (bLongOut==false){
						hTempDoc=hDoc;
						bLongOut=true;
					}
					else{
						dtTempDoc=(DateTime)(hTempDoc.getItemValueDateTimeArray("longoutDateFrom").elementAt(0));
						dtDoc=(DateTime)(hDoc.getItemValueDateTimeArray("longoutDateFrom").elementAt(0));
						//  Если дата выхода в длит.отсут. из текущиго документа больше той, которая храниться во временном  документе...
						if (dtDoc.timeDifference(dtTempDoc)>0) {
							hTempDoc=hDoc;
						}
					}
				}
				else{
					//  Сотрудник не длительного отсутствующий...
					//  Любой вид сотрудника в этой ветке алгоритма, кроме уволенного, будет следствием False на выходе функции.
					//  "Уволенные" записи игнорируются и цикл продолжается
					if (!(hDoc.hasItem("signCancelled") & hDoc.getItemValueString("signCancelled")=="1")) {
						bOtherPerson=true;
						break;
					}
				}

				hDoc=hDocCollection.getNextDocument();
			}

			//  Ниже единственный маршрут при котором функция вернет true
			if (bOtherPerson==false){
				if (bLongOut){

					// неверное изменение значения переменной типа DateTime
					// dtLongOutFrom=(DateTime)(hTempDoc.getItemValueDateTimeArray("longoutDateFrom").elementAt(0));

					// неверное изменение значения переменной типа DateTime
					// Session hSession;
					// hSession=NotesFactory.createSession();
					// dtLongOutFrom=hSession.createDateTime(  ((DateTime)hTempDoc.getItemValueDateTimeArray("longoutDateFrom").elementAt(0)).toJavaDate()  );

					dtLongOutFrom.setLocalTime( ((DateTime)hTempDoc.getItemValueDateTimeArray("longoutDateFrom").elementAt(0)).toJavaDate() );

					bFnRes=true;
				}
			}

			return bFnRes;
		} catch (Exception e){
			e.printStackTrace();
			return false;
		}
	}
	*/

    /*
     * Является ли сотрудник длительно отсутствующим? (анализ коллекции документов из Списка сотрудников)
     *
     * Входные параметры:
     * 		in	hDocCollection	коллекция документов относящаяяся к заданному сотруднику ииз Списка сотрудников
     * 		out	dtLongOutFrom	если функция возвращает true, то дата выхода в длительное отсутствие
     * Выход:
     * 		true/false			является/не является
     *
     * ТРЕБУЕТСЯ ТЕСТИРОВНИЕ
     */
	/*
	static boolean IsLongOut_Collection(DocumentCollection hDocCollection, DateTime dtLongOutFrom){
		try{
			Document hDoc, hTempDoc;
			boolean bCancelledOrLongout;
			DateTime dtDoc, dtTempDoc;
			boolean bCancelled, bLongout;


			//  Признак того, что найдена по крайней мере одна запись, с установленным признаком уволенности или длит.отсутствия
			bCancelledOrLongout=false;

			hDoc=hDocCollection.getFirstDocument();
			while (hDoc!=null){

				//  Признак уволенности для текущего документа
				bCancelled=hDoc.getItemValueString("signCancelled").equals("1");
				//  Признак длит.отсут. для текущего документа
				bLongout=hDoc.getItemValueString("longOut").equals("1");

				if (bCancelled==false && bLongout==false) {
					return false;
				} else {
					//  В данной точке кода, сотрудник либо уволенный либо длит.отстутствующий
					if (bCancelledOrLongout==false){
						bCancelledOrLongout=true;
						hTempDoc=hDoc;
					}else{
						//  Получение либо даты увольнения либо даты выхода в длит.отсут. для текущего документа
						if (bCancelled) dtDoc=(DateTime) hDoc.getItemValueDateTimeArray("signDateClose").elementAt(0);
						 else dtDoc=(DateTime) hDoc.getItemValueDateTimeArray("longoutDateFrom").elementAt(0);


					}
				}

				hDoc=hDocCollection.getNextDocument();
			}

			return true;
		} catch (Exception e){
			//e.printStackTrace();
			return false;
		}
	}
	*/



    /*
     * Является ли сотрудник уволенным? (анализ коллекции документов из Списка сотрудников)
     * Входные параметры:
     * 		in	hDocCollection	коллекция документов из Списка сотрудников по пользователю
     * 		out	dtClose			если функция возвращает true, то дата выхода в длительное отсутствие
     * Выход:
     * 		true/false			является/не является
     *
     * ТРЕБУЕТСЯ ТЕСТИРОВНИЕ
     */
    static boolean IsCancelled_Collection(DocumentCollection hDocCollection, DateTime dtClose){
        try{
            Document hDoc;
            Document hTempDoc;
            boolean bCancelled, bOtherPerson;
            DateTime dtDoc, dtTempDoc;
            boolean bFnRes;

            //  По умолчанию
            bFnRes=false;

            //  Признак того, что в коллекции найден хотя-бы один уволенный
            bCancelled=false;

            //  Признак того, что в коллекции обнаружен по крайней мере один сотрудник не являющийся уволенным(каким он является - значения не имеет)
            bOtherPerson=false;

            hDoc=hDocCollection.getFirstDocument();

            //  Дань правилам Java - невозможности использовать неинициализированную переменную.
            //  Я не хочу создавать для этой переменной пустой объект Document, чтобы только инициализировать ее. Ведь я
            //  планирую использовать данную переменную только как ссылку на уже существующий объект.
            hTempDoc=hDoc;

            while (hDoc!=null){

                if (hDoc.hasItem("signCancelled") & hDoc.getItemValueString("signCancelled")=="1"){
                    if (bCancelled){
                        //  Имеем первого уволенного в коллекции. Фиксируем это изменяя флаг на true
                        bCancelled=true;
                        hTempDoc=hDoc;
                    }
                    else{
                        //  Очередной уволенный в коллекции. Определяем у кого дата увольнения больше(т.е. позже) и теперь запоминаем
                        //  именно этот документ

                        dtTempDoc=(DateTime)(hTempDoc.getItemValueDateTimeArray("signDateClose").elementAt(0));
                        dtDoc=(DateTime)(hDoc.getItemValueDateTimeArray("signDateClose").elementAt(0));
                        //  Если дата выхода в длит.отсут. из текущиго документа больше той, которая храниться во временном  документе...
                        if (dtDoc.timeDifference(dtTempDoc)>0) {
                            hTempDoc=hDoc;
                        }
                    }
                }
                else{
                    //  Явный неуволенный!
                    bOtherPerson=true;
                    break;
                }

                hDoc=hDocCollection.getNextDocument();
            }

            if (bOtherPerson==false){
                if (bCancelled){
                    // dtClose=(DateTime)(hTempDoc.getItemValueDateTimeArray("signDateClose").elementAt(0));
                    Session hSession;
                    hSession=NotesFactory.createSession();
                    dtClose=hSession.createDateTime(  ((DateTime)hTempDoc.getItemValueDateTimeArray("signDateClose").elementAt(0)).toJavaDate()  );
                    bFnRes=true;
                }
            }

            return bFnRes;

        }catch (Exception e){
            return false;
        }
    }


}