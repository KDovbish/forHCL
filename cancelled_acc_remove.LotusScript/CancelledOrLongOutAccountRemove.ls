%REM
	Agent CancelledOrLongOutAccountRemove
	Created Apr 6, 2018 by Konstantin G Dovbish/KIB
	Description: Comments for Agent

	Удаление учетных записей уволенных/длительно-отстуствующих.
	Принятие решения о том, какой статус у сотрудника, осуществляется путем анализа документов из Списка сотрудников.
	Производиться выборка всех возможных документов по именам прописанным в FullName учетной записи.

	Подразумевается, что агент должен запускаться из адресной книги.


%END REM
Option Declare

'  Местоположение адресной книги
'Dim NAMESSERVER As String
'Dim NAMESDB As String

'  Местоположение списка сотрудников
Dim STAFFSERVER As String
Dim STAFFDB As String
'  Время(в днях), которое должно пройти после увольнения/выхода в длит.отсут., прежде чем учетная запись будет удалена
Dim DAYSDELAY As Integer
'  Массив получаталей уведомления об удаленных учетных записях
Dim RECIPIENTLIST As Variant
'  Массив с именами учетных записей, которые удалять нельзя
Dim ACCOUNTEXCEPTIONS As Variant
'  Имя сервера, на котором будет расположена база, в которой будут сохранены удаляемые записи
Dim BACKUPSERVER As String
'  Имя базы, в которой будут сохранены удаляемые записи
Dim BACKUPDB As String
'  IDVault. Сервер
Dim IDVAULTSERVER As String
'  IDVault. База
Dim IDVAULTDB As String
'  Осуществлять ли удаление id-файлов из IDVault?
Dim FORCE_IDVAULT_DOC_REMOVE As Boolean




Sub Initialize

	Dim hSession As New NotesSession
	Dim hNames As NotesDatabase
	Dim hStaff As NotesDatabase
	Dim hPeople As NotesView
	Dim hPerson As NotesDocument
	Dim hStaffDocColl As NotesDocumentCollection
	Dim dt As Variant
	Dim hLog As Integer
	Dim hStaffDoc As NotesDocument

	Dim aAccountNameForRemove() As String
	Dim aAccountDocForRemove() As NotesDocument

	Dim aStaffDoc() As NotesDocument
	Dim nAccountForRemoveBound As Integer
	Dim hPDoc As NotesDocument


	'NAMESSERVER="Parsek-010/KIB"
	'NAMESDB="names.nsf"
	'STAFFSERVER="APP-001/KIB"
	'STAFFDB="staff.nsf"
	'DAYSDELAY=3

	'Dim tmpRecipients(0) As String
	'tmpRecipients(0)="CN=Konstantin G Dovbish/O=KIB"
	'RECIPIENTLIST=tmpRecipients


	Print "НАЧАЛО"

	'  Получаение настроек из профильного документа
	Set hPDoc=hSession.Currentdatabase.Getprofiledocument("CancelledOrLongOutAccountRemove")
	STAFFSERVER=hPDoc.GetItemValue("StaffServer")(0)
	STAFFDB=hPDoc.GetItemValue("StaffDb")(0)
	DAYSDELAY=hPDoc.Getitemvalue("DaysDelay")(0)
	RECIPIENTLIST=hPDoc.GetItemValue("RecipientList")
	ACCOUNTEXCEPTIONS=hPDoc.GetItemValue("AccountExceptions")
	BACKUPSERVER=hPDoc.GetItemValue("BackupServer")(0)
	BACKUPDB=hPDoc.GetItemValue("BackupDb")(0)
	IDVAULTSERVER=hPDoc.GetItemValue("IDVaultServer")(0)
	IDVAULTDB=hPDoc.GetItemValue("IDVaultDb")(0)
	If hPDoc.GetItemValue("ForceIDVaultDocRemove")(0)="1" Then FORCE_IDVAULT_DOC_REMOVE=True Else FORCE_IDVAULT_DOC_REMOVE=False


	' --- DEBUG ---
	'IDVAULTSERVER="Parsek-011/KIB"
	'IDVAULTDB="temp\Vault_KIB.nsf"
	'FORCE_IDVAULT_DOC_REMOVE=True


	' --- DEBUG ---
	'BACKUPSERVER="Parsek-011/KIB"
	'BACKUPDB="mail\removedpersons.nsf"

	'  --- DEBUG ---
	'Dim aAccountExceptions(1) As String
	'aAccountExceptions(0)={}
	'aAccountExceptions(1)={CN=Nataliya Azarova/OU=ecall/O=KIB}
	'ACCOUNTEXCEPTIONS=aAccountExceptions

	'  --- DEBUG ---
	Set hNames=New NotesDatabase("Parsek-010/KIB", "names.nsf")
	If hNames.IsOpen=False Then Exit Sub
	'  --- Продуктив ---
	'Set hNames=hSession.Currentdatabase

	'  Проверки допустимости работы дальнейшего кода
	Set hStaff=New NotesDatabase(STAFFSERVER, STAFFDB)
	If hStaff.IsOpen=False Then Exit Sub
	Set hPeople=hNames.GetView("$People")
	If hPeople Is Nothing Then Exit Sub

	nAccountForRemoveBound=-1
	Set hPerson=hPeople.GetFirstDocument()
	Do While Not(hPerson Is Nothing)

		'  --- DEBUG ---
		Print hPerson.GetItemValue("FullName")(0)

		'  Проверяем, входит ли рассматриваемая учетная запись в исключения, которые нельзя удалять
		If IsNull(ArrayGetIndex(ACCOUNTEXCEPTIONS, hPerson.GetItemValue("FullName")(0),5)) Then
			'  Получить из Списка сотрудников все карточки которые закреплены за заданным FullName
			Set hStaffDocColl=GetStaffCollection(hStaff, hPerson.GetItemValue("FullName"))
			If Not(hStaffDocColl Is Nothing) Then
				'  Проверяем, является ли сотрудник, представленный коллекцией документов из Списка сотрудников уволенным или длит.отсутствующим
				If IsCancelledOrLongOut_Collection(hStaffDocColl, hStaffDoc) Then

					'  Получаем дату увольнения/выхода в длит.отсут.
					If hStaffDoc.GetItemValue("signCancelled")(0)="1" Then dt=hStaffDoc.GetItemValue("signDateClose")(0) Else dt=hStaffDoc.GetItemValue("longoutDateFrom")(0)

					'  Прошла ли заданная задержка в днях после увольнения/выхода в длит.отсут.
					If DaysDatesDifference(Now(), dt)>=DAYSDELAY Then

						nAccountForRemoveBound=nAccountForRemoveBound+1

						'  Сохраняем ссылки на документы, которые будем удалять. Данный массив будет использоваться
						'  для предварительного переноса в отдельную базу удаляемых документов
						ReDim Preserve aAccountDocForRemove(nAccountForRemoveBound)
						Set aAccountDocForRemove(nAccountForRemoveBound)=hPerson
						'  Сохраняем полное имя учетной записи, которую нужно будет удалить через админпроцесс и
						'  которое нужно будет добавлять в информационное сообщение. Именно потому что информационное сообщение
						'  отправляется последним, уже после бекапирования и самое главное, запуска админпроцесса на удаление,
						'  нужно иметь отдельный массив с именами учетных записей, которые были удалены
						ReDim Preserve aAccountNameForRemove(nAccountForRemoveBound)
						aAccountNameForRemove(nAccountForRemoveBound)=hPerson.GetItemValue("FullName")(0)
						'  Сохраняем ссылку на документ в Списке сотрудников, который завязан на удаляемую учетную запись
						ReDim Preserve aStaffDoc(nAccountForRemoveBound)
						Set aStaffDoc(nAccountForRemoveBound)=hStaffDoc

						' --- DEBUG ---
						Exit Do

					End If
				End If
			End If
		End If


		Set hPerson=hPeople.Getnextdocument(hPerson)
	Loop

	If nAccountForRemoveBound>=0 Then
		'  Резервное копирование удаляемых документов
		If BACKUPSERVER<>"" And BACKUPDB<>"" Then
			'  Делаем резервные копии учетных записей из адресной книги в сторонней базе
			Call CopyDocumentArrayToOtherDb(aAccountDocForRemove, BACKUPSERVER, BACKUPDB)
			'  Если разрешено удаление документов в IDVault, то переносим документы, завязанные на массив найденных учетных записей, в стороннюю базу
			If FORCE_IDVAULT_DOC_REMOVE Then
				If IDVAULTSERVER<>"" And IDVAULTDB<>"" Then
					Call MoveIDFileToOtherDb(aAccountNameForRemove, IDVAULTSERVER, IDVAULTDB, BACKUPSERVER, BACKUPDB)
				End If
			End If
		End If
		'  Удаляем через admin-процесс найденные учетные записи
		'Call CreateDeleteUserAdminPRequests(aAccountNameForRemove)
		'  Рассылаем уведомление
		Call SendNotification(RECIPIENTLIST, aAccountNameForRemove, aStaffDoc)
	End If

	Print "КОНЕЦ"

End Sub


Sub Terminate

End Sub










%REM
	Function GetStaffCollection
	Description: Comments for Function

	Получить коллекцию документов из Списка сотрудников, соответствущую именам учетных записей из массива

	Входные параметры:
		hStaff		Список сотрудников
		aPerson		Массив имен
	Выход:
		NotesDocumentCollection/Nothing

%END REM
Function GetStaffCollection(hStaff As NotesDatabase, aPerson As Variant) As NotesDocumentCollection

	Dim hAllStaff As NotesView
	Dim hMainColl As NotesDocumentCollection
	Dim i As Integer
	Dim hDColl As NotesDocumentCollection
	Dim bIsMainColl As Boolean
	Dim hMainCollCopy As NotesDocumentCollection
	Dim hDoc As NotesDocument


	'  Предварительные проверки
	If hStaff.Isopen=False Then Exit Function
	If IsArray(aPerson)=False Then Exit Function
	Set hAllStaff=hStaff.GetView("AllStaff")
	If hAllStaff Is Nothing Then Exit Function

	'  Поиск имен из массива в Списке сотрудников и формирование на основе результат поиска
	'  сводной коллекции документов
	'
	'  Подобный алгоритм пришел в результате экспериментов. По видимому, нельзя использовать метод Merge()
	'  на пустой коллекции.
	bIsMainColl=False
	For i=0 To UBound(aPerson)

		Set hDColl=hAllStaff.GetAllDocumentsByKey(aPerson(i), True)
		If hDColl.Count>0 Then
			If bIsMainColl Then
				Call hMainColl.Merge(hDColl)
			Else
				Set hMainColl=hDColl
				bIsMainColl=True
			End If
		End If

	Next

	'  Если в коллекции что то есть, то проверяем каждый документ этой коллекции на точное совпадение по полю signPersonRLat.
	If Not(hMainColl Is Nothing) Then

		'  Создаем копию основной коллекции
		Set hMainCollCopy=hMainColl.Clone()

		'  Перебираем все документы копии
		Set hDoc=hMainCollCopy.GetFirstDocument()
		Do While Not(hDoc Is Nothing)

			If IsNull(ArrayGetIndex(aPerson, hDoc.GetItemValue("signPersonRLat")(0), 5)) Then

				'  Если ни одно из имен в массиве не совпадает со значением в поле signPersonRLat,
				'  то удаляаем этот документ в основной коллекции
				Call hMainColl.Subtract(hDoc)

			End If

			Set hDoc=hMainCollCopy.GetNextDocument(hDoc)
		Loop

		If hMainColl.Count=0 Then
			Set GetStaffCollection=Nothing
		Else
			Set GetStaffCollection=hMainColl
		End If

	End If

	'  В этой точке алгоритма либо hMainColl равно Nothing, т.к. индексный поикс был изначально неуспешен. Тогда и функция вернет Nothing.
	'  Либо прошла обработка непустой hMainColl на предмет точного совпадения с полем signPersonRLat. По концовке возврат
	'  либо Nothing либо коллекция с найденным точном совпадением по входному массиву.





	%REM
	'  Получаем коллекцию из Списка сотрудников по первому имени из массива имен
	Set hMainColl=hAllStaff.Getalldocumentsbykey(aPerson(0), True)
	If UBound(aPerson)>0 Then

		'  Получаем коллекции из Списка сотрудников по остальным именам из массива имен
		'  Доливаем полученные коллекции в основную коллекцию
		For i=1 To UBound(aPerson)

			Set hTempColl=hAllStaff.Getalldocumentsbykey(aPerson(i), True)
			If hTempColl.Count>0 Then Call hMainColl.Merge(hTempColl)



			If hTempColl.Count>0 Then

				Set hTempCollDoc=hTempColl.Getfirstdocument()
				Do
					Call hMainColl.Adddocument(hTempCollDoc)
					Set hTempCollDoc=hTempColl.Getnextdocument(hTempCollDoc)
				Loop Until hTempCollDoc Is Nothing

			End If


		Next

	End If

	Set GetStaffCollection=hMainColl
	%END REM


End Function

%REM
	Sub CreateDeleteUserAdminPRequests
	Description: Comments for Sub

	Создать админ-запросы на удаление учетных записей, имена которых содержаться в массиве

	Входные параметры:
		aAccountForRemove		массив с именами учетных записей

%END REM
Sub CreateDeleteUserAdminPRequests(aAccountForRemove() As String)

	Dim hSession As New NotesSession
	Dim hAdminP As NotesAdministrationProcess
	Dim i As Integer


	If IsArrayEmpty(aAccountForRemove) Then Exit Sub

	Set hAdminP=hSession.CreateAdministrationProcess(hSession.CurrentDatabase.Server)

	For i=0 To UBound(aAccountForRemove)
		Call hAdminp.Deleteuser(aAccountForRemove(i), False, MAILFILE_DELETE_NONE, "")
	Next

End Sub
%REM
	Sub SendNotification
	Description:

	Отсылка уведомления об удаленных учетных записях

	Входные параметры:
		aRecipients			список получателей информационного уведомления
		aAccountForRemove	массив dn-имен учетных записей
		aStaffDoc			массив ссылок на документы карточке в Списке сотрудников

	Примечание: оба массива из входных парамтров являются синхронными по отошению друг к другу

%END REM
Sub SendNotification(aRecipients As Variant, aAccountForRemove() As String, aStaffDoc() As NotesDocument)
	Dim hSession As New NotesSession
	Dim hMessage As NotesDocument
	Dim hSendToItem As NotesItem
	Dim hBody As NotesRichTextItem
	Dim hRTextStyle As NotesRichTextStyle
	Dim i As Integer


	If IsArrayEmpty(aRecipients) Or IsArrayEmpty(aAccountForRemove) Or IsArrayEmpty(aStaffDoc) Then Exit Sub

	Set hMessage=hSession.CurrentDatabase.CreateDocument()

	Call hMessage.ReplaceItemValue("Form", "Memo")
	Set hSendToItem=New NotesItem(hMessage, "SendTo", aRecipients, NAMES)
	Call hMessage.ReplaceItemValue("Principal", "CancelledOrLongOutAccountRemove Agent")
	Call hMessage.ReplaceItemValue("Subject", {Учетные записи уволенных или длит.отсут., которые были удалены})

	Set hBody=hMessage.CreateRichTextItem("Body")
	Set hRTextStyle=hSession.CreateRichTextStyle()

	hRTextStyle.Fontsize=10
	hRTextStyle.Bold=True
	Call hBody.Appendstyle(hRTextStyle)
	Call hBody.AppendText("Количество удаленных учетных записей: " & UBound(aAccountForRemove)+1)
	Call hBody.Addnewline(1)
	Call hBody.Appendtext("Количество дней, прошедших после увольнения/выхода в длит.остут.: " & DAYSDELAY)
	hRTextStyle.Bold=False
	Call hBody.Appendstyle(hRTextStyle)
	Call hBody.AddNewLine(2)

	For i=0 To UBound(aAccountForRemove)
		Call hBody.AppendDocLink(aStaffDoc(i), "")
		Call hBody.AppendText(Chr(9))
		Call hBody.AppendText(GetAbbreviatedName(aAccountForRemove(i)))
		Call hBody.Addnewline(1)
	Next

	Call hMessage.Send(False)

End Sub
%REM
	Function GetAbbriviatedName
	Description:

	Получить abbreviated имя

	Примеры:

		Имя: "KIB"								Возврат: "KIB"
		Имя: "/KIB"								Возврат: "/KIB"
		Имя: "O=KIB"							Возврат: "/KIB"
		Имя: "OU=POS/O=KIB"						Возврат: "/POS/KIB"
		Имя: "/OU=POS/O=KIB"					Возврат: "/POS/KIB"
		Имя: "Ivan I Testov"					Возврат: "Ivan I Testov"
		Имя: "CN=Ivan I Testov/OU=POS/O=KIB"	Возврат: "Ivan I Testov/POS/KIB"
		Имя: "Ivan I Testov/POS/KIB"			Возврат: "Ivan I Testov/POS/KIB"
		Имя: ""									Возврат: ""


%END REM
Public Function GetAbbreviatedName(cAccount As String) As String
	Dim hN As New NotesName(cAccount)
	GetAbbreviatedName=hN.Abbreviated
End Function
%REM
	Sub CopyDocumentToOtherDatabase
	Description: Comments for Sub

	Копировать все документы из массива в заданную базу

	Входные параметры
		aDoc		массив со ссылками на документы
		cServer		имя сервера для бекапа
		cDb			имя базы для бекапа

%END REM
Sub CopyDocumentArrayToOtherDb(aDoc() As NotesDocument, cServer As String, cDb As String)

	Dim hBackupDb As NotesDatabase
	Dim i As Integer


	Set hBackupDb=New NotesDatabase(cServer, cDb)
	If hBackupDb.IsOpen=False Then
		Call hBackupDb.Create(cServer, cDb, True)
		If hBackupDb.IsOpen=False Then Exit Sub
	End If

	For i=0 To UBound(aDoc)
		Call aDoc(i).CopyToDatabase(hBackupDb)
	Next

End Sub
%REM
	Sub MoveIDFileToOtherDb
	Description: Comments for Sub

	Произвести перенос id-файлов из IDVault в стороннюю базу
	Входные параметры:
		aAccount			массив DN-имен учетных записей
		cIDVaultServer		имя сервера, на котором храниться IDVault
		cIDVaultDb			имя базы IDVault
		cBackupServer		имя сервера, на котором храниться сторонняя база, в которую будут переноситься id-файлы
		cBackupDb			имя базы, в которую будут переноситься id-файлы

%END REM
Sub MoveIDFileToOtherDb(aAccount() As String, cIDVaultServer As String, cIDVaultDb As String, cBackupServer As String, cBackupDb As String)

	Dim hBackupDb As NotesDatabase
	Dim hIDVault As NotesDatabase
	Dim hIDFileView As NotesView
	Dim i As Integer
	Dim hDoc As NotesDocument


	'  Проверяем, открывается ли база IDVault
	Set hIDVault=New NotesDatabase(cIDVaultServer, cIDVaultDb)
	If hIDVault.IsOpen=False Then Exit Sub

	'  Проверяем, открывается ли база, в которую будут переноситься id-файлы. В случае необходимости - создаем новую.
	Set hBackupDb=New NotesDatabase(cBackupServer, cBackupDb)
	If hBackupDb.IsOpen=False Then
		Call hBackupDb.Create(cBackupServer, cBackupDb, True)
		If hBackupDb.IsOpen=False Then Exit Sub
	End If

	'  Перенос id-файлов для имен уч.записей из массива в стороннюю базу
	Set hIDFileView=hIDVault.GetView("$IDFile")
	If Not(hIDFileView Is Nothing) Then

		For i=0 To UBound(aAccount)
			Set hDoc=hIDFileView.GetDocumentByKey(aAccount(i))
			If Not(hDoc Is Nothing) Then
				Call hDoc.Copytodatabase(hBackupDb)
				Call hDoc.Remove(True)
			End If
		Next

	End If

End Sub

%REM
	Function DateDifferenceInDays
	Description:

	Разница между двумя датами в днях

%END REM
Function DaysDatesDifference(dtFirst As Variant, dtSecond As Variant) As Long
Dim nSecondsDiff As Long

	nSecondsDiff=DatesDifference(dtFirst, dtSecond)
	DaysDatesDifference=Fix(Abs(nSecondsDiff)/86400)

End Function

%REM
	Function DatesDifference
	Description:

	Количество секунд между двумя DateTime
	"Правильное" вычитание: dtFirst - dtSecond ,т.е. dtFirst ближе к настоящему, а dtSecond - это прошлое

%END REM

Function DatesDifference(dtFirst As Variant, dtSecond As Variant) As Long
Dim notesdateFirst As NotesDateTime, notesdateSecond As NotesDateTime

	Set notesdateFirst=New NotesDateTime(CStr(dtFirst))
	Set notesdateSecond=New NotesDateTime(CStr(dtSecond))

	DatesDifference=notesdateFirst.Timedifference(notesdateSecond)

End Function




%REM
	Function IsCancelledOrLongOut_Collection
	Description: Comments for Function

	Является ли сотрудник уволенным или длительно-отсутствующим? Анализ коллекции документов из Списка сотрудников.

	Входные параметры:
		in	hDColl		коллекция документов из Списка сотрудников
		out	hStaffDoc	в случае, если функция возвращает True, документ с самой поздней датой увольнения/выхода в длит.отсут.
	Выход:
		True/False		является/не является

%END REM
Function IsCancelledOrLongOut_Collection(hDColl As NotesDocumentCollection, hStaffDoc As NotesDocument) As Boolean

	Dim hDoc As NotesDocument
	Dim bDocCancelled As Boolean, bDocLongOut As Boolean
	Dim bCancelledOrLongOutDoc As Boolean
	Dim hTempDoc As NotesDocument
	Dim dtDoc As Variant, dtTempDoc As Variant


	'  Признак того, что найден по крайней мере один документ, который является уволенным или длительно-отсутствующим
	bCancelledOrLongOutDoc=False

	Set hDoc=hDColl.Getfirstdocument()
	Do While Not(hDoc Is Nothing)

		If hDoc.GetItemvalue("signCancelled")(0)="1" Then bDocCancelled=True Else bDocCancelled=False
		If hDoc.GetItemvalue("longOut")(0)="1" Then bDocLongOut=True Else bDocLongOut=False

		If bDocCancelled=False And bDocLongOut=False  Then
			'  Явный не уволенный!
			'  Выход из функции. Ответ отрицательный.
			IsCancelledOrLongOut_Collection=False
			Exit Function
		Else

			'  Сотрудник либо уволен либо длительно-отсутствует

			If bCancelledOrLongOutDoc=False Then
				'  Подобная запись обнаружена в первый раз
				'  Сохраняем ссылку на найденный документ во временной переменной
				Set hTempDoc=hDoc
				bCancelledOrLongOutDoc=True
			Else
				'  Подобная запись обнаружена в очередной раз
				'  Производим сравнение дат увольнения/выхода в длит.отсут. для текущего документа и ранее найденного документа. Сохраняем ссылку на
				'  документ с более позней датой.
				If bDocCancelled Then dtDoc=hDoc.GetItemvalue("signDateClose")(0) Else dtDoc=hDoc.GetItemValue("lngoutDateFrom")(0)
				If hTempDoc.GetItemValue("signCancelled")(0)="1" Then dtTempDoc=hTempDoc.GetItemValue("signDateClose")(0) Else dtTempDoc=hTempDoc.GetItemValue("longoutDateFrom")(0)
				If dtDoc>dtTempDoc Then	Set hTempDoc=hDoc
			End If

		End If

		Set hDoc=hDColl.Getnextdocument(hDoc)
	Loop

	'  В этой точке алгоритма может быть только два варианта:
	'  		1. В коллекции не было ни одного документа. Соответственно hTempDoc осталось Nothing.
	'  		2. Сотрудник действительно является уволенным или длит.отсут. и в переменной hTempDoc храниться ссылка на самый "последний по дате" документ коллекции

	If Not(hTempDoc Is Nothing) Then
		Set hStaffDoc=hTempDoc
		IsCancelledOrLongOut_Collection=True
	Else
		IsCancelledOrLongOut_Collection=False
	End If

End Function
%REM

	Пуст ли массив?

	Если заходит пустая Variant функция вернит True
	Если заходит статический массив функция вернет False
	Если заходит неинициализированный динамический массив функция вернет True
	Если заходит ициализированный динамический массив функция вернет False

%END REM

Function IsArrayEmpty(a As Variant)
	Err=0
	On Error Resume Next
	Dim nBound As Integer
	nBound=UBound(a)
	On Error GoTo 0
	If Err=0 Then IsArrayEmpty=False Else IsArrayEmpty=True
End Function
