%REM
	Agent AllGroupAutoFilling
	Created Jul 9, 2018 by Konstantin G Dovbish/KIB
	Description: Comments for Agent
%END REM
Option Declare

'Use "CGroupExposer"
'Use "CAggrGroupHandler"
'Use "CAggrGroupNormalizer"

'  ********** ПАРАМЕТРЫ СКРИПТА(ИЗ ПРОФИЛЬНОГО ДОКУМЕНТА) **********

'  Группа для наполнения
Dim MAINGROUP As String
'  Шаблон(начальная неизменная часть) имени дочерней группы
Dim CHILDGROUPTEMPLATE As String
'  Иерархическая составляющая дочерних групп
Dim CHILDGROUPHIERARHY As String
'  Предельный лимит группы
Dim GROUPLIMIT As Long
'  Список сотрудников
Dim STAFFSERVER As String
Dim STAFFDB As String
'  ИСКЛЮЧЕНИЯ
'  Перечень учетных записей (возможно использование групп) которые НИКОГДА не должны присутствовать в группе All/KIB
Dim EXCEPTIONS_NEVER As Variant
'  ИСКЛЮЧЕНИЯ
'  Перечень учетных записей(возможно использование групп) которые ВСЕГДА должны присутствовать в группе All/KIB
Dim EXCEPTIONS_ALWAYS As Variant


'  ********** ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ СКРИПТА  **********
'  Ссылка на адресную книгу
Dim hNames As NotesDatabase
'  Ссылка на Список сотрудников
Dim hStaff As NotesDatabase
'  Раскрытый(в списке могли быть группы) список исключений "НИКОГДА"
Dim aExceptionsNever As Variant
'  Раскрытый(в списке могли быть группы) список исключений "ВСЕГДА"
Dim aExceptionsAlways As Variant
'  Объект для работы с аггрегированной группой
Dim hAGH As CAggrGroupHandler










%REM

	Обработчик аггрегированной группы

	Входные параметры конструктра:

		NotesDocument		ссылка на основную группу
		String				шаблон(начальная неизменная часть common-имени) дочерних групп
		String				иерархическая составляющая дочерних групп
		Long				предельно допустимый размер группы


	Публичные методы:

		AddArr( <массив> ) As Boolean
		Добавить массив в аггрегированную группу.
		Указанный массив разливается по дочерним группам. Если не хватает дочерних групп,
		то создаются новые дочерние группы

		RemoveArr( <массив> ) As Boolean
		Удалить массив из аггрегированной группы

		IsMember( <строка с именем учетной записи в DN-формате> ) As Boolean
		Является ли учетная запись членом аггрегированной группы?

		Property Get GroupCycling As Boolean
		Получить признак того, что аггрегированная группа может быть зациклена

		Property Get Members As Variant
		Получить массив раскрытой аггрегированной группы

%END REM
Public Class CAggrGroupHandler

	'  ПАРАМЕТРЫ, КОТОРЫЕ ЗАДАЮТСЯ ЧЕРЕЗ КОНСТРУКТОР
	'  Ссылка на основную группу
	Private hMainGroup As NotesDocument
	'  Шаблон(начальная неизменная часть common-имени) дочерних групп
	Private cChildGroupCNTemplate As String
	'  Иерархическая составляющая дочерних групп
	Private cHierarchy As String
	'  Предельно допустимый размер группы
	Private nLimit As Long
	'  Максимально допустимая глубина вложенности дочерних групп(Константа! Задается непосредственно в конструкторе)
	Private nChildGroupsDeepMax As Integer


	'  Синхронные массивы: члены раскрытой группы и ссылки на группы в которых храняться эти члены
	'  В случае операций по добавлению/удалению членов, содержимое данных массивов модифицируются
	Private aMember As Variant
	Private aMemberGroup As Variant

	'  Массив ссылок на дочерние группы
	'  В случае добавления новых/удаления существующих дочерних группы, содержимое данного массива модифицируется
	Private aChildGroup As Variant

	'  Счетчик вложенности дочерних групп
	'  Используется в процедуре раскрытия группы
	Private nChildGroupsCount As Integer

	'  Признак зацикленности группы
	'  Выставляется каждый раз после выполнения процедуры раскрытия группы
	Private bGroupCycling As Boolean

	'  Признак допустимости использования ключевых публичных методов, функционирование
	'  которых завязано на то, корректно или нет была раскрыта группа, т.е. успешно
	'  ли были проведены все подготовительные операции для полноценного функционирования
	'  объекта класса
	Private bRunEnable As Boolean


	'  Дать признак зацикленности рассматриваемой группы
	Property Get GroupCycling As Boolean
		GroupCycling=Me.bGroupCycling
	End Property

	'  Получить раскрытую группу
	Property Get Members As Variant
		Members=aMember
	End Property



	%REM
	' --- DEBUG ---
	Property Get dbg_aMember As Variant
		dbg_aMember=aMember
	End Property

	' --- DEBUG ---
	Property Get dbg_aMemberGroup As Variant
		dbg_aMemberGroup=aMemberGroup
	End Property

	' --- DEBUG ---
	Property Get dbg_aChildGroup As Variant
		dbg_aChildGroup=aChildGroup
	End Property
	%END REM




	%REM
		Конструктор
		Параметры
			hMainGroup				ссылка на основную группу
			cChildGroupCNTemplate	шаблон(начальная неизменная часть common-имени) дочерних групп
			cHierarchy				иерархическая составляющая дочерних групп
			nLimit					предельно допустимый размер группы
	%END REM
	Sub New(hMainGroup As NotesDocument, cChildGroupCNTemplate As String, cHierarchy As String, nLimit As Long)

		'  То, что передается через параметры...
		Set Me.hMainGroup=hMainGroup
		Me.cChildGroupCNTemplate=cChildGroupCNTemplate
		Me.cHierarchy=cHierarchy
		Me.nLimit=nLimit

		'  Я не счел нужным передавать максимально допустимую глубину вложенности групп через конструктор.
		'  При необходимости можно это сделать. Фактически мы имеем константу.
		Me.nChildGroupsDeepMax=10

		'  По умолчанию ключевые публичные методы класса запускать нельзя.
		'  После раскрытия группы, если все пройдет нормально, это признак изменится
		Me.bRunEnable=False

		Call refresh()

	End Sub


	%REM
		Удалить из группы все учетные записи, содержащиеся в массиве
		Тестирование проведено 24-07-2018
	%END REM
	Function RemoveArr(aAccount As Variant) As Boolean

		Dim i As Integer
		Dim nIndex As Variant
		Dim bRemoveError As Boolean


		If Me.bRunEnable=False Then
			RemoveArr=False
			Exit Function
		End If

		If IsArrayEmpty(aAccount) Then
			RemoveArr=True
			Exit Function
		End If

		If IsArrayEmpty(Me.aMember) Then
			RemoveArr=True
			Exit Function
		End If


		bRemoveError=False
		'  Перебираем учетные записи для удаления...
		For i=0 To UBound(aAccount)
			'  Если учетная запись присутствует во внутреннем массиве(члены раскрытой группы)...
			nIndex=ArrayGetIndex(Me.aMember, aAccount(i), 5)
			If IsNull(nIndex)=False Then
				'  ... то данная учетная запись удаляется из группы
				If RemoveMemberFromSpecificGroup(Me.aMemberGroup(nIndex), CStr(Me.aMember(nIndex))) Then
					'  Если удаление в адресной книге прошло успешно, то вычищаем соответствующие элементы внутренних массивов
					Me.aMember(nIndex)=""
					Set Me.aMemberGroup(nIndex)=Nothing
				Else
					bRemoveError=True
					Exit For
				End If
			End If
		Next

		RemoveArr=Not bRemoveError

	End Function


	%REM
	Удалить заданный элемент из заданной группы

	Вход:
		in	hGroup		ссылка на группу
		in	сAccount	учетная запись для удаления
	Выход:
		True/False
	Тестирование проведено 23-07-2018
	%END REM
	Private Function RemoveMemberFromSpecificGroup(hGroup As NotesDocument, cAccount As String) As Boolean

		Dim aMember As Variant
		Dim nElementIndex As Variant
		Dim aMemberNew() As String
		Dim nMemberNewArrIndex As Integer
		Dim i As Integer


		RemoveMemberFromSpecificGroup=False

		'  Получаем текущее содержимое группы
		aMember=hGroup.GetItemValue("Members")
		If IsArrayEmpty(aMember) Then
			RemoveMemberFromSpecificGroup=True
			Exit Function
		End If

		'  Получаем индекс того члена группы, который нужно удалить
		nElementIndex=ArrayGetIndex(aMember, cAccount, 5)
		If IsNull(nElementIndex) Then
			RemoveMemberFromSpecificGroup=True
			Exit Function
		End If

		'  Случай последнего удаляемого члена
		If UBound(aMember)=0 Then
			If ReplaceGroup(hGroup, "")=0 Then
				RemoveMemberFromSpecificGroup=True
				Exit Function
			End If
		End If

		'  Формируем новый массив для группы без удаляемого члена...
		ReDim aMemberNew(UBound(aMember)-1)
		nMemberNewArrIndex=-1
		For i=0 To nElementIndex-1
			nMemberNewArrIndex=nMemberNewArrIndex+1
			aMemberNew(nMemberNewArrIndex)=aMember(i)
		Next
		For i=nElementIndex+1 To UBound(aMember)
			nMemberNewArrIndex=nMemberNewArrIndex+1
			aMemberNew(nMemberNewArrIndex)=aMember(i)
		Next
		'  ... и сохраняем его в группе
		If ReplaceGroup(hGroup, aMemberNew)=0 Then RemoveMemberFromSpecificGroup=True

	End Function






	%REM
		Явяляется ли учетная запись с заданным именем членом раскрытой аггрегированной группы
	%END REM
	Function IsMember(cAccount As String) As Boolean
		IsMember=False
		If bRunEnable Then
			If IsArrayEmpty(Me.aMember)=False Then
				IsMember=Not IsNull(ArrayGetIndex(Me.aMember, cAccount, 5))
			End If
		End If
	End Function



	%REM
		Добавить массив учетных записей в аггрегированную группу
		Вход:
			in 	aAccount		массив для добавления в аггрегированную группу
		Выход:
			True/False
		Тестирование проведено 19-07-2018
	%END REM
	Function AddArr(aAccount As Variant) As Boolean

		AddArr=False

		If Me.bRunEnable=False Then Exit Function

		If IsArrayEmpty(aAccount)=False Then
			If AddArrToExistChildGroups(aAccount) Then
				If IsArrayEmpty(aAccount)=False Then
					If AddArrToNewChildGroups(aAccount) Then
						AddArr=True
					End If
				Else
					AddArr=True
				End If
			End If
		Else
			AddArr=True
		End If

	End Function



	%REM
		Попытаться добавить массив(или его часть) в конкретную группу
		с учетом лимита группы

		Входные параметры:
			in			hGroup		ссылка на группу
			in/out		aAccount	массив для добавления
		Выход:
			True/False	сигнализирует о том, можно ли проводить анализ параметра aAccount

		Примечание:
		Результат False говорит о проблемах записи в физический файл адресной книги.
		Фактически, настоящим результатом работы функции будет содержимое массива hAccount. Он может вернуться неизменным,
		может вернуться его часть(та, которая не была записана в группу), может вернуться EMPTY. Однако все эти значения
		имеет смысл анализировать только в случае если функция в целом вернет True.

		Тестирование завершено 17.07.2018
	%END REM
	Private Function AddArrToOneGroup(hGroup As NotesDocument, aAccount As Variant) As Boolean

		Dim Acurrent As Variant, Atemp As Variant
		Dim B As Variant
		Dim i As Integer
		Dim AIndex As Integer
		Dim nReplaceResult As Integer
		Dim aEmpty As Variant
		Dim nBound As Integer
		Dim aTmp() As String


		If IsArrayEmpty(aAccount)=True Then
			AddArrToOneGroup=True
			Exit Function
		End If

		'  Сохраняем текущее состояние группы
		Acurrent=hGroup.GetItemValue("Members")
		Call DeleteDublicateAndEmpty(Acurrent)

		'  Вычисляем массив элементов, которые можно будет добавить в группу
		Atemp=Acurrent
		For i=0 To UBound(aAccount)
			Call AddElementToStringArray(Atemp, CStr(aAccount(i)))
			If ArrayMatchingToGroupLimit(Atemp, Me.nLimit) Then
				'  Накапливаем массив тех элементов, которые возможно добавить в группу
				Call AddElementToStringArray(B, CStr(aAccount(i)))
				'  Фиксируем индекс последнего элемента в исходном массиве, который еще будет добавлен в группу.
				'  Все что правее индекса добавить уже не представляется возможным из-за превышения лимита группы.
				AIndex=i
			Else
				Exit For
			End If
		Next

		If IsArrayEmpty(B)=False Then

			'  Если хотя бы один раз было произведено присвоение массиву "В", т.е. если был найден
			'  один или несколько элементов, которые возможно сохранить в группе...

			nReplaceResult=ReplaceGroup(hGroup, ArrayAddition(Acurrent, B))
			If nReplaceResult=0 Then

				'  Сохранение в группу прошло успешно!
				AddArrToOneGroup=True

				'  Фиксируем(добавляем) во внутренние массивы информацию о добавленных в группу членах
				For i=0 To UBound(B)
					Call AddElementToStringArray(aMember, CStr(B(i)))
					Call AddElementToNotesDocumentArray(aMemberGroup, hGroup)
				Next

				'  Оставляем в исходном массиве только те значения, которые не удалось сохранить в группе
				If AIndex=UBound(aAccount) Then
					aAccount=aEmpty
				Else
					nBound=-1
					For i=AIndex+1 To UBound(aAccount)
						nBound=nBound+1
						ReDim Preserve aTmp(nBound)
						aTmp(nBound)=aAccount(i)
					Next
					aAccount=aTmp
				End If

			Else
				AddArrToOneGroup=False
			End If

		Else
			'  Если ни одно значение из исходного массива не смогло поместиться в группе...
			'  Функция возращает True и неизменный исходной массив. Функция нормально отработала и
			'  честно пыталась сохранить предложенный массив в составе группы.
			AddArrToOneGroup=True
		End If

	End Function



	%REM
		Попытаться разлить массив по существующим дочерним группам

		Вход:
			in/out	aAccount	массив для разлития по дочерним группам
		Выход:
			True/False	сигнализирует о том, можно ли проводить анализ параметра aAccount

		Примечание:
		Фактически, настоящим результатом функции является массив aAccount. Он может вернуться неизменными, может вернуться
		его часть, может быть возвращено EMPTY. Но анализ массива aAccount имеет смысл только если функция в целом вернет True

		Тестирование завершено 17.07.2018
	%END REM
	Private Function AddArrToExistChildGroups(aAccount As Variant) As Boolean

		Dim i As Integer


		AddArrToExistChildGroups=True

		If IsArrayEmpty(Me.aChildGroup)=False Then

			For i=0 To UBound(Me.aChildGroup)
				If AddArrToOneGroup( Me.aChildGroup(i), aAccount )=False Then
					AddArrToExistChildGroups=False
					Exit Function
				End If

				If IsArrayEmpty(aAccount) Then Exit For
			Next

		End If

	End Function





	%REM
		Создавать новые дочерние группы и добавлять в них массив до тех
		пор пока весь массив не будет сохранен.
		Параметры:
			in	aAccount	массив для разлития по новым дочерним группам
		Выход:
			True/False
		Тестирование проведено 18-07-2018
	%END REM
	Private Function AddArrToNewChildGroups(aAccount As Variant) As Boolean

		Dim hChildGroup As NotesDocument

		Do While True
			Set hChildGroup=CreateNewChildGroup()
			If Not(hChildGroup Is Nothing) Then
				If AddArrToOneGroup(hChildGroup, aAccount) Then
					If IsArrayEmpty(aAccount) Then Exit Do
				Else
					Exit Do
				End If
			Else
				Exit Do
			End If
		Loop

		AddArrToNewChildGroups=IsArrayEmpty(aAccount)

	End Function


	%REM
		Создать новую дочернюю группу
		В случае успешного создания, сохранить ссылку на эту группу во внутреннем массиве
		Выход:
			ссылка на созданную дочернюю группу или Nothing
		Тестирование проведено 18-07-2018
	%END REM
	Private Function CreateNewChildGroup() As NotesDocument

		Dim cChildGroup As String
		Dim hChildGroupTemplate As NotesDocument
		Dim hChildGroup As NotesDocument
		Dim aMainGroupMembers As Variant


		'  Получаем имя которое можно использовать для создания новой дочерней группы
		'  и, если уже существуют другие дочерние группы, то и ссылку на предыдущую по счету существующую
		'  дочернюю группу, которую можно использовать как шаблон для создания новой группы
		cChildGroup=GetNewChildGroupName(hChildGroupTemplate)

		'  Создаем группу в адресной книге(либо с нуля либо по шаблону)
		If hChildGroupTemplate Is Nothing Then
			Set hChildGroup=CreateNewGroup(cChildGroup)
		Else
			Set hChildGroup=CreateNewGroupPerTemplate(cChildGroup, hChildGroupTemplate)
		End If

		'  Прописываем созданную дочернюю группу в основной группе и во внутреннем массиве
		If Not (hChildGroup Is Nothing) Then

			aMainGroupMembers=Me.hMainGroup.GetItemValue("Members")
			Call AddElementToStringArray(aMainGroupMembers, cChildGroup)
			Call DeleteDublicateAndEmpty(aMainGroupMembers)

			If ReplaceGroup(Me.hMainGroup, aMainGroupMembers)=0 Then
				'  Фиксируем во внутреннем массиве, что создана новая дочерняя группа
				Call AddElementToNotesDocumentArray(Me.aChildGroup, hChildGroup)
				'  Возвращаем ссылку на созданную дочернюю группу
				Set CreateNewChildGroup=hChildGroup
			End If

		End If

	End Function



	%REM
	Создать новую группу в адресной книге
	Параметры:
		cGroupName	имя новой группы
	Возврат:
		NotesDocument-ссылка на новую группу/Nothing
	Тестирование проведено 18-07-2018
	%END REM
	Private Function CreateNewGroup(cGroupName As String) As NotesDocument

		Dim hGroup As NotesDocument
		Dim hItem As NotesItem

		Set hGroup=Me.hMainGroup.Parentdatabase.CreateDocument()
		Call hGroup.AppendItemValue("Form", "Group")
		Call hGroup.AppendItemValue("Type", "Group")
		Set hItem=New NotesItem(hGroup, "ListName", GetCanonicalizeName(cGroupName), NAMES)
		Set hItem=New NotesItem(hGroup, "Members", "", NAMES)
		Set hItem=New NotesItem(hGroup, "DocumentAccess", "[GroupModifier]", AUTHORS)

		If SaveGroup(hGroup)=0 Then Set CreateNewGroup=hGroup

	End Function


	%REM
	Создать новую группу взяв за основу другую группу
	Входные параметры:
		in	cNewGroupName	имя новой группы
		in	hTemplateGroup	существующая группа, которая будет взята за основу при создании новой группы
	Выход:
		ссылка на группу или Nothing, если группа не была создана

	Тестирование проведено 18-07-2018
	%END REM
	Private Function CreateNewGroupPerTemplate(cNewGroupName As String, hGroupTemplate As NotesDocument) As NotesDocument

		Dim hNewGroup As NotesDocument
		Dim cItemName As String
		Dim hSession As New NotesSession


		Set hNewGroup=Me.hMainGroup.Parentdatabase.Createdocument()
		Call hGroupTemplate.CopyAllItems(hNewGroup)
		Call hNewGroup.ReplaceItemValue("ListName", GetCanonicalizeName(cNewGroupName))
		Call hNewGroup.ReplaceItemValue("ListOwner", hSession.UserName)
		Call hNewGroup.ReplaceItemValue("LocalAdmin", hSession.UserName)

		Call hNewGroup.ReplaceItemValue("Members", "")
		Call hNewGroup.ReplaceItemValue("InternetAddress", "")

		Call hNewGroup.RemoveItem("$Revisions")
		Call hNewGroup.RemoveItem("$UpdatedBy")

		If SaveGroup(hNewGroup)=0 Then Set CreateNewGroupPerTemplate=hNewGroup

	End Function





	%REM
		Вычислить имя для новой дочерней группы
		Параметры:
			out		hChildGroupTemplate		ссылка на предыдущую по счету существующую дочернюю группу
											(для использования ее в качестве шаблона при дальнейшем
											создании новой группы)
		Выход:
			имя новой дочерней группы

		Тестирование проведено 18-07-2018
	%END REM
	Private Function GetNewChildGroupName(hChildGroupTemplate) As String

		Dim cAbbrHierarchy As String
		Dim nIndex As Integer
		Dim cChildGroupCN As String
		Dim cChildGroup As String
		Dim hChildGroup As NotesDocument


		'  Дополнительный контроль: иерархическая составляющая должна быть в сокращенном виде и без начального слеша
		cAbbrHierarchy=RemoveFirstSlash(GetAbbreviatedName(Me.cHierarchy))

		'  Находим индекс первой несуществующей дочерней группы, а заодно ссылку на последнюю встретившуюся дочернюю группу
		nIndex=0
		Do
			nIndex=nIndex+1
			cChildGroupCN=Me.cChildGroupCNTemplate & nIndex
			If cAbbrHierarchy<>"" Then cChildGroup=cChildGroupCN & "/" & cAbbrHierarchy Else cChildGroup=cChildGroupCN
			Set hChildGroup=IsChildGroupName(cChildGroup)
			If Not(hChildGroup Is Nothing) Then Set hChildGroupTemplate=hChildGroup
		Loop Until hChildGroup Is Nothing

		GetNewChildGroupName=GetCanonicalizeName(cChildGroup)

	End Function


	%REM
		Проверить внутренний массив с дочерними группами на предмет наличия
		дочерней группы с заданным именем
		Параметры:
			in		cGroupName		имя искомой группы
		Выход:
			ссылка на найденную группу или Nothing

		Тестирование проведено 18-07-2018
	%END REM
	Private Function IsChildGroupName(cGroupName As String) As NotesDocument

		Dim i As Integer
		Dim cAbbrGroupName As String


		If IsArrayEmpty(Me.aChildGroup)=False Then

			cAbbrGroupName=LCase(GetAbbreviatedName(cGroupName))

			For i=0 To UBound(Me.aChildGroup)
				If LCase(GetAbbreviatedName(Me.aChildGroup(i).GetItemValue("ListName")(0)))=cAbbrGroupName Then
					Set IsChildGroupName=Me.aChildGroup(i)
					Exit Function
				End If
			Next
		End If

	End Function



	%REM
		Заполнение массивов...
			aMember(члены раскрытой группы)
			aMemberGroup(группы к которые входят члены раскрытой группы)
			aChildGroup(ссылки на дочерние группы)
		... из адресной книги
		Т.е. фактически в вышеуказанных массивах отражается реальная ситуация из адресной книги.
		Внимание! Раскрытие группы и получение списка дочерних групп являются самыми трудоемкими операциями!
	%END REM
	Private Sub refresh

		'  Раскрываем основную группу
		Call Expose()
		'  Было ли обнаружено зацикливание в процесса раскрытия группы?
		If Me.bGroupCycling=False Then
			'  Получаем список дочерних групп
			Call GetChildGroups()
			'  Начальная инициализация объекта проведена нормально.
			'  Теперь возможен запуск ключевых публичных методов.
			Me.bRunEnable=True
		End If

	End Sub




	'  Получить ссылки на все дочерние группы и сохранить их во внутреннем массиве aChildGroup
	'  Тестирование проведено
	Private Sub GetChildGroups

		Dim aEmpty As Variant
		Dim cFormula As String
		Dim hDColl As NotesDocumentCollection
		Dim nBound As Integer
		Dim cAbbrHierarchy As String
		Dim hGroup As NotesDocument
		Dim cChildGroupCN As String
		Dim nNumberTemp As Integer
		Dim aChildGroupTemp() As NotesDocument


		'  Очищаем внутренний массив, в котором будут сохраняться дочерние группы
		Me.aChildGroup=aEmpty

		'  Находим все группы, которые подпадают под шаблон common-имени.
		'  Иерархия в расчет не берется. Ее анализ будет производиться в цикле для каждой найденной группы.
		cFormula={SELECT Form="Group" & @Like( @LowerCase(@Name([CN]; ListName)); "} & LCase(Me.cChildGroupCNTemplate) & {%")}

		Set hDColl=Me.hMainGroup.Parentdatabase.Search(cFormula, Nothing, 0)
		If hDColl.Count>0 Then

			'  Имя каждой найденной группы анализируем на предмет того, что стоит после шаблона.
			'  Если это число - считаем найденную группу дочерней.

			'  Индекс массива с найденными дочерними группами
			nBound=-1

			'  сокращенное имя иерархии без начального слеша для последующего сравнения
			cAbbrHierarchy=LCase(RemoveFirstSlash(GetAbbreviatedName(Me.cHierarchy)))

			Set hGroup=hDColl.GetFirstDocument()
			Do While Not (hGroup Is Nothing)
				cChildGroupCN=GetCommonName(hGroup.GetItemvalue("ListName")(0))
				If Len(cChildGroupCN)>Len(Me.cChildGroupCNTemplate) Then
					'  Если иерархия найденной группы, соответствует запрошенной...
					If LCase( RemoveFirstSlash(GetAbbreviatedName(GetHierarchyName(hGroup.GetItemvalue("ListName")(0)))) )=cAbbrHierarchy Then
						'  Пытаемся преобразовать в число, то что находиться правее шаблона.
						'  Преобразование прошло без ошибок? Если да, то считаем что нашли очередную дочернюю группу.
						Err=0
						On Error Resume Next
						nNumberTemp=CInt( Right(cChildGroupCN, Len(cChildGroupCN)-Len(Me.cChildGroupCNTemplate)) )
						On Error GoTo 0
						If Err=0 Then
							'  Запоминаем ссылку на группу
							nBound=nBound+1
							ReDim Preserve aChildGroupTemp(nBound)
							Set aChildGroupTemp(nBound)=hGroup
						End If
					End If
				End If
				Set hGroup=hDColl.GetNextDocument(hGroup)
			Loop

			If nBound>=0 Then
				Me.aChildGroup=aChildGroupTemp
			End If
		End If

	End Sub




	%REM
		Раскрыть дочернюю группу
		Все члены группы и все члены подгрупп, которые в нее входят, будут зафиксированы в массивах aMember и aMemberGroup
		Данная функция является вспомогательной(рекурсивной) для функции Expose()

		Тестирование проведено
	%END REM
	Private Sub ChildGroupExpose(hGroup As NotesDocument)

		Dim nChildGroupsCountOnThisLevel As Integer
		Dim aGroupMembers As Variant
		Dim cMember As String
		Dim hChildGroup As NotesDocument


		'  Вошли в дочернюю группу. Увеличиваем счетчик...
		Me.nChildGroupsCount=Me.nChildGroupsCount+1
		'  Запоминаем глубину вложенности групп на текущем уровне...
		nChildGroupsCountOnThisLevel=Me.nChildGroupsCount

		'  Обнаружено, что группа в целом имеет признаки зацикливания! Прерываем обработку...
		If Me.nChildGroupsCount>Me.nChildGroupsDeepMax Then	Exit Sub

		aGroupMembers=hGroup.GetItemValue("Members")
		'  Перебираем всех членов группы. Работаем только с непустыми членами.
		ForAll vMember In hGroup.GetItemValue("Members")
			cMember=CStr(vMember)
			If Trim(cMember)<>"" Then
				Set hChildGroup=IsGroup(hGroup.ParentDatabase, cMember)
				If hChildGroup Is Nothing Then
					'  Найдена не группа
					'  Добавляем найденную учетную запись во внутренний массив
					Call AddElementToStringArray(aMember, cMember)
					'  Добавляем группу в которую входит учетная запись во внутренний массив
					Call AddElementToNotesDocumentArray(aMemberGroup, hGroup)
				Else
					'  Найдена дочерняя группа. Раскрываем...
					Call ChildGroupExpose(hChildGroup)
					If Me.nChildGroupsCount>Me.nChildGroupsDeepMax Then
						Exit Sub
					Else
						'  Восстанавливаем уровнь вложенности групп для текущего уровня для
						'  последующих вызовов ChildGroupExpose() на этом уровне
						Me.nChildGroupsCount=nChildGroupsCountOnThisLevel
					End If
				End If
			End If
		End ForAll

	End Sub


	%REM
		Раскрыть основную группу

		Ссылка на основную группу передается через конструктор и сохраняется
		в переменной Me.hMainGroup. Раскрытая группа сохраняется в двух параллельных(синхронных)
		массивах:
			Me.aMember				String()
			Me.aMemberGroup			NotesDocument()

		Тестирование проведено
	%END REM
	Private Sub Expose

		Dim aGroupMembers As Variant
		Dim cMember As String
		Dim hChildGroup As NotesDocument


		'  По умолчанию зацикленности нет.
		Me.bGroupCycling=False

		aGroupMembers=Me.hMainGroup.GetItemValue("Members")
		'  Перебираем все членов группы. Работаем только с непустыми членами.
		ForAll vMember In aGroupMembers
			cMember=CStr(vMember)
			If Trim(cMember)<>"" Then
				'  Является ли рассматриваемый член группы сам группой или нет?
				Set hChildGroup=IsGroup(Me.hMainGroup.ParentDatabase, cMember)
				If hChildGroup Is Nothing Then
					'  Найдена не группа
					'  Добавляем найденную учетную запись во внутренний массив
					Call AddElementToStringArray(aMember, cMember)
					'  Добавляем группу в которую входит учетная запись во внутренний массив
					Call AddElementToNotesDocumentArray(aMemberGroup, Me.hMainGroup)
				Else
					'  Найдена дочерняя группа. Раскрываем ее...
					Me.nChildGroupsCount=0
					Call ChildGroupExpose(hChildGroup)
					'  Обнаружено ли в процессе раскрытия дочерней группы, что вложенность дочерних групп превышает допустимую?
					If Me.nChildGroupsCount>Me.nChildGroupsDeepMax Then
						'  Устанавливаем признак зацикливания и прерываем работу функции
						Me.bGroupCycling=True
						Exit Sub
					End If
				End If
			End If
		End ForAll

	End Sub


	%REM
	Сложение двух массивов
	Входные параметры:
		a1	in		первый массив
		a2	in		второй массив
	Выход:
		массив a1+a2

	Тестирование проведено 05-06-2018

	%END REM
	Private Function ArrayAddition(a1 As Variant, a2 As Variant) As Variant

		Dim ba1Empty As Boolean
		Dim ba2Empty As Boolean


		ba1Empty=IsArrayEmpty(a1)
		ba2Empty=IsArrayEmpty(a2)

		If ba1Empty And ba2Empty Then Exit Function
		If ba1Empty Then
			ArrayAddition=a2
			Exit Function
		End If
		If ba2Empty Then
			ArrayAddition=a1
			Exit Function
		End If

		ArrayAddition=ArrayAppend(a1, a2)

	End Function


	%REM
		Добавить элемент в конец текстового массива
		Входные параметры:
			in		a	массив к которому будет производиться добавление
			in		s	значение нового добавляемого элемента
	%END REM
	Private Sub AddElementToStringArray(a As Variant, s As String)

		If IsArrayEmpty(a) Then
			Dim aRes(0) As String
			aRes(0)=s
			a=aRes
		Else
			a=ArrayAppend(a, s)
		End If

	End Sub


	%REM
		Добавить элемент в конец массива типа NotesDocument
		Входные параметры
			in		a			массив к которому будет производиться добавление
			in		hDoc		ссылка на новый член массива
	%END REM
	Private Sub AddElementToNotesDocumentArray(a As Variant, hDoc As NotesDocument)

		If IsArrayEmpty(a) Then
			Dim aRes(0) As NotesDocument
			Set aRes(0)=hDoc
			a=aRes
		Else
			a=ArrayAppend(a, hDoc)
		End If

	End Sub


	'  ***************************** БАЗОВЫЕ ФУНКЦИИ *****************************

	'  Является ли группой?
	Private Function IsGroup(hNames As NotesDatabase, cGroup As String) As NotesDocument
		Dim hGroupView As NotesView
		Set hGroupView=hNames.GetView("$RegisterGroups")
		Set IsGroup=hGroupView.GetDocumentByKey(GetAbbreviatedName(cGroup), True)
	End Function


	'  Получить common-имя
	Private Function GetCommonName(cAccount As String) As String
		Dim hN As New NotesName(cAccount)
		GetCommonName=hN.Common
	End Function

	'  Получить сокращенное имя
	Private Function GetAbbreviatedName(cAccount As String) As String
		Dim hN As New NotesName(cAccount)
		GetAbbreviatedName=hN.Abbreviated
	End Function

	%REM
		Получить иерархическую часть имени(то, что после common-имени)
		(иерархия возвращается без начального слеша)

		Примеры:
			Имя: "Ivan I Testov/KIB"		Возврат: "KIB"
			Имя: "CN=Ivan I Testov/O=KIB"	Возврат: "O=KIB"
			Имя: "Ivan I Testov"			Возврат: ""
	%END REM
	Private Function GetHierarchyName(cAccount As String) As String
		Dim vAccount As Variant
		vAccount=Evaluate({@Name([HIERARCHYONLY]; "} & cAccount & {")})
		GetHierarchyName=vAccount(0)
	End Function


	%REM
	Преобразовать имя в каноническую форму
	Тестирование: проведено 11-07-2016
	%END REM
	Private Function GetCanonicalizeName(cAccount As String) As String
		Dim vAccount As Variant
		vAccount=Evaluate({@Name([CANONICALIZE]; "} & cAccount & {")})
		GetCanonicalizeName=vAccount(0)
	End Function


	'  Если у строки есть начальный слеш, то удалить его
	Private Function RemoveFirstSlash(cString)
		If Left(cString,1)="/" Then RemoveFirstSlash=StrRight(cString, "/") Else RemoveFirstSlash=cString
	End Function


	%REM
		Пуст ли массив?
		(под пустотой понимается отсутствие памяти выделенной под массив)

		Если заходит пустая Variant функция вернит True
		Если заходит статический массив функция вернет False
		Если заходит неинициализированный динамический массив функция вернет True
		Если заходит ициализированный динамический массив функция вернет False

		/
		Если заходит не массив, то функция вернет True.
		Т.е. выполнение даннной функции имеет смысл только для массивов.
		/
	%END REM
	Private Function IsArrayEmpty(a As Variant)
		Err=0
		On Error Resume Next
		Dim nBound As Integer
		nBound=UBound(a)
		On Error GoTo 0
		If Err=0 Then IsArrayEmpty=False Else IsArrayEmpty=True
	End Function


	%REM

	Соответствует ли массив имен, который предполагается записать в группу, заданному лимиту
	размера поля Members?

	Входные параметры:
		a			массив
		nLimit		лимит(только в пределах <=32767, в противном случае функция вернет False)

	Возврат: True/False

	Протокол тестирования функции:
		Тест 1: объем передаваемого массива больше 64К						Пройден. Возврат False. Сработало исключение
		Тест 2: объем передаваемого массива больше 32767 но меньше 64К		Пройден. Возврат False. NotesItem.ValueLength>32767. До сравнения с заданным лимитом даже не доходит
		Тест 3: объем передаваемого массива равен 32767						Пройден. Возврат True
		Тест 4: объем передаваемого массива меньше 32767					Пройден. Возврат True
		Тест 5: передаваемый динамический массив неинициализирован			Пройден. Возврат True. ReplaceItemvalue() выполняется без ошибки. В NotesItem.Values пустая строка. NotesItem.Valuelength=2
		Тест 6: передается не массив										Пройден. Возврат True. Передавалась строка. NotesItem.Valuelength установилось в значение длины строки+2

	%END REM
	Private Function ArrayMatchingToGroupLimit(a As Variant, nLimit As Long) As Boolean

		Dim hSession As New NotesSession
		Dim hDoc As NotesDocument
		Dim hMembers As NotesItem


		ArrayMatchingToGroupLimit=False

		'  Данная функция пишется исключительно для работы с группами в адресной книге. Ни одна группа не может
		'  вместить в поле Members объем данных, который соответствует NotesItem.ValueLength=32767(в Ytria: 32765).
		'  Поэтому, передача параметра лимита больше предельного значения, изначально считается ошибкой и
		'  функция, каким бы ни был объем передаваемого массива, возвращает False
		If nLimit>32767 Then Exit Function

		Err=0
		On Error Resume Next
		Set hDoc=hSession.Currentdatabase.CreateDocument()
		Set hMembers=hDoc.ReplaceItemValue("Members", a)
		On Error GoTo 0
		'  Исключение сработает, если размер массива превышает 64K
		If Err<>0 Then Exit Function
		'  Произвести успешное сохранение документа при выполнении условия ниже
		'  невозможно - NotesDocument.Save() даст False
		If hMembers.Valuelength>32767 Then Exit Function
		'  Проверяем допустимость будущего размера поля Members заданному лимиту.
		'  Помним, что Ytria будет всегда показывать на 2 байта меньше.
		If hMembers.Valuelength<=nLimit Then ArrayMatchingToGroupLimit=True

	End Function


	%REM
	Произвести операцию сохрания группы, контролируя при этом ошибки.
	Версия 2.

	Вход:
	in	hGroup				ссылка на группу

	Выход:
		ноль(успешное заверешние)
		код ошибки: Err() или 1000(пользовательский код ошибки) если операция NotesDocument::Save() вернула False

	Тестирование: проведено 23-06-2016

	%END REM
	Private Function SaveGroup(hGroup As NotesDocument) As Integer
		Dim bSave As Boolean

		SaveGroup=0

		'  Сохраняем...
		On Error Resume Next
		Err=0
		bSave=hGroup.Save(True, False)
		If Err()<>0 Then
			SaveGroup=Err()
		Else
			If bSave=False Then
				'  Ошибки времени выполнения нет, однако фунция Save() вернула False. Сигнализируем об этом
				'  пользовательским кодом ошибки 1000
				SaveGroup=1000
			End If
		End If
		On Error GoTo 0

	End Function

	%REM
	Function ReplaceAndSaveGroup
	Description:

	Прописать новое значение Members в группу и сохранить
	Вход:
		in	hGroup				ссылка на группу
		in	aNewGroupMembers	массив, который нужно прописать в поле Members
	Выход:
		код ошибки(Err() при ошибке времени выполения или 1000 если функция NotesDocument::Save() вернула False)

	Тестирование: не нужно
	%END REM
	Private Function ReplaceGroup(hGroup As NotesDocument, aNewGroupMembers As Variant) As Integer
		'  Прописываем новый состав группы
		Call hGroup.ReplaceItemValue("Members", aNewGroupMembers)
		ReplaceGroup=SaveGroup(hGroup)
	End Function


	%REM
		Удалить дубликаты и пустые строки в массиве
		Вход
			in	a	массив
		Выход: массив без дубликатов и пустых строк либо EMPTY
	%END REM
	Private Sub DeleteDublicateAndEmpty(a As Variant)
		Dim i As Long, j As Long
		Dim aTemp() As String
		Dim aEmpty As Variant

		If IsEmpty(a)=False Then
			a=ArrayUnique(a)
			If IsNull(ArrayGetIndex(a, ""))=False Then
				j=-1
				For i=0 To UBound(a)
					If a(i)<>"" Then
						j=j+1
						ReDim Preserve aTemp(j) As String
						aTemp(j)=a(i)
					End If
				Next
				If j>-1 Then a=aTemp Else a=aEmpty
			End If
		End If
	End Sub

End Class

%REM
	Class CAggrGroupNormalizer
	Description: Comments for Class
%END REM
Public Class CAggrGroupNormalizer

	'  Массив ссылок на дочерние группы
	Private aChildGroup As Variant

	'  Ссылка на основную группу(из конструктора)
	Private hMainGroup As NotesDocument

	'  Шаблон(начальная неизменная часть common-имени) дочерних групп(из конструктора)
	Private cChildGroupCNTemplate As String

	'  Иерархическая составляющая дочерних групп(из конструктора)
	Private cChildGroupHierarchy As String


	%REM
		Конструктор
		Входные параметры:
			hMainGroup					ссылка на основную группу
			cChildGroupCNTemplate		шаблон дочерних групп(неизменная часть common-имени)
			cChildGroupHierarchy		иерархическая составляющая дочерних групп
	%END REM
	Sub New(hMainGroup As NotesDocument, cChildGroupCNTemplate As String, cChildGroupHierarchy As String)
		Set Me.hMainGroup=hMainGroup
		Me.cChildGroupCNTemplate=cChildGroupCNTemplate
		Me.cChildGroupHierarchy=cChildGroupHierarchy

		'  Получаем список дочерних групп
		Call GetChildGroups()
	End Sub


	%REM
		Нормализация аггрегированной группы
		Нормализация проходит исходя из принципов:
			- в головной группе должны быть только подгруппы подпадающие под заданный шаблон
			- все группы, подпадающие под шаблон дочерних, должные входить в основную группу
			- в дочерних группах не может быть подгрупп
	%END REM
	Function normalize() As Boolean

		normalize=False

		'  Удаляем из головной группы все члены не подпадающие под шаблон дочерних групп
		If RemoveFromMainGroupWithoutTemplate() Then
			'  Проверяем чтобы все дочерние группы входили в головную группу
			If CheckChildGroupsOccurrences() Then
				If RemoveFromChildGroupsOtherGroups() Then
					normalize=True
				End If
			End If
		End If

	End Function


	%REM
		Пройтись по всем дочерним группам и удалить все найденные члены-группы
		Подгруппы могут быть только в головной группе!
	%END REM
	Function RemoveFromChildGroupsOtherGroups() As Boolean

		Dim i As Integer


		RemoveFromChildGroupsOtherGroups=True

		If IsArrayEmpty(Me.aChildGroup)=False Then
			For i=0 To UBound(Me.aChildGroup)
				If RemoveFromGroupOtherGroups(Me.aChildGroup(i))=False Then
					RemoveFromChildGroupsOtherGroups=False
					Exit For
				End If
			Next
		End If

	End Function

	%REM
		В заданной группе проверить содержимое на наличие групп. Группы удалить.
	%END REM
	Private Function RemoveFromGroupOtherGroups(hGroup As NotesDocument) As Boolean

		Dim aMember As Variant
		Dim aMemberNew As Variant
		Dim i As Integer


		RemoveFromGroupOtherGroups=True

		aMember=hGroup.GetItemValue("Members")
		Call DeleteDublicateAndEmpty(aMember)

		If IsArrayEmpty(aMember)=False Then
			For i=0 To UBound(aMember)
				If IsGroup(hGroup.Parentdatabase, aMember(i)) Is Nothing Then
					Call AddElementToStringArray(aMemberNew, aMember(i))
				End If
			Next

			If ArrayLength(aMember)<>ArrayLength(aMemberNew) Then
				If ReplaceGroup(hGroup, aMemberNew)<>0 Then RemoveFromGroupOtherGroups=False
			End If
		End If

	End Function



	%REM
		Проверка вхождения дочерних групп в головную групппу.
		В случае отсутствия - добавить
		Тестирование проведено 02-08-2018
	%END REM
	Function CheckChildGroupsOccurrences() As Boolean

		Dim i As Integer
		Dim aMainGroupMember As Variant
		Dim aMainGroupMemberNew As Variant


		CheckChildGroupsOccurrences=True

		If IsArrayEmpty(Me.aChildGroup)=False Then

			'  Члены основной группы
			aMainGroupMember=Me.hMainGroup.GetItemValue("Members")
			Call DeleteDublicateAndEmpty(aMainGroupMember)

			'  Предполагаемый новый массив для основной группы
			'  На старте анализа идентичен текущему массиву группы
			aMainGroupMemberNew=aMainGroupMember

			For i=0 To UBound(Me.aChildGroup)
				'  Если дочерняя группа не входит в головную группу, то добавляем ее в предполагаемый новый массив...
				If IsStringArrayMember(aMainGroupMember, Me.aChildGroup(i).GetItemValue("ListName")(0))=False Then
					Call AddElementToStringArray(aMainGroupMemberNew, Me.aChildGroup(i).GetItemValue("ListName")(0))
				End If
			Next

			'  Если было хотя бы одно добавление в предполагаемый новый массив, то сохраняем этот массив в групе
			If ArrayLength(aMainGroupMember)<>ArrayLength(aMainGroupMemberNew) Then
				If ReplaceGroup(Me.hMainGroup, aMainGroupMemberNew)<>0 Then CheckChildGroupsOccurrences=False
			End If

		End If


	End Function


	'  Получить количество элементов в массиве
	Private Function ArrayLength(a As Variant) As Integer
		If IsArrayEmpty(a) Then ArrayLength=0 Else ArrayLength=UBound(a)+1
	End Function

	'  Получить ссылки на все дочерние группы и сохранить их во внутреннем массиве aChildGroup
	'  Тестирование проведено
	Private Sub GetChildGroups

		Dim aEmpty As Variant
		Dim cFormula As String
		Dim hDColl As NotesDocumentCollection
		Dim nBound As Integer
		Dim cAbbrHierarchy As String
		Dim hGroup As NotesDocument
		Dim cChildGroupCN As String
		Dim nNumberTemp As Integer
		Dim aChildGroupTemp() As NotesDocument


		'  Очищаем внутренний массив, в котором будут сохраняться дочерние группы
		Me.aChildGroup=aEmpty

		'  Находим все группы, которые подпадают под шаблон common-имени.
		'  Иерархия в расчет не берется. Ее анализ будет производиться в цикле для каждой найденной группы.
		cFormula={SELECT Form="Group" & @Like( @LowerCase(@Name([CN]; ListName)); "} & LCase(Me.cChildGroupCNTemplate) & {%")}

		Set hDColl=Me.hMainGroup.Parentdatabase.Search(cFormula, Nothing, 0)
		If hDColl.Count>0 Then

			'  Имя каждой найденной группы анализируем на предмет того, что стоит после шаблона.
			'  Если это число - считаем найденную группу дочерней.

			'  Индекс массива с найденными дочерними группами
			nBound=-1

			'  сокращенное имя иерархии без начального слеша для последующего сравнения
			cAbbrHierarchy=LCase(RemoveFirstSlash(GetAbbreviatedName(Me.cChildGroupHierarchy)))

			Set hGroup=hDColl.GetFirstDocument()
			Do While Not (hGroup Is Nothing)
				cChildGroupCN=GetCommonName(hGroup.GetItemvalue("ListName")(0))
				If Len(cChildGroupCN)>Len(Me.cChildGroupCNTemplate) Then
					'  Если иерархия найденной группы, соответствует запрошенной...
					If LCase( RemoveFirstSlash(GetAbbreviatedName(GetHierarchyName(hGroup.GetItemvalue("ListName")(0)))) )=cAbbrHierarchy Then
						'  Пытаемся преобразовать в число, то что находиться правее шаблона.
						'  Преобразование прошло без ошибок? Если да, то считаем что нашли очередную дочернюю группу.
						Err=0
						On Error Resume Next
						nNumberTemp=CInt( Right(cChildGroupCN, Len(cChildGroupCN)-Len(Me.cChildGroupCNTemplate)) )
						On Error GoTo 0
						If Err=0 Then
							'  Запоминаем ссылку на группу
							nBound=nBound+1
							ReDim Preserve aChildGroupTemp(nBound)
							Set aChildGroupTemp(nBound)=hGroup
						End If
					End If
				End If
				Set hGroup=hDColl.GetNextDocument(hGroup)
			Loop

			If nBound>=0 Then
				Me.aChildGroup=aChildGroupTemp
			End If
		End If

	End Sub


	%REM
		Удалить из основной группы любые члены не подпадающие под шаблон дочерней группы
		Тестирование проведено
	%END REM
	Function RemoveFromMainGroupWithoutTemplate() As Boolean

		Dim aMember As Variant
		Dim i As Integer
		Dim cAbbrChildGroupHierarchy As String
		Dim cCN As String, cHierarchy As String
		Dim nNumberTemp As Integer
		Dim bConformity As Boolean
		Dim aNewMember As Variant


		RemoveFromMainGroupWithoutTemplate=True

		aMember=Me.hMainGroup.GetItemValue("Members")
		Call DeleteDublicateAndEmpty(aMember)

		If IsArrayEmpty(aMember)=False Then

			'  Дополнительный контроль: иерархическая составляющая должна быть в сокращенном виде и без начального слеша
			cAbbrChildGroupHierarchy=LCase(RemoveFirstSlash(GetAbbreviatedName(Me.cChildGroupHierarchy)))

			For i=0 To UBound(aMember)
				If Not(IsGroup(Me.hMainGroup.ParentDatabase, aMember(i)) Is Nothing) Then

					cCN=LCase(GetCommonName(aMember(i)))
					cHierarchy=LCase(RemoveFirstSlash(GetAbbreviatedName(GetHierarchyName(aMember(i)))))

					'  Проверяем подходит ли название члена группы под заданный шаблон
					bConformity=False
					If cHierarchy=cAbbrChildGroupHierarchy Then
						If Left(cCN, Len(Me.cChildGroupCNTemplate))=LCase(Me.cChildGroupCNTemplate) Then

							'  Пытаемся преобразовать в число, то что находиться правее шаблона.
							'  Преобразование прошло без ошибок? Если да, то считаем что нашли очередную дочернюю группу.
							Err=0
							On Error Resume Next
							nNumberTemp=CInt( Right(cCN, Len(cCN)-Len(Me.cChildGroupCNTemplate)) )
							On Error GoTo 0
							If Err=0 Then
								bConformity=True
							End If

						End If
					End If

					'  Наполняем массив только теми членами-группами, которые подходят под шаблон
					If bConformity Then Call AddElementToStringArray(aNewMember, CStr(aMember(i)))

				End If
			Next

			'  Сохраняем новый массив в группе
			If IsArrayEmpty(aNewMember) Then
				If ReplaceGroup(Me.hMainGroup, aNewMember)<>0 Then RemoveFromMainGroupWithoutTemplate=False
			Else
				If UBound(aMember)<>UBound(aNewMember) Then
					If ReplaceGroup(Me.hMainGroup, aNewMember)<>0 Then RemoveFromMainGroupWithoutTemplate=False
				End If
			End If

		End If


	End Function


	%REM
		Пуст ли массив?
		(под пустотой понимается отсутствие памяти выделенной под массив)

		Если заходит пустая Variant функция вернит True
		Если заходит статический массив функция вернет False
		Если заходит неинициализированный динамический массив функция вернет True
		Если заходит ициализированный динамический массив функция вернет False

		/
		Если заходит не массив, то функция вернет True.
		Т.е. выполнение даннной функции имеет смысл только для массивов.
		/
	%END REM
	Private Function IsArrayEmpty(a As Variant)
		Err=0
		On Error Resume Next
		Dim nBound As Integer
		nBound=UBound(a)
		On Error GoTo 0
		If Err=0 Then IsArrayEmpty=False Else IsArrayEmpty=True
	End Function

	'  Получить сокращенное имя
	Private Function GetAbbreviatedName(cAccount As String) As String
		Dim hN As New NotesName(cAccount)
		GetAbbreviatedName=hN.Abbreviated
	End Function

	'  Если у строки есть начальный слеш, то удалить его
	Private Function RemoveFirstSlash(cString)
		If Left(cString,1)="/" Then RemoveFirstSlash=StrRight(cString, "/") Else RemoveFirstSlash=cString
	End Function

	'  Является ли группой?
	Private Function IsGroup(hNames As NotesDatabase, cGroup As String) As NotesDocument
		Dim hGroupView As NotesView
		Set hGroupView=hNames.GetView("$RegisterGroups")
		Set IsGroup=hGroupView.GetDocumentByKey(GetAbbreviatedName(cGroup), True)
	End Function

	'  Получить common-имя
	Private Function GetCommonName(cAccount As String) As String
		Dim hN As New NotesName(cAccount)
		GetCommonName=hN.Common
	End Function

	%REM
		Получить иерархическую часть имени(то, что после common-имени)
		(иерархия возвращается без начального слеша)

		Примеры:
			Имя: "Ivan I Testov/KIB"		Возврат: "KIB"
			Имя: "CN=Ivan I Testov/O=KIB"	Возврат: "O=KIB"
			Имя: "Ivan I Testov"			Возврат: ""
	%END REM
	Private Function GetHierarchyName(cAccount As String) As String
		Dim vAccount As Variant
		vAccount=Evaluate({@Name([HIERARCHYONLY]; "} & cAccount & {")})
		GetHierarchyName=vAccount(0)
	End Function

	%REM
		Добавить элемент в конец текстового массива
		Входные параметры:
			in		a	массив к которому будет производиться добавление
			in		s	значение нового добавляемого элемента
	%END REM
	Private Sub AddElementToStringArray(a As Variant, s As String)

		If IsArrayEmpty(a) Then
			Dim aRes(0) As String
			aRes(0)=s
			a=aRes
		Else
			a=ArrayAppend(a, s)
		End If

	End Sub

	%REM
	Function ReplaceAndSaveGroup
	Description:

	Прописать новое значение Members в группу и сохранить
	Вход:
		in	hGroup				ссылка на группу
		in	aNewGroupMembers	массив, который нужно прописать в поле Members
	Выход:
		код ошибки(Err() при ошибке времени выполения или 1000 если функция NotesDocument::Save() вернула False)

	Тестирование: не нужно
	%END REM
	Private Function ReplaceGroup(hGroup As NotesDocument, aNewGroupMembers As Variant) As Integer
		'  Прописываем новый состав группы
		Call hGroup.ReplaceItemValue("Members", aNewGroupMembers)
		ReplaceGroup=SaveGroup(hGroup)
	End Function

	%REM
	Произвести операцию сохрания группы, контролируя при этом ошибки.
	Версия 2.

	Вход:
	in	hGroup				ссылка на группу

	Выход:
		ноль(успешное заверешние)
		код ошибки: Err() или 1000(пользовательский код ошибки) если операция NotesDocument::Save() вернула False

	Тестирование: проведено 23-06-2016

	%END REM
	Private Function SaveGroup(hGroup As NotesDocument) As Integer
		Dim bSave As Boolean

		SaveGroup=0

		'  Сохраняем...
		On Error Resume Next
		Err=0
		bSave=hGroup.Save(True, False)
		If Err()<>0 Then
			SaveGroup=Err()
		Else
			If bSave=False Then
				'  Ошибки времени выполнения нет, однако фунция Save() вернула False. Сигнализируем об этом
				'  пользовательским кодом ошибки 1000
				SaveGroup=1000
			End If
		End If
		On Error GoTo 0

	End Function


	%REM
		Удалить дубликаты и пустые строки в массиве
		Вход
			in	a	массив
		Выход: массив без дубликатов и пустых строк либо EMPTY
	%END REM
	Private Sub DeleteDublicateAndEmpty(a As Variant)
		Dim i As Long, j As Long
		Dim aTemp() As String
		Dim aEmpty As Variant

		If IsEmpty(a)=False Then
			a=ArrayUnique(a)
			If IsNull(ArrayGetIndex(a, ""))=False Then
				j=-1
				For i=0 To UBound(a)
					If a(i)<>"" Then
						j=j+1
						ReDim Preserve aTemp(j) As String
						aTemp(j)=a(i)
					End If
				Next
				If j>-1 Then a=aTemp Else a=aEmpty
			End If
		End If
	End Sub


	%REM
		Входит ли заданная строка в массив?
	%END REM
	Private Function IsStringArrayMember(a As Variant, s As String) As Boolean
		IsStringArrayMember=False
		If IsArrayEmpty(a)=False Then
			IsStringArrayMember=Not IsNull(ArrayGetIndex(a, s, 5))
		End If
	End Function

End Class

%REM
	Class CGroupExpose
	Description:

	"Раскрыватель" группы


	Параметры конструктора:
		nChildGroupsDeepMax	максимальная глубина вложенности дочерних групп

	Функции-члены и свойства:
		Expose(hGroup As NotesDocument) As Variant
			Функция "Раскрыть группу"(входящий параметр: ссылка на группу)
			Возвращает массив членов раскрытой группы или EMPTY. Обязательно после вызова группы
			анализировать признак зацикленности! Только в случае если нет зацикленности результат
			будет верным.

		ExposeMembers(aMembers As Variant) As Variant
			Функция "Раскрыть группу"(входящий параметр: массив из поля Members)

		GetMembersCount(hGroup As NotesDocument) As Integer
			Функция "Получить количество членов раскрытой группы"
			Функция возвращает ноль в двух случаях. Первый - когда группа действительно пуста.
			И второй - когда в наличии зацикливание группы и подсчитать точное количество членов
			группы не представляется возможным. Поэтому, в случае необходимости, после вызова
			функции проверить значение свойства Cycling.

		property Cycling
			Свойство "Зацикленность группы"
			Boolean. Только чтение. Прописывается каждый раз при вызове в функциях "Раскрыть группу".



%END REM
Public Class CGroupExposer

	'  Максимально допустимая глубина вложенности дочерних групп. Устанавливается через конструктор.
	nChildGroupsDeepMax As Integer

	'  Счетчик дочерних групп
	nChildGroupsCount As Integer

	'  Признак зацикливания группы
	bGroupCycling As Boolean


	'  Конструктор
	Sub New(nChildGroupsDeepMax As Integer)
		Me.nChildGroupsDeepMax=nChildGroupsDeepMax
	End Sub

	'  Получить признак зацикливания группы
	Property Get Cycling As Boolean
		Cycling=Me.bGroupCycling
	End Property

	'  Получить сокращенное имя
	Private Function GetAbbreviatedName(cAccount As String) As String
		Dim hN As New NotesName(cAccount)
		GetAbbreviatedName=hN.Abbreviated
	End Function

	'  Является ли группой?
	Private Function IsGroup(hNames As NotesDatabase, cGroup As String) As NotesDocument
		Dim hGroupView As NotesView
		Set hGroupView=hNames.GetView("$RegisterGroups")
		Set IsGroup=hGroupView.GetDocumentByKey(GetAbbreviatedName(cGroup), True)
	End Function


	%REM
	Добавить в конец текстового массива новый элемент
	Входные параметры:
		in		a	массив к которому будет производиться добавление
		in		s	значение нового добавляемого элемента

	Выход: суммарный массив

	Примечание:
	Особенность функции в том, что в качестве первого параметра может передаваться EMPTY или неинициализированный
	динамический массив. В этом случае функция все равно вернет массив из одного члена.
	%END REM
	Private Function AddElementToArray(a As Variant, s As String) As Variant

		If IsArrayEmpty(a) Then
			Dim aRes(0) As String
			aRes(0)=s
			AddElementToArray=aRes
		Else
			AddElementToArray=ArrayAppend(a, s)
		End If

	End Function


	%REM
	Добавить в конец текстового массива другой текстовый массив
	Входные параметры:
		in		a1	первый массив
		in		a2	второй массив
	Выход:
		массив a1+a2
	%END REM
	Private Function AddArrayToArray(a1 As Variant, a2 As Variant) As Variant
		Dim ba1Empty As Boolean
		Dim ba2Empty As Boolean


		ba1Empty=IsArrayEmpty(a1)
		ba2Empty=IsArrayEmpty(a2)

		If ba1Empty And ba2Empty Then Exit Function
		If ba1Empty Then
			AddArrayToArray=a2
			Exit Function
		End If
		If ba2Empty Then
			AddArrayToArray=a1
			Exit Function
		End If

		AddArrayToArray=ArrayAppend(a1, a2)
	End Function


	%REM
	Пуст ли массив?
	(под пустотой понимается отсутствие памяти выделенной под массив)

	Если заходит пустая Variant функция вернит True
	Если заходит статический массив функция вернет False
	Если заходит неинициализированный динамический массив функция вернет True
	Если заходит ициализированный динамический массив функция вернет False

	/
	Если заходит не массив, то функция вернет True.
	Т.е. выполнение даннной функции имеет смысл только для массивов.
	/
	%END REM
	Private Function IsArrayEmpty(a As Variant)
		Err=0
		On Error Resume Next
		Dim nBound As Integer
		nBound=UBound(a)
		On Error GoTo 0
		If Err=0 Then IsArrayEmpty=False Else IsArrayEmpty=True
	End Function


	'  Раскрыть дочернюю группу
	Private Function ChildGroupExpose(hGroup As NotesDocument) As Variant

		Dim cMember As String
		Dim hChildGroup As NotesDocument
		Dim aExposedGroupMembers As Variant
		Dim aGroupMembers As Variant
		Dim aChildGroupMembers As Variant
		Dim nChildGroupsCountOnThisLevel As Integer


		'  Вошли в дочернюю группу. Увеличиваем счетчик...
		Me.nChildGroupsCount=Me.nChildGroupsCount+1
		'  Запоминаем глубину вложенности групп на текущем уровне...
		nChildGroupsCountOnThisLevel=Me.nChildGroupsCount

		'  Обнаружено, что группа имеет признаки зацикливания! Прерываем обработку...
		If Me.nChildGroupsCount>Me.nChildGroupsDeepMax Then
			Exit Function
		End If

		aGroupMembers=hGroup.GetItemValue("Members")
		'  Перебираем всех членов группы. Работаем только с непустыми членами.
		ForAll vMember In hGroup.GetItemValue("Members")
			cMember=CStr(vMember)
			If Trim(cMember)<>"" Then
				Set hChildGroup=IsGroup(hGroup.ParentDatabase, cMember)
				If hChildGroup Is Nothing Then
					'  Найдена не группа
					'  Добавляем найденную учетную запись в конечный массив
					aExposedGroupMembers=AddElementToArray(aExposedGroupMembers, cMember)
				Else
					'  Найдена дочерняя группа. Раскрываем...
					aChildGroupMembers=ChildGroupExpose(hChildGroup)
					If Me.nChildGroupsCount>Me.nChildGroupsDeepMax Then
						Exit Function
					Else
						'  Восстанавливаем уровнь вложенности групп для текущего уровня для
						'  последующих вызовов ChildGroupExpose() на этом уровне
						Me.nChildGroupsCount=nChildGroupsCountOnThisLevel
						'  Добавляем члены раскрытой дочерней группы в конечный массив
						aExposedGroupMembers=AddArrayToArray(aExposedGroupMembers, aChildGroupMembers)
					End If
				End If
			End If
		End ForAll

		'  Устраняем дубликаты...
		If IsEmpty(aExposedGroupMembers)=False Then
			aExposedGroupMembers=ArrayUnique(aExposedGroupMembers)
		End If

		ChildGroupExpose=aExposedGroupMembers
	End Function




	'  Раскрыть группу
	Function Expose(hGroup As NotesDocument) As Variant
		Expose=ExposeMembers(hGroup.GetItemValue("Members"), hGroup.ParentDatabase)
	End Function


	%REM
	Раскрыть группу, используя в качестве входного параметра массив членов группы
	Входные параметры:
		aGroupMembers		массив учетных записей, прочитанный из поля Members какой-либо группы
		hNames				ссылка на адресную книгу
	Возврат:
		"Раскрытый" массив либо EMPTY
	%END REM
	Function ExposeMembers(aGroupMembers As Variant, hNames As NotesDatabase) As Variant

		Dim aChildGroupMembers As Variant
		Dim aMember As Variant, cMember As String
		Dim aExposedGroupMembers As Variant
		Dim hChildGroup As NotesDocument


		'  По умолчанию зацикленности нет.
		Me.bGroupCycling=False

		'  Перебираем все членов группы. Работаем только с непустыми членами.
		ForAll vMember In aGroupMembers
			cMember=CStr(vMember)
			If Trim(cMember)<>"" Then
				'  Является ли рассматриваемый член группы сам группой или нет?
				Set hChildGroup=IsGroup(hNames, cMember)
				If hChildGroup Is Nothing Then
					'  Найдена не группа
					'  Добавляемы найденный член в конечный массив
					aExposedGroupMembers=AddElementToArray(aExposedGroupMembers, cMember)
				Else
					'  Найдена дочерняя группа. Раскрываем ее...
					Me.nChildGroupsCount=0
					aChildGroupMembers=ChildGroupExpose(hChildGroup)
					'  Обнаружено ли в процессе раскрытия дочерней группы, что вложенность дочерних групп превышает допустимую?
					If Me.nChildGroupsCount>Me.nChildGroupsDeepMax Then
						'  Устанавливаем признак зацикливания и прерываем работу функции
						Me.bGroupCycling=True
						Exit Function
					Else
						'  Добавляем члены дочерней группы в оконечный массив
						aExposedGroupMembers=AddArrayToArray(aExposedGroupMembers, aChildGroupMembers)
					End If
				End If
			End If
		End ForAll

		'  Устраняем дубликаты...
		If IsEmpty(aExposedGroupMembers)=False Then
			aExposedGroupMembers=ArrayUnique(aExposedGroupMembers)
		End If

		ExposeMembers=aExposedGroupMembers
	End Function


	'  Получить количество членов группы, предварительно раскрыв группу
	Function GetMembersCount(hGroup As NotesDocument) As Integer

		Dim aMembers As Variant

		GetMembersCount=0
		aMembers=Expose(hGroup)
		If IsEmpty(aMembers)=False Then
				GetMembersCount=UBound(aMembers)+1
		End If

	End Function

End Class
Sub Initialize

	Dim hSession As New NotesSession
	Dim hGroupExposer As CGroupExposer
	Dim hGroupNormalizer As CAggrGroupNormalizer
	Dim hMainGroup As NotesDocument
	Dim vBeginTime As Variant


	On Error GoTo lErrHandler

	'  Фиксируем время начала работы агента. Будет в конце выведено в консоль для анализа
	'  длительности работы агента.
	vBeginTime=Now()

	PrintToConsole "НАЧАЛО"

	'  Получаем параметры агента
	Call ProfileDocReading()

	'  Адресная книга
	'  Продуктив
	Set hNames=hSession.Currentdatabase
	'  --- DEBUG ---
	'Set hNames=New NotesDatabase("EMA/KIB", "temp\names_copy.nsf")
	'If hNames.IsOpen=False Then GoTo lExitSub

	'  Находим группу All/KIB
	Set hMainGroup=IsGroup(hNames, MAINGROUP)
	If hMainGroup Is Nothing Then GoTo lExitSub

	'  Список сотрудников
	Set hStaff=New NotesDatabase(STAFFSERVER, STAFFDB)
	If hStaff.IsOpen=False Then GoTo lExitSub

	'  Раскрываем списки исключений(на тот случай, если среди записей встречаются группы)
	'  После раскрытия, в массивах будет содержаиться исключительно учетные записи пользователей
	Set hGroupExposer=New CGroupExposer(10)
	aExceptionsNever=hGroupExposer.ExposeMembers(EXCEPTIONS_NEVER, hNames)
	If hGroupExposer.Cycling Then Exit Sub
	aExceptionsAlways=hGroupExposer.ExposeMembers(EXCEPTIONS_ALWAYS, hNames)
	If hGroupExposer.Cycling Then Exit Sub

	'  Предварительная нормализация аггрегированной группы
	Set hGroupNormalizer=New CAggrGroupNormalizer(hMainGroup, CHILDGROUPTEMPLATE, CHILDGROUPHIERARHY)
	If hGroupNormalizer.normalize()=False Then GoTo lExitSub Else Delete hGroupNormalizer

	'  Создание объекта для работы с аггрегированной группой
	Set hAGH=New CAggrGroupHandler(hMainGroup, CHILDGROUPTEMPLATE, CHILDGROUPHIERARHY, GROUPLIMIT)
	If hAGH.GroupCycling Then GoTo lExitSub

	'  Наполняем новыми записями группу All/KIB
	If GroupFilling() Then
		'  Убираем лишние записи из группы All/KIB
		Call GroupCleaning()
	End If


	lExitSub:
	PrintToConsole "Длительность работы агента: " + CStr(vBeginTime) + " - " + CStr(Now())
	PrintToConsole "КОНЕЦ"
	Exit Sub


	lErrHandler:
	PrintToConsole "Line: " + CStr(Erl()) + "   Error: " + CStr(Err()) + ": " + Error()
	Exit Sub

End Sub


Sub Terminate

End Sub






















%REM
	Sub GroupFillling
	Description:

	Пробежаться по всем учетным записям типа Person в адресной книге, собрать в отдельный массив
	тех штатных которых еще нет в целевой группе и потом добавить этот массив в целевую группу.

	Тестирование проведено.
%END REM
Function GroupFilling() As Boolean

	Dim hPeopleVw As NotesView
	Dim hPerson As NotesDocument
	Dim cFullName As String
	Dim aAccountForAdd As Variant


	GroupFilling=False

	Set hPeopleVw=hNames.GetView("$People")
	If hPeopleVw Is Nothing Then Exit Function


	'  --- DEBUG ---
	'Dim nCnt As Integer
	'Dim cAllPeopleNumber As String
	'nCnt=0
	'cAllPeopleNumber=CStr(hPeopleVw.Entrycount)


	Set hPerson=hPeopleVw.Getfirstdocument()
	Do While Not(hPerson Is Nothing)

		cFullName=hPerson.GetItemValue("FullName")(0)

		'  Учетная запись входит в исключения "НИКОГДА"?
		'  Если входит, то пропускаем подобную учетную запись и берем следующую...
		If IsArrayEmpty(aExceptionsNever)=False Then
			If IsNull(ArrayGetIndex(aExceptionsNever, cFullName, 5))=False Then GoTo lNextPerson
		End If

		'  Учетная запись входит в исключения "ВСЕГДА"?
		'  Если входит, то добавляем такую учетку в группу All/KIB и идем дальше...
		If IsArrayEmpty(aExceptionsAlways)=False Then
			If IsNull(ArrayGetIndex(aExceptionsAlways, cFullName, 5))=False Then
				'  Если рассматриваемая учетная запись отсутствует в группе, то
				'  накапливаем массив учетных записей, которые нужно будет добавить в группу
				If hAGH.IsMember(cFullName)=False Then
					Call AddElementToStringArray(aAccountForAdd, cFullName)
					GoTo lNextPerson
				End If
			End If
		End If

		'  У учетной записи отсутствует привязка к почте.
		'  Такую запись в любом случае добавлять в группу рассылки нельзя: будут только генерироваться
		'  отбойники на недоставку
		If hPerson.GetItemValue("MailSystem")(0)="100" Then GoTo lNextPerson

		'  Если учетная запись принадлежит штатному сотруднику, сохраняем эту учетку во временном массиве
		If IsStaffMember(hStaff, cFullName) Then
			'  Если рассматриваемая учетная запись отсутствует в группе, то
			'  накапливаем массив учетных записей, которые нужно будет добавить в группу
			If hAGH.IsMember(cFullName)=False Then 	Call AddElementToStringArray(aAccountForAdd, cFullName)
		End If


		lNextPerson:

		'  --- DEBUG ---
		'nCnt=nCnt+1
		'Print "Всего: " + cAllPeopleNumber + "   Обработано: " + CStr(nCnt)


		Set hPerson=hPeopleVw.Getnextdocument(hPerson)
	Loop

	'  Физическое добавление найденных новых учетных записей штатных сотрудников в группу и
	'  параллельное отображение этих учетных записей во внутренних массивах объекта типа CAggrGroupHandler
	If IsArrayEmpty(aAccountForAdd)=False Then GroupFilling=hAGH.AddArr(aAccountForAdd)

End Function

%REM
	Function GetHierarchyName
	Description:

	Получить иерархическую часть имени(то, что после common-имени)
	(иерархия возвращается без начального слеша)

	Примеры:
		Имя: "Ivan I Testov/KIB"		Возврат: "KIB"
		Имя: "CN=Ivan I Testov/O=KIB"	Возврат: "O=KIB"
		Имя: "Ivan I Testov"			Возврат: ""

%END REM
Public Function GetHierarchyName(cAccount As String) As String
	Dim vAccount As Variant
	vAccount=Evaluate({@Name([HIERARCHYONLY]; "} & cAccount & {")})
	GetHierarchyName=vAccount(0)
End Function

%REM
	Sub GroupCleanning
	Description:

	Пройтись по учетным записям раскрытой группы(храняться в объекте типа CAggrGroupHandler) и выяснить соответствует
	ли каждая учетная запись условиям нахождения в этой группе: не входит ли в исключения? есть ли у нее почтовый ящик?
	является ли штатным? Если условия не соблюдены, то удалить подобную учетную запись из группы.

	Тестирвование проведено 25-07-2018
%END REM
Function GroupCleaning() As Boolean

	Dim aMember As Variant
	Dim i As Integer
	Dim aAccountForRemove As Variant
	Dim hPerson As NotesDocument

	' --- DEBUG ---
	'Dim hForRemoveFile As Integer
	'hForRemoveFile=FreeFile()
	'Open "c:\space\out\GroupCleaning\1ForDelete.txt" For Output As hForRemoveFile
	'Dim hGroupBeforeFile As Integer, hGroupAfterFile As Integer
	'hGroupBeforeFile=FreeFile()
	'Open "c:\space\out\GroupCleaning\2GroupBefore.txt" For Output As hGroupBeforeFile
	'hGroupAfterFile=FreeFile()
	'Open "c:\space\out\GroupCleaning\3GroupAfter.txt" For Output As hGroupAfterFile
	'Dim hGroupExposed As New CGroupExposer(10)
	'Dim aGroupExposed As Variant




	GroupCleaning=False

	aMember=hAGH.Members
	If IsArrayEmpty(aMember) Then
		GroupCleaning=True
		Exit Function
	End If

	For i=0 To UBound(aMember)

		'  Учетная запись входит в исключения "НИКОГДА"?
		'  Если входит, то фиксируем во временном массиве эту учетную запись для последующего удаления из группы
		If IsArrayEmpty(aExceptionsNever)=False Then
			If IsNull(ArrayGetIndex(aExceptionsNever, aMember(i), 5))=False Then
				Call AddElementToStringArray(aAccountForRemove, CStr(aMember(i)))
				GoTo lNextPerson
			End If
		End If

		'  Учетная запись входит в исключения "ВСЕГДА"?
		'  Если входит, то переходим к следующей учетке. Эту мы в любом случае удалять не будем.
		If IsArrayEmpty(aExceptionsAlways)=False Then
			If IsNull(ArrayGetIndex(aExceptionsAlways, aMember(i), 5))=False Then
				GoTo lNextPerson
			End If
		End If

		'  Проверяем существование учетной записи в адресной книге и если учетка существует, то
		'  проверяем наличие возможности получения писем
		Set hPerson=IsPerson(hNames, CStr(aMember(i)))
		If hPerson Is Nothing Then
			'  Учетная запись не существует! Зачем она в группе рассылки?
			Call AddElementToStringArray(aAccountForRemove, CStr(aMember(i)))
			GoTo lNextPerson
		Else
			If hPerson.GetItemValue("MailSystem")(0)="100" Then
				'  Учетная запись не имеет почтового ящика. Зачем она в группе рассылки?
				Call AddElementToStringArray(aAccountForRemove, CStr(aMember(i)))
				GoTo lNextPerson
			End If
		End If

		'  Если учетная запись принадлежит не штатному сотруднику, удаляем ее из группы
		If IsStaffMember(hStaff, CStr(aMember(i)))=False Then Call AddElementToStringArray(aAccountForRemove, CStr(aMember(i)))

		lNextPerson:
	Next

	' --- DEBUG ---
	'Call PrintArray(hForRemoveFile, aAccountForRemove)
	'aGroupExposed=hGroupExposed.Expose(IsGroup(hNames, "TestGroup/KIB"))
	'Call PrintArray(hGroupBeforeFile, aGroupExposed)

	If IsArrayEmpty(aAccountForRemove)=False Then GroupCleaning=hAGH.RemoveArr(aAccountForRemove)

	' --- DEBUG ---
	'aGroupExposed=hGroupExposed.Expose(IsGroup(hNames, "TestGroup/KIB"))
	'Call PrintArray(hGroupAfterFile, aGroupExposed)

End Function
%REM
	Sub ProfileDocReading
	Description: Comments for Sub

	Читаем из профайл-документа ТЕКУЩЕЙ БАЗЫ параметры агента

%END REM
Sub ProfileDocReading

	Dim hSession As New NotesSession
	Dim hPDoc As NotesDocument

	Set hPDoc=hSession.Currentdatabase.Getprofiledocument("AllGroupAutoFilling")

	MAINGROUP=hPDoc.GetItemValue("MainGroup")(0)
	CHILDGROUPTEMPLATE=hPDoc.GetItemValue("ChildGroupTemplate")(0)
	CHILDGROUPHIERARHY=hPDoc.GetItemValue("ChildGroupHierarhy")(0)
	GROUPLIMIT=hPDoc.GetItemValue("GroupLimit")(0)
	EXCEPTIONS_NEVER=hPDoc.GetItemValue("Exceptions_Never")
	EXCEPTIONS_ALWAYS=hPDoc.GetItemValue("Exceptions_Always")
	STAFFSERVER=hPDoc.GetItemValue("StaffServer")(0)
	STAFFDB=hPDoc.GetItemValue("StaffDb")(0)

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
		Имя: "iiavnov@alfabank.kiev.ua"			Возврат: "iiavnov@alfabank.kiev.ua"


%END REM
Public Function GetAbbreviatedName(cAccount As String) As String
	Dim hN As New NotesName(cAccount)
	GetAbbreviatedName=hN.Abbreviated
End Function

%REM
	Function IsStaffMember
	Description:

	Принадлежит ли имя учетной записи штатному сотруднику?
	Параметры:
		in	hStaff			ссылка на Список сотрудников
		in	cFullName		имя учетной записи
	Выход:
		True/False		является штатным/не является штатным

	13-08-2019
	Обнаружена и исправлена ошибка включения в группу All/KIB "чистых" УСБ-сотрудников.
	Дело в том, что если взять штатника и добавить в данный документ USB="1", то карточка превращается в "чистого" УСБшника.
	К сожалению, представление ALL_LN_LOOKUP, которое осуществляет поиск по dn-имени не знает об этом признаке, любого
	штатника с полем USB="1" или без него воспринимает как штатного.

	22-01-2021
	Опять проблема по бывшим УСБ-сотрудникам. По представлению ALL_LN_LOOKUP выбираем теперь не первый документ,
	а всю коллекцию. И перебираем ее. Аксима: если втретилась запись не усб-сотрудника(USB<>"1"), то значит
	мы нашли карточку штатного сотрудника Альфа-Банка.
	Эта доработка связана со случаем, когда в представлении	ALL_LN_LOOKUP было по сотруднику две записи в индексе и первая
	была усб-шная, а второя - штатная. Алгоритм, естественно, считал что это сотрудник усб по первой записи и возращал False.

	Тестирование проведено 20-07-2018
	Тестирование проведено 13-08-2019
	Тестирование проведено 22-01-2021

%END REM
Function IsStaffMember(hStaff As NotesDatabase, cFullName As String) As Boolean

	Dim hStaffMemberVw As NotesView
	Dim hCardColl As NotesDocumentCollection
	Dim hCard As NotesDocument


	IsStaffMember=False

	Set hStaffMemberVw=hStaff.GetView("All_LN_LOOKUP")
	If Not(hStaffMemberVw Is Nothing) Then
		Set hCardColl=hStaffMemberVw.GetAllDocumentsByKey(GetAbbreviatedName(cFullName), True)
		If hCardColl.Count>0 Then
			Set hCard=hCardColl.Getfirstdocument()
			Do While Not(hCard Is Nothing)
				'  Как только будет найдена первая же запись не УСБ сотрудника, это
				'  сигнал того, что найдена карточка не уволенного, не внештатного, не длительно-отсутствуюещго и
				'  не усб-сотрудника. Это класический штатный сотрудник Альфа-Банка.
				If hCard.GetItemValue("USB")(0)<>"1" Then
					IsStaffMember=True
					Exit Do
				End If
				Set hCard=hCardColl.Getnextdocument(hCard)
			Loop
		End If
	End If





	%REM
	Dim hStaffMemberVw As NotesView
	Dim hCard As NotesDocument


	IsStaffMember=False

	Set hStaffMemberVw=hStaff.GetView("All_LN_LOOKUP")
	If Not(hStaffMemberVw Is Nothing) Then
		Set hCard=hStaffMemberVw.GetDocumentByKey(GetAbbreviatedName(cFullName), True)
		If Not(hCard Is Nothing) Then
			If hCard.GetItemValue("USB")(0)<>"1" Then IsStaffMember=True
		End If
	End If
	%END REM




	%REM
	Dim hStaffMemberVw As NotesView


	IsStaffMember=False

	Set hStaffMemberVw=hStaff.GetView("All_LN_LOOKUP")
	If Not(hStaffMemberVw Is Nothing) Then
		IsStaffMember=Not(hStaffMemberVw.GetDocumentByKey(GetAbbreviatedName(cFullName), True) Is Nothing)
	End If
	%END REM

End Function

%REM
	Function GetLastName
	Description:

	Получить фамилию из имени(предполагается FullName)

%END REM
Public Function GetLastName(cAccount As String) As String
	GetLastName=StrRightBack(GetCommonName(cAccount), " ")
End Function




%REM
	Существует ли учетная запись типа Person.

	Входные параметры:
		in	hNames			ссылка на адресную книгу
		in	cAccount		имя учетной записи

	Результат:
		ссылка на найденную учетную запись либо Nothing

%END REM
Function IsPerson(hNames As NotesDatabase, cAccount As String) As NotesDocument
	Dim hPeopleVw As NotesView
	Dim aKey(1) As String

	Set hPeopleVw=hNames.GetView("$People")

	aKey(0)=Left(GetLastName(cAccount), 1)
	aKey(1)=cAccount
	Set IsPerson=hPeopleVw.GetDocumentbyKey(aKey, True)
End Function

%REM
	Function IsGroup2
	Description:

	Существует ли заданная группа?
	Вариант 2

	Входные параметры:
		in	hNames	ссылка на АК
		in	cGroup	имя искомой группы

	Выход:
		ссылка на группу, если таковая найдена
		Nothing, если группа не найдена

	Тестирование: проведено 11-07-2016

%END REM
Public Function IsGroup(hNames As NotesDatabase, cGroup As String) As NotesDocument
	Dim hGroupView As NotesView

	Set hGroupView=hNames.GetView("$RegisterGroups")

	'  Вследствии того, что индекс представления $RegisterGroups построен на основе abbreviated имен,
	'  передаем в функцию abbreviated имя искомой группы
	Set IsGroup=hGroupView.GetDocumentByKey(GetAbbreviatedName(cGroup), True)
End Function

%REM
	Sub PrintToConsole
	Description: Comments for Sub
%END REM
Private Sub PrintToConsole(cS As String)
	Print "AllGroupAutoFilling agent  " + cS
End Sub
%REM
	Добавить элемент в конец текстового массива
	Входные параметры:
		in		a	массив к которому будет производиться добавление
		in		s	значение нового добавляемого элемента
%END REM
Sub AddElementToStringArray(a As Variant, s As String)
	If IsArrayEmpty(a) Then
		Dim aRes(0) As String
		aRes(0)=s
		a=aRes
	Else
		a=ArrayAppend(a, s)
	End If
End Sub
%REM
	Function SaveGroup
	Description:

	Произвести операцию сохрания группы, контролируя при этом ошибки.
	Версия 2.

	Вход:
	in	hGroup				ссылка на группу

	Выход:
		ноль(успешное заверешние)
		код ошибки: Err() или 1000(пользовательский код ошибки) если операция NotesDocument::Save() вернула False

	Тестирование: проведено 23-06-2016

%END REM
Function SaveGroup(hGroup As NotesDocument) As Integer
	Dim bSave As Boolean

	SaveGroup=0

	'  Сохраняем...
	On Error Resume Next
	Err=0
	bSave=hGroup.Save(True, False)
	If Err()<>0 Then
		SaveGroup=Err()
	Else
		If bSave=False Then
			'  Ошибки времени выполнения нет, однако фунция Save() вернула False. Сигнализируем об этом
			'  пользовательским кодом ошибки 1000
			SaveGroup=1000
		End If
	End If
	On Error GoTo 0

End Function


%REM
	Function GetCommonName
	Description:

	Получить common-имя

%END REM
Public Function GetCommonName(cAccount As String) As String
	Dim hN As New NotesName(cAccount)
	GetCommonName=hN.Common
End Function




Sub PrintArray(hFile As Integer, a As Variant)

	Dim i As Integer

	If IsArrayEmpty(a) Then
		Print #hFile, "IsArrayEmpty(): True"
		Exit Sub
	End If


	For i=0 To UBound(a)
		Print #hFile, CStr(i)+": "+a(i)
	Next

End Sub

%REM
	Пуст ли массив?
	(под пустотой понимается отсутствие памяти выделенной под массив)

	Если заходит пустая Variant функция вернит True
	Если заходит статический массив функция вернет False
	Если заходит неинициализированный динамический массив функция вернет True
	Если заходит ициализированный динамический массив функция вернет False

	/
	Если заходит не массив, то функция вернет True.
	Т.е. выполнение даннной функции имеет смысл только для массивов.
	/
%END REM
Private Function IsArrayEmpty(a As Variant)
	Err=0
	On Error Resume Next
	Dim nBound As Integer
	nBound=UBound(a)
	On Error GoTo 0
	If Err=0 Then IsArrayEmpty=False Else IsArrayEmpty=True
End Function
%REM
	Function ReplaceAndSaveGroup
	Description:

	Прописать новое значение Members в группу и сохранить
	Вход:
		in	hGroup				ссылка на группу
		in	aNewGroupMembers	массив, который нужно прописать в поле Members
	Выход:
		код ошибки(Err() при ошибке времени выполения или 1000 если функция NotesDocument::Save() вернула False)

	Тестирование: не нужно


%END REM
Function ReplaceGroup(hGroup As NotesDocument, aNewGroupMembers As Variant) As Integer
	'  Прописываем новый состав группы
	Call hGroup.ReplaceItemValue("Members", aNewGroupMembers)
	ReplaceGroup=SaveGroup(hGroup)
End Function




