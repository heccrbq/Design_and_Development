'******************************************************************************
'* Файл:     Обогащение_вер_1.vbs
'* Цель:     Скрипт предназначен для добавления полей метаданных и исторических таблиц в физическую модель данных.
'* Версия:   1.0
'* Автор:	 Быков Д.
'******************************************************************************

Option Explicit

'-----------------------------------------------------------------------------
' Main function
'-----------------------------------------------------------------------------

ValidationMode = True

Dim mdl ' the current model

' get the current active model
Set mdl = ActiveModel
If (mdl Is Nothing) Then
   MsgBox "There is no current Model."
ElseIf Not mdl.IsKindOf(PdPDM.cls_Model) Then
   MsgBox "The current model is not a Physical Data model."
Else
   ProcessFolder mdl
End If


'--------------------------------------------------------------------------------
Private Sub ProcessFolder(byref folder)
   Dim Tab, nTab 
   Dim Col, nCol, kCol 
   Dim Key, iKey, nKey, nInd
   Dim name, code

   ' Get the class diagram
   Dim diagram
   Dim sym
   Set diagram = folder.PhysicalDiagrams.Item(0)

   For each Tab in folder.Tables
      If not Tab.isShortcut then
         If right(Tab.Code,2) = "_H" then
            Output "Drop Hist Table: " + Tab.Code
            folder.Tables.Remove Tab, true
         end if
      End If
   Next
   
   For each Tab in folder.Tables
      If not Tab.isShortcut then
         
		   name = Tab.Name + "_H"
         code = Tab.Code + "_H"
         If Not isTable(folder, name, code) Then         
            Output "Create Hist Table: " + code
            
            set nTab = folder.Tables.CreateNew
            nTab.Name = name
            nTab.Code = code
            nTab.Comment = Tab.Comment
            
            'вывести на диаграму
            'Set sym = diagram.AttachObject(nTab)
            
            Output " - Add columns." 
            For each Col in tab.Columns
				   set nCol = nTab.Columns.CreateNew
				   nCol.Name = Col.Name
				   nCol.Code = Col.Code
				   nCol.DataType = Col.DataType
				   nCol.Comment = Col.Comment
            Next
			   
            Output " - Add meta columns." 
			   addColumnsMD_H nTab
            
            'получить первичный ключь исходной таблицы
            Key = "null"
            For each iKey in Tab.Keys
               if iKey.Primary = true then
                  set Key = iKey
                  exit for
               end if
            next
            
            if Key <> "null" then
               Output " - Add primary key." 
               set nKey = nTab.Keys.CreateNew            
               nKey.Name = name
   				nKey.Code = name
               nKey.Primary = true
               'nKey.Clustered = false
               nKey.ConstraintName = "PK_" + name
               For each Col in nTab.Columns
                  For each kCol in Key.Columns
                     if Col.Code = kCol.Code then
   				         nKey.Columns.Add Col
                     end if
                  next
                  if Col.Code = "MD_DATETIME_FROM" then
   				      nKey.Columns.Add Col
                  end if
               Next

               Output " - Add primary key index." 
               set nInd = nTab.Indexes.CreateNew 
               nInd.Name = "PK_" + name
				nInd.Code =  "PK_" + name
               nInd.LinkedObject = nKey
               nInd.Unique = true
               'nInd.Clustered = false
               For each Col in nTab.Columns
                  For each kCol in Key.Columns
                     if Col.Code = kCol.Code then
   				         set nCol = nInd.IndexColumns.CreateNew
                        set nCol.Column = Col
                     end if
                  next
                  if Col.Code = "MD_DATETIME_FROM" then
                     set nCol = nInd.IndexColumns.CreateNew
                     set nCol.Column = Col
                  end if
               Next            
            end if
         end if
         
         Output "Add meta columns into: " + Tab.Name
         addColumnsMD tab
      End If
   Next
End Sub

'--------------------------------------------------------------------------------
Private Function isTable(byref folder, name, code)
	Dim vTab
	For each vTab in folder.Tables
		If Not vTab.isShortcut Then
			If vTab.Name = name or vTab.Code = code Then
				isTable = true
				Exit Function
			End If
		End If
	Next   
	isTable = false
End Function

'--------------------------------------------------------------------------------
Private Function isColumn(byref tab, name, code)
	Dim vCol
	For each vCol in tab.Columns
		If vCol.Name = name or vCol.Code = code Then
            isColumn = true
            Exit Function
        End If
	Next 
    isColumn = false
End Function

'--------------------------------------------------------------------------------
Private Sub addColumnsMD(byref tab)
	Dim nCol, name

    name = "MD_ID_ETL_LOG"
    if not isColumn(tab, name, name) then
		set nCol = tab.Columns.CreateNew
        nCol.Name = name
        nCol.Code = name
        nCol.DataType = "integer"
        nCol.Comment = "Идентификатор экземпляра процесса, изменивший запись последним."
    end if

    name = "MD_FLAG_DELETED"
    if not isColumn(tab, name, name) then
		set nCol = tab.Columns.CreateNew
        nCol.Name = name
        nCol.Code = name
        nCol.DataType = "varchar2(1 char)"
        nCol.Comment = "Признак удаленной запись: Y - удалена, N - нет."
    end if

    name = "MD_CODE_SOURCE_SYSTEM"
    if not isColumn(tab, name, name) then
		set nCol = tab.Columns.CreateNew
        nCol.Name = name
        nCol.Code = name
        nCol.DataType = "varchar2(10 char)"
        nCol.Comment = "Код системы источника."
    end if	
	
End Sub

'--------------------------------------------------------------------------------
Private Sub addColumnsMD_H(byref tab)
	Dim nCol, name

    addColumnsMD tab
    
    name = "MD_DATETIME_FROM"
    if not isColumn(tab, name, name) then
		set nCol = tab.Columns.CreateNew
        nCol.Name = name
        nCol.Code = name
        nCol.DataType = "date"
        nCol.Comment = "Дата изменения записи в источнике."
    end if

End Sub

'--------------------------------------------------------------------------------
' показывает атрибуты и колекции объкта
Private Sub viewObj(byref obj)
   Dim metaclass, attr, coll
   
   Set metaclass = obj.MetaClass
   Output "Metaclass: " + metaclass.PublicName 
   Output "Parent: " + metaclass.Parent.PublicName
   Output "Metalibrary: " + metaclass.Library.PublicName
   Output "Attributes:"
   For each attr in metaclass.attributes
      Output " - " + attr.PublicName
   Next
   Output "Collections:"
   For each coll in metaclass.collections
      Output " - " + coll.PublicName
   Next
   
End Sub
