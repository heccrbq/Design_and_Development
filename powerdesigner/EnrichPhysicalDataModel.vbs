'******************************************************************************
'* File:     	EnrichPhysicalDataModel.vbs
'* Purpose:     The script is designed to add metadata fields and historical tables to the physical data model.
'* Version:   	2.0
'* Author:	 	Bykov D.
'******************************************************************************

Option Explicit
ValidationMode = True

'-----------------------------------------------------------------------------
' DECLARE GLOBAL CONSTANTS [gc]
'-----------------------------------------------------------------------------
' Global Object's Constants
const gcMainLoaderUser = "LOADER_DWH"
const gcHistoricalTablePostfix = "_H"
const gcPrimaryKeyConstraintPrefix = "PK_"
const gcUniqueConstraintPrefix = "AK_"

' Grants
const gcGrantUpdate = "UPDATE"
const gcGrantDelete = "DELETE"
const gcGrantInsert = "INSERT"
const gcGrantSelect = "SELECT"


' let's go
Call Main(ActiveModel)

'-----------------------------------------------------------------------------
' Main procedure
'-----------------------------------------------------------------------------
' Checks specified model and runs enrichModel procedure.
'-----------------------------------------------------------------------------
Private Sub Main(byref model)
   If model is Nothing Then
      MsgBox "There is no current Model."
   ElseIf Not model.IsKindOf(PdPDM.cls_Model) Then
      MsgBox "The current model is not a Physical Data model."
   Else
      enrichModel model
   End If
End Sub

'-----------------------------------------------------------------------------
' EnrichModel prodecure
'-----------------------------------------------------------------------------
' Checks basic parameters of model and creates new tables and service fields..
'-----------------------------------------------------------------------------
Private Sub enrichModel(byref model)
   '#1. Drop & Create historical tables
   recreateHistoricalTables model
   
   '#2. Grant all tables
   grantTablesToLoaderUser model
End Sub

'-----------------------------------------------------------------------------
' RecreateHistoricalTables procedure
'-----------------------------------------------------------------------------
' Shows attributes and collections of object
'-----------------------------------------------------------------------------
Private Sub recreateHistoricalTables(byref model)
   Dim table, column, key, dependency
   Dim histTable, histColumn, histKey, histIndex, histDependency
   
   'Drop All Hist tables
   For Each table in model.Tables
      If Not table.isShortcut And right(table.Code,2) = gcHistoricalTablePostfix Then
         model.Tables.Remove table, true 
         Output "Drop Hist Table: " + table.Code
      End If
   Next
   
   'Add service columns and mandatory flag for PK_ columns 
   For Each table in model.Tables
      If Not table.isShortcut Then
         checkMandatoryFlag table
	      addServiceColumns table, false
      End If
   Next
   
   'Create All Hist tables
   For Each table in model.Tables
      If Not table.isShortcut Then
         'Copy table
         Set histTable = model.Tables.CreateNew
         histTable.Name = table.Name & gcHistoricalTablePostfix
         histTable.Code = table.Code & gcHistoricalTablePostfix
         histTable.Comment = table.Comment
         
         'Copy columns
         For Each column in table.Columns
            createColumn histTable, column.Code, column.DataType, column.Mandatory, column.Comment
         Next
         addServiceColumns histTable, true
         
         'Copy keys & indexes
         For Each key in table.Keys
            'Create keys
            Set histKey = histTable.Keys.CreateNew
            histKey.Primary = key.Primary
            If histKey.Primary Then               
               histKey.Name = histTable.Name
               histKey.Code = histTable.Code
               'PK_%.U27:TABLE%
               histKey.ConstraintName = gcPrimaryKeyConstraintPrefix & left(histTable.Code,27)
            Else
               histKey.Name = key.Name
               histKey.Code = key.Code
               'AK_%.U20:TABLE%_%.U6:AKEY%
               histKey.ConstraintName = gcUniqueConstraintPrefix & left(histTable.Code,20) & "_" & left(histKey.Code,6)               
            End If
         
            For Each column in key.Columns
               For Each histColumn in histTable.Columns
                  If column.Name = histColumn.Name or histColumn.Name = "MD_DATETIME_FROM" Then
                     histKey.Columns.Add histColumn
                  End If
               Next
            Next
           
            'Create indexes
            Set histIndex = histTable.Indexes.CreateNew
            histIndex.Name = histKey.ConstraintName
            histIndex.Code = histKey.ConstraintName
            histIndex.LinkedObject = histKey
         Next
         
         'Create dependencies
         'For Each dependency in table.OutReferences
         '   Set histDependency = histTable.OutReferences.CreateNew
         '   histDependency.Name = dependency.Name
         '   histDependency.ParentTable = table.ParentTable
         'Next
         
         Output "Historical Table " & histTable.Name & " created"
      End If
   Next
End Sub

'-----------------------------------------------------------------------------
' GrantTablesToLoaderUser procedure
'-----------------------------------------------------------------------------
' Grants permissions on table to Loader User
'-----------------------------------------------------------------------------
Private Sub grantTablesToLoaderUser(byref model)
   Dim table, perm, loaderUser, loaderGrant
   
   'Init variables
   Set loaderUser = findUser(model, gcMainLoaderUser)
   loaderGrant = gcGrantUpdate & "," & gcGrantDelete & "," & gcGrantInsert & "," & gcGrantSelect
      
   If Not loaderUser is Nothing Then
      'Loop on tables
      For Each table in model.Tables
         If Not table.isShortcut Then
            'Delete permissions if they exist. Only gcMainLoaderUser has to have access to tables
            If table.Permissions.Count > 0 Then
         	   For Each perm in table.Permissions
                  perm.Delete
                  Output "Permission on table " & table.Name & " droppped"
               Next
            End If
   
            'Create permissions. We know, that gcMainLoaderUser has to have access to load data into CORE schema
            Set perm = table.Permissions.CreateNew
            Set perm.DBIdentifier = loaderUser
            perm.grant = loaderGrant
            Output "Table " & table.Name & " was granted to user " & loaderUser.Name & " with permissions: " & loaderGrant
         End If
      Next
   End If
End Sub

'-----------------------------------------------------------------------------
' FindUser function
'-----------------------------------------------------------------------------
' Finds specified user in model. Returns user if found.
'-----------------------------------------------------------------------------
Private Function findUser(byref model, byref name)
   Dim user
   
   'Find user and exit if found
   For Each user in model.Users
      If user.Name = name Then
	      Set FindUser = user
		   Output "User " & name & " found."
         Exit Function
	   End If
   Next
   
   'If user wasn't found
   Set FindUser = Nothing
   Output "User " & name & " not found."
End Function

'-----------------------------------------------------------------------------
' CheckMandatoryFlag procedure
'-----------------------------------------------------------------------------
' Adds not null constraint for PK_* columns
'-----------------------------------------------------------------------------
Private Sub checkMandatoryFlag(byref table)
   Dim column
   
   'We have to add not null constraint for PK_* columns. Nothing else
   For Each column in table.Columns
      If left(column.Code,3) = gcPrimaryKeyConstraintPrefix and Not column.Mandatory Then
         column.Mandatory = true
      End If
   Next
End Sub

'-----------------------------------------------------------------------------
' AddServiceColumns procedure
'-----------------------------------------------------------------------------
' Adds MD_* service columns into table
'-----------------------------------------------------------------------------
Private Sub addServiceColumns(byref table, isHistoricalTable)
   Dim code, dataType, comment

   code = "MD_ID_ETL_LOG"
   dataType = "INTEGER"
   comment = "Идентификатор экземпляра процесса, изменившего запись последним."
   createColumn table, code, dataType, true, comment

   code = "MD_FLAG_DELETED"
   dataType = "VARCHAR2(1)"
   comment = "Признак удаленной записи: Y - удалена, N - нет."
   createColumn table, code, dataType, true, comment

   code = "MD_CODE_SOURCE_SYSTEM"
   dataType = "VARCHAR2(10)"
   comment = "Код системы источника."
   createColumn table, code, dataType, true, comment
	
	If isHistoricalTable Then
	   code = "MD_DATETIME_FROM"
      dataType = "DATE"
      comment = "Дата изменения записи в источнике."
      createColumn table, code, dataType, true, comment
	End If
End Sub

'-----------------------------------------------------------------------------
' CreateColumn procedure
'-----------------------------------------------------------------------------
' Creates a new service column and sets common attributes
'-----------------------------------------------------------------------------
Private Sub createColumn(byref table, code, dataType, mandatory, comment)
	Dim column
   
   If Not isColumnExist(table, code) Then
	   Set column = table.Columns.CreateNew
		column.Name = code
		column.Code = code
      column.Mandatory = mandatory
		column.DataType = dataType
		column.Comment = comment
	End If
End Sub

'-----------------------------------------------------------------------------
' IsColumnExist function
'-----------------------------------------------------------------------------
' Find specified column amongs columns of table. Returns true if found
'-----------------------------------------------------------------------------
Private Function isColumnExist(byref table, code)
	Dim col
	
	'Return true if column is exist
   For Each col in table.Columns
	   If col.Code = code Then
         isColumnExist = true
         Exit Function
      End If
	Next
	
   isColumnExist = false
End Function

'-----------------------------------------------------------------------------
' viewObj procedure
'-----------------------------------------------------------------------------
' Shows attributes and collections of object
'-----------------------------------------------------------------------------
Private Sub viewObj(byref obj)
   Dim metaClass, attr, coll
   
   Set metaClass = obj.MetaClass   
   Output "Metaclass: " & metaclass.PublicName 
   Output "Parent: " & metaclass.Parent.PublicName
   Output "Metalibrary: " & metaclass.Library.PublicName
   
   Output "Attributes:"   
   For Each attr in metaClass.Attributes
      Output " - " + attr.PublicName
   Next
   
   Output "Collections:"   
   For each coll in metaClass.Collections
      Output " - " + coll.PublicName
   Next   
End Sub
