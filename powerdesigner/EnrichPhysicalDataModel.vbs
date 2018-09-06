'******************************************************************************
'* File:     	EnrichPhysicalDataModel.vbs
'* Purpose:     The script is designed to enrich the physical data model - creating partitioned historical tables and adding sorted metadata fileds.
'* Version:   	3.2
'* Author:	 	Bykov D.
'******************************************************************************

Option Explicit
ValidationMode = True
InteractiveMode = im_Batch

'------------------------------------------------------------------------------
' DECLARE GLOBAL CONSTANTS [gc]
'------------------------------------------------------------------------------
' Global Object's Constants
const gcMainUser = "CORE"
const gcMainLoaderUser = "LOADER_DWH"
const gcHistoricalTablePostfix = "_H"
const gcPrimaryKeyConstraintPrefix = "PK_"
const gcUniqueConstraintPrefix = "AK_"

' Partitioning
const gcPartitionColumNameHistTable = "MD_DATETIME_FROM"
const gcPartitionRangeIntervalHistTable = "INTERVAL '1' MONTH"
const gcFirstPartitionName = "PARTMM_01"
const gcFirstPartitionCondition = "VALUES LESS THAN (TO_DATE('2000-01-01 00:00:00', 'YYYY-MM-DD HH24:MI:SS', 'NLS_CALENDAR=GREGORIAN'))"
const gcTypeOfPartitionedIndexes = "LOCAL"   '[GLOBAL | LOCAL]

' Grants
const gcGrantUpdate = "UPDATE"
const gcGrantDelete = "DELETE"
const gcGrantInsert = "INSERT"
const gcGrantSelect = "SELECT"


' let's go
Call Main(ActiveModel)

'------------------------------------------------------------------------------
' The procedure checks the specified model and starts the procedure to enrich the model
'------------------------------------------------------------------------------
Private Sub Main(byref model)
   If model Is Nothing Then
      MsgBox "There is no current Model."
   ElseIf Not model.IsKindOf(PdPDM.cls_Model) Then
      MsgBox "The current model is not a Physical Data model."
   Else
      EnrichModel model
   End If
End Sub

'------------------------------------------------------------------------------
' The procedure checks the basic parameters of the model and creates hist tables
'------------------------------------------------------------------------------
Private Sub EnrichModel(byref model)
   '#1. Create Loader User
   CreateUserIfNotExist model

   '#2. Drop & Create historical tables
   RecreateHistoricalTables model
   
   '#3. Disable all references
   DisableReferences model
   
   '#4. Grant all tables
   GrantTablesToLoaderUser model
End Sub

'------------------------------------------------------------------------------
' The procedure creates a new user if he doesn't exist
'------------------------------------------------------------------------------
Private Sub CreateUserIfNotExist(byref model)
   Dim user, pwdUser
   Dim userArray : userArray = Array(gcMainUser, gcMainLoaderUser)
   
   'Create main loader user
   For Each user in userArray
      If FindUser(model, user) Is Nothing Then
         Output "Create User : " & user
      
         Set pwdUser = model.Users.CreateNew
         pwdUser.Name = user
         pwdUser.Code = user
      Else
         Output "User " & user & " found."
      End If
   Next
End Sub

'------------------------------------------------------------------------------
' The procedure creates historical tables, keys, constraints and indexes
'------------------------------------------------------------------------------
Private Sub RecreateHistoricalTables(byref model)
   Dim table, histTable
   
   'Drop all hist tables
   DropAllHistTables model
   
   'Add service columns and mandatory flag for PK_* columns.
   EnrichTables model   
   
   'Create all hist tables. Using loop on snapshot table because hist table is a copy of snapshot table
   For Each table In model.Tables
      If Not table.isShortcut Then         
         Output "Create Hist Table : " & table.Name & gcHistoricalTablePostfix
      
         'Copy table
         Set histTable = CopyTableToHist(model, table)
         
         'Copy columns
         CopyColumnsToHist table, histTable
         
         'Partitioning. Split table by range using as a part_key_column gcPartitionColumNameHistTable
         PartitionByRangeInterval histTable
         
         'Copy keys & indexes
         CopyKeysAndIndexesToHist table, histTable         
         
         'Copy outer dependencies
         CopyReferences model, table, histTable
      End If
   Next
End Sub

'------------------------------------------------------------------------------
' The procedure disables all reference constraints on tables
'------------------------------------------------------------------------------
Private Sub DisableReferences(byref model)
   Dim ref
  
   'All references have to have disable and novalidate options
   For Each ref In model.References
      Output "Disable Reference : " & ref.Name
      
      ref.setExtendedAttribute "Validate", false
      If Not ref.getExtendedAttribute("Disable") Then
         ref.setExtendedAttribute "Disable", true         
      End If 
  Next
End Sub

'------------------------------------------------------------------------------
' The procedure grants permissions on table to Loader User
'------------------------------------------------------------------------------
Private Sub GrantTablesToLoaderUser(byref model)
   Dim table, perm, loaderUser, loaderGrant
   
   'Init variables
   Set loaderUser = FindUser(model, gcMainLoaderUser)
   loaderGrant = gcGrantUpdate & "," & gcGrantDelete & "," & gcGrantInsert & "," & gcGrantSelect
   
   'We know, that gcMainLoaderUser has to have access to load data into CORE schema
   If Not loaderUser Is Nothing Then
      'Loop on tables
      For Each table In model.Tables
         If Not table.isShortcut Then            
            'Delete permissions of incorrect user. Only gcMainLoaderUser has to have access to tables
            For Each perm In table.Permissions
               If (perm.DBIdentifier.Code <> loaderUser.Code) or (perm.DBIdentifier.Code = loaderUser.Code and perm.Grant <> loaderGrant) Then
                  Output "Revoke Grant : Table[" & table.Name & "] User[" & perm.DBIdentifier.Code & "] Grant[" & perm.Grant & "]"
                  perm.Delete
               End If
            Next
            
            'It means gcMainLoaderUser hasn't been added into Permissions block yet
            If table.Permissions.Count = 0 Then            
               'Create permissions
               Output "Grant Table : " & table.Name
               Set perm = table.Permissions.CreateNew
               Set perm.DBIdentifier = loaderUser
               perm.Grant = loaderGrant         	   
            End If
         End If
      Next
   End If
End Sub

'------------------------------------------------------------------------------
' The function finds specified user in model. Returns user if found.
'------------------------------------------------------------------------------
Private Function FindUser(byref model, user)
   'Find user. Return Nothing if not found
   Set FindUser = model.FindChildByCode(user, PdPDM.cls_User)
End Function

'------------------------------------------------------------------------------
' The procedure Drops all hist tables from specified model
'------------------------------------------------------------------------------
Private Sub DropAllHistTables(byref model)
   Dim table
   
   For Each table in model.Tables
      If Not table.isShortcut And IsHistoricalTable(table) Then
         Output "Drop Hist Table : " + table.Code
         model.Tables.Remove table, true         
      End If
   Next
End Sub

'------------------------------------------------------------------------------
' The procedure adds mandatory flag for PK_* columns and adds service columns
'------------------------------------------------------------------------------
Private Sub EnrichTables(byref model)
   Dim table
   
   For Each table In model.Tables
      If Not table.isShortcut Then
         CheckMandatoryFlag table
	      AddServiceColumns table
      End If
   Next
End Sub

'------------------------------------------------------------------------------
' The procedure adds not null constraint for PK_* columns
'------------------------------------------------------------------------------
Private Sub CheckMandatoryFlag(byref table)
   Dim column
   
   'We have to add not null constraint for all PK_* columns. No one else
   For Each column In table.Columns
      If left(column.Code,3) = gcPrimaryKeyConstraintPrefix and Not column.Mandatory Then
         column.Mandatory = true
      End If
   Next
End Sub

'------------------------------------------------------------------------------
' The function creates new hist table based on sourceTable
'------------------------------------------------------------------------------
Private Function CopyTableToHist(byref model, byref sourceTable)
   Dim histTable

   'As a prefix we'will use gcHistoricalTablePostfix global constant
   Set histTable = model.Tables.CreateNew
   histTable.Name = sourceTable.Name & gcHistoricalTablePostfix
   histTable.Code = sourceTable.Code & gcHistoricalTablePostfix
   histTable.Comment = sourceTable.Comment
   
   Set CopyTableToHist = histTable
End Function

'------------------------------------------------------------------------------
' The procedure copies all columns from sourceTable to TargetTable and creates missing service columns
'------------------------------------------------------------------------------
Private Sub CopyColumnsToHist(byref sourceTable, byref targetTable)
   Dim column
   
   'Loop on columns of sourceTable
   For Each column In sourceTable.Columns
      CreateColumn targetTable, column.Code, column.DataType, column.Mandatory, column.Comment
   Next
   'Add missing service columns
   AddServiceColumns targetTable
End Sub

'------------------------------------------------------------------------------
' The procedure adds MD_* service columns into specified table
'------------------------------------------------------------------------------
Private Sub AddServiceColumns(byref table)
   Dim code, dataType, comment, mandatory : mandatory = true

   code = "MD_ID_ETL_LOG"
   dataType = "INTEGER"
   comment = "Идентификатор экземпляра процесса, изменившего запись последним."
   CreateColumn table, code, dataType, mandatory, comment

   code = "MD_FLAG_DELETED"
   dataType = "VARCHAR2(1)"
   comment = "Признак удаленной записи: Y - удалена, N - нет."
   CreateColumn table, code, dataType, mandatory, comment

   code = "MD_CODE_SOURCE_SYSTEM"
   dataType = "VARCHAR2(10)"
   comment = "Код системы источника."
   CreateColumn table, code, dataType, mandatory, comment
	
	If IsHistoricalTable(table) Then
	   code = gcPartitionColumNameHistTable
      dataType = "DATE"
      comment = "Дата изменения записи в источнике."
      CreateColumn table, code, dataType, mandatory, comment
	End If
End Sub

'------------------------------------------------------------------------------
' The function returns true if table is historical
'------------------------------------------------------------------------------
Private Function IsHistoricalTable(byref table)
	If right(table.Code,2) = gcHistoricalTablePostfix Then
      IsHistoricalTable = true
   Else
      IsHistoricalTable = false
   End If
End Function

'------------------------------------------------------------------------------
' The procedure creates a new service column and sets common attributes
'------------------------------------------------------------------------------
Private Sub CreateColumn(byref table, code, dataType, mandatory, comment)
	Dim column
   
   'If column doesn't exist then create the new
   If table.FindChildByCode(code, PdPDM.cls_column) Is Nothing Then
	   Set column = table.Columns.CreateNew
		column.Name = code
		column.Code = code
      column.Mandatory = mandatory
		column.DataType = dataType
		column.Comment = comment
	End If
End Sub

'------------------------------------------------------------------------------
' The procedure splits the table into partitions by range
'------------------------------------------------------------------------------
Private Sub PartitionByRangeInterval(byref table)
   'Range / Composite
   table.SetExtendedAttribute "TablePropertiesTablePartitioningClausesRangeOrCompositePartitioningClausePresence", true
   
   'Column list
   table.SetExtendedAttribute "TablePropertiesTablePartitioningClausesRangeOrCompositePartitioningClausePartitionByRangeColumnListColumn", gcPartitionColumNameHistTable
   
   'Define interval & Expression
   table.SetExtendedAttribute "RangePartitionIntervalPresence", true
   table.SetExtendedAttribute "RangePartitionIntervalExpression", gcPartitionRangeIntervalHistTable
   
   'Partition details
   table.setextendedattribute "TablePropertiesTablePartitioningClausesRangeOrCompositePartitioningClausePartitionByRangePartitionListPartitionDefinition", gcFirstPartitionName & " " &  gcFirstPartitionCondition
End Sub

'------------------------------------------------------------------------------
' The procedure copy all keys from sourceTable to targetTable and creates corresponding indexes
'------------------------------------------------------------------------------
Private Sub CopyKeysAndIndexesToHist(byref sourceTable, byref targetTable)
   Dim column, key, histIndex, cuttedTableName
   Dim histColumn, histKey

   For Each key In sourceTable.Keys
      'Create keys
      Set histKey = targetTable.Keys.CreateNew
      histKey.Primary = key.Primary
      If histKey.Primary Then
         histKey.Code = targetTable.Code 
         histKey.Name = histKey.Code
         'PK_%.U27:TABLE%
         histKey.ConstraintName = gcPrimaryKeyConstraintPrefix & left(targetTable.Code,27)
      Else
         histKey.Name = key.Name
         histKey.Code = key.Code
         'AK_%.U20:CUTTEDTABLE%_%.U6:AKEY%
         cuttedTableName = replace(targetTable.Code,"_","")               
         If len(cuttedTableName) > 20 Then
            cuttedTableName = left(cuttedTableName,18) & gcHistoricalTablePostfix
         End If
         histKey.ConstraintName = gcUniqueConstraintPrefix & cuttedTableName & "_" & left(histKey.Code,6)               
      End If
         
      For Each column In key.Columns
         For Each histColumn In targetTable.Columns
            If column.Name = histColumn.Name or histColumn.Name = gcPartitionColumNameHistTable Then
               histKey.Columns.Add histColumn
            End If
         Next
      Next
      'Using index <local|global> for embedded keys
      histKey.SetPhysicalOptionValue "<constraint_state>/using index/<local_partitioned_index>/local", gcTypeOfPartitionedIndexes
           
      'Create indexes based on keys
      Set histIndex = targetTable.Indexes.CreateNew
      histIndex.Code = histKey.ConstraintName
      histIndex.Name = histIndex.Code
      histIndex.LinkedObject = histKey
   Next
End Sub

'------------------------------------------------------------------------------
' The procedure helps to create foreign key referencing on a TABLE
'------------------------------------------------------------------------------
Private Sub CopyReferences(byref model, byref table, byref histTable)
   Dim reference, histReference
   
   For Each reference In table.OutReferences
      Set histReference = model.References.CreateNew
      histReference.Code = left(replace(reference.Code, reference.ChildTable.Code, histTable.Code),30)
      histReference.Name = histReference.Code
      'FK_%.U27:REFR%
      'histReference.ForeignKeyConstraintName = ...
      histReference.ParentTable = reference.ParentTable
      histReference.ChildTable = histTable
      histReference.ParentKey = reference.ParentKey
   Next
   
   'Create reference between snapshot and historical table
   Set histReference = model.References.CreateNew
   histReference.Code = histTable.Code & "S"
   histReference.Name = histReference.Code
   histReference.ParentTable = table
   histReference.ChildTable = histTable
   histReference.ParentKey = table.PrimaryKey
End Sub

'------------------------------------------------------------------------------
' The procedure shows attributes and collections of object
'------------------------------------------------------------------------------
Private Sub ViewObj(byref obj)
   Dim metaClass, attr, coll
   
   Set metaClass = obj.MetaClass   
   Output "Metaclass: " & metaclass.PublicName 
   Output "Parent: " & metaclass.Parent.PublicName
   Output "Metalibrary: " & metaclass.Library.PublicName
   
   Output "Attributes:"   
   For Each attr In metaClass.Attributes
      Output " - " + attr.PublicName
   Next
   
   Output "Collections:"   
   For each coll in metaClass.Collections
      Output " - " + coll.PublicName
   Next   
End Sub
