'******************************************************************************
'* File:     	EnrichPhysicalDataModel.vbs
'* Purpose:     The script is designed to add metadata fields and historical tables to the physical data model.
'* Version:   	3.0
'* Author:	 	Bykov D.
'******************************************************************************

Option Explicit
ValidationMode = True
InteractiveMode = im_Batch

'-----------------------------------------------------------------------------
' DECLARE GLOBAL CONSTANTS [gc]
'-----------------------------------------------------------------------------
' Global Object's Constants
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

'-----------------------------------------------------------------------------
' Main procedure
'-----------------------------------------------------------------------------
' Checks specified model and runs enrichModel procedure.
'-----------------------------------------------------------------------------
Private Sub Main(byref model)
   If model Is Nothing Then
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

'-----------------------------------------------------------------------------
' CreateUserIfNotExist procedure
'-----------------------------------------------------------------------------
' Creates a new user if he doesn't exist
'-----------------------------------------------------------------------------
Private Sub CreateUserIfNotExist(byref model)
   Dim loaderUser : loaderUser = gcMainLoaderUser
   
   'Create main loder user
   If FindUser(model, loaderUser) Is Nothing Then
      Output "Create User : " & loaderUser
      
      Set loaderUser = model.Users.CreateNew
      loaderUser.Name = gcMainLoaderUser
      loaderUser.Code = gcMainLoaderUser
   Else
      Output "User " & loaderUser & " found."
   End If
End Sub

'-----------------------------------------------------------------------------
' RecreateHistoricalTables procedure
'-----------------------------------------------------------------------------
' Shows attributes and collections of object
'-----------------------------------------------------------------------------
Private Sub RecreateHistoricalTables(byref model)
   Dim table, column, key, cuttedTableName
   Dim histTable, histColumn, histKey, histIndex
   
   'Drop All Hist tables
   For Each table in model.Tables
      If Not table.isShortcut And IsHistoricalTable(table) Then
         Output "Drop Hist Table : " + table.Code
         model.Tables.Remove table, true         
      End If
   Next
   
   'Add service columns and mandatory flag for PK_ columns
   For Each table In model.Tables
      If Not table.isShortcut Then
         checkMandatoryFlag table
	      addServiceColumns table
      End If
   Next
   
   'Create All Hist tables
   For Each table In model.Tables
      If Not table.isShortcut Then         
         Output "Create Hist Table : " & table.Name & gcHistoricalTablePostfix
      
         'Copy table
         Set histTable = model.Tables.CreateNew
         histTable.Name = table.Name & gcHistoricalTablePostfix
         histTable.Code = table.Code & gcHistoricalTablePostfix
         histTable.Comment = table.Comment
         
         'Copy columns
         For Each column In table.Columns
            createColumn histTable, column.Code, column.DataType, column.Mandatory, column.Comment
         Next
         addServiceColumns histTable
         
         'Partitioning
         partitionByRangeInterval histTable
         
         'Copy keys & indexes
         For Each key In table.Keys
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
               'AK_%.U20:CUTTEDTABLE%_%.U6:AKEY%
               cuttedTableName = replace(histTable.Code,"_","")               
               If len(cuttedTableName) > 20 Then
                  cuttedTableName = left(cuttedTableName,18) & gcHistoricalTablePostfix
               End If
               histKey.ConstraintName = gcUniqueConstraintPrefix & cuttedTableName & "_" & left(histKey.Code,6)               
            End If
         
            For Each column In key.Columns
               For Each histColumn in histTable.Columns
                  If column.Name = histColumn.Name or histColumn.Name = gcPartitionColumNameHistTable Then
                     histKey.Columns.Add histColumn
                  End If
               Next
            Next
            histKey.SetPhysicalOptionValue "<constraint_state>/using index/<local_partitioned_index>/local", gcTypeOfPartitionedIndexes
           
            'Create indexes
            Set histIndex = histTable.Indexes.CreateNew
            histIndex.Name = histKey.ConstraintName
            histIndex.Code = histKey.ConstraintName
            histIndex.LinkedObject = histKey
         Next
         
         'Copy dependencies
         copyDependencies model, table, histTable
      End If
   Next
End Sub

'-----------------------------------------------------------------------------
' CopyDependencies procedure
'-----------------------------------------------------------------------------
' helps to create foreign key referencing on TABLE
'-----------------------------------------------------------------------------
Private Sub CopyDependencies(byref model, byref table, byref histTable)
   Dim reference, histReference
   
   For Each reference in table.OutReferences
      Set histReference = model.References.CreateNew
      histReference.Name = replace(reference.Name, reference.ChildTable.Name, histTable.Name)
      histReference.Code = replace(reference.Code, reference.ChildTable.Name, histTable.Code)
      'FK_%.U27:REFR%
      'histReference.ForeignKeyConstraintName = ...
      histReference.ParentTable = reference.ParentTable
      histReference.ChildTable = histTable
      histReference.ParentKey = reference.ParentKey
   Next
End Sub

'-----------------------------------------------------------------------------
' PartitionByRangeInterval procedure
'-----------------------------------------------------------------------------
' Splits table on partitions by range
'-----------------------------------------------------------------------------
Private Sub partitionByRangeInterval(byref table)
   'Range / Composite      
   table.setextendedattribute "TablePropertiesTablePartitioningClausesRangeOrCompositePartitioningClausePresence", true
   
   'Column list
   table.setextendedattribute "TablePropertiesTablePartitioningClausesRangeOrCompositePartitioningClausePartitionByRangeColumnListColumn", gcPartitionColumNameHistTable
   
   'Define interval & Expression
   table.setextendedattribute "RangePartitionIntervalPresence", true
   table.setextendedattribute "RangePartitionIntervalExpression", gcPartitionRangeIntervalHistTable
   
   'Partition details
   table.setextendedattribute "TablePropertiesTablePartitioningClausesRangeOrCompositePartitioningClausePartitionByRangePartitionListPartitionDefinition", gcFirstPartitionName & " " &  gcFirstPartitionCondition
End Sub

'-----------------------------------------------------------------------------
' disableReferences procedure
'-----------------------------------------------------------------------------
' Disables all reference constraints on tables
'-----------------------------------------------------------------------------
Private Sub disableReferences(byref model)
   Dim ref
  
   'All references have to have disable and novalidate options
   For Each ref In model.References       
      'Output "Reference: " & ref.name & "/" & ref.foreignKeyConstraintName
      Output "Disable Reference : " & ref.Name
      
      ref.setExtendedAttribute "Validate", false
      If Not ref.getExtendedAttribute("Disable") Then
         ref.setExtendedAttribute "Disable", true         
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

'-----------------------------------------------------------------------------
' FindUser function
'-----------------------------------------------------------------------------
' Finds specified user in model. Returns user if found.
'-----------------------------------------------------------------------------
Private Function findUser(byref model, code)
   Dim user
   
   'Find user and exit if found
   For Each user in model.Users
      If user.Code = code Then
	      Set FindUser = user
         Exit Function
	   End If
   Next
   
   'If user wasn't found
   Set FindUser = Nothing
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
Private Sub AddServiceColumns(byref table)
   Dim code, dataType, comment

   code = "MD_ID_ETL_LOG"
   dataType = "INTEGER"
   comment = "Идентификатор экземпляра процесса, изменившего запись последним."
   CreateColumn table, code, dataType, true, comment

   code = "MD_FLAG_DELETED"
   dataType = "VARCHAR2(1)"
   comment = "Признак удаленной записи: Y - удалена, N - нет."
   CreateColumn table, code, dataType, true, comment

   code = "MD_CODE_SOURCE_SYSTEM"
   dataType = "VARCHAR2(10)"
   comment = "Код системы источника."
   CreateColumn table, code, dataType, true, comment
	
	If isHistoricalTable(table) Then
	   code = gcPartitionColumNameHistTable
      dataType = "DATE"
      comment = "Дата изменения записи в источнике."
      CreateColumn table, code, dataType, true, comment
	End If
End Sub

'-----------------------------------------------------------------------------
' CreateColumn procedure
'-----------------------------------------------------------------------------
' Creates a new service column and sets common attributes
'-----------------------------------------------------------------------------
Private Sub CreateColumn(byref table, code, dataType, mandatory, comment)
	Dim column
   
   If Not IsColumnExist(table, code) Then
	   Set column = table.Columns.CreateNew
		column.Name = code
		column.Code = code
      column.Mandatory = mandatory
		column.DataType = dataType
		column.Comment = comment
	End If
End Sub

'-----------------------------------------------------------------------------
' IsHistoricalTable function
'-----------------------------------------------------------------------------
' Returns true if table is historical
'-----------------------------------------------------------------------------
Private Function IsHistoricalTable(byref table)
	If right(table.Code,2) = gcHistoricalTablePostfix Then
      IsHistoricalTable = true
   Else
      IsHistoricalTable = false
   End If
End Function

'-----------------------------------------------------------------------------
' IsColumnExist function
'-----------------------------------------------------------------------------
' Find specified column amongs columns of table. Returns true if found
'-----------------------------------------------------------------------------
Private Function IsColumnExist(byref table, code)
	Dim col
	
	'Return true if column is exist
   For Each col in table.Columns
	   If col.Code = code Then
         IsColumnExist = true
         Exit Function
      End If
	Next
	
   IsColumnExist = false
End Function

'-----------------------------------------------------------------------------
' ViewObj procedure
'-----------------------------------------------------------------------------
' Shows attributes and collections of object
'-----------------------------------------------------------------------------
Private Sub ViewObj(byref obj)
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
