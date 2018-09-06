Dim tab, col, ref, jn

Output "TABLES :"
For Each tab In ActiveModel.Tables
   If tab.Code <> tab.Name Then
      Output "   The CODE is not equal to the NAME : " & tab.Name
   End If
Next

Output vbCrLf

Output "COLUMNS :"
For Each tab In ActiveModel.Tables
   For Each col In tab.Columns
      If col.Name = "MD_ID_ETL_LOG" or col.Name = "MD_FLAG_DELETED" or col.Name = "MD_CODE_SOURCE_SYSTEM" or col.Name = "MD_DATETIME_FROM" Then
         Output "   Table has already have service columns : " & tab.Name
         Exit For
      End If
   Next
Next

Output vbCrLf

Output "REFERENCES :"
For Each ref In ActiveModel.References
   If ref.Name <> ref.Code Then
      Output "   The CODE is not equal to the NAME : " & ref.Name
   Else
      If Instr(ref.Code,ref.ChildTable.Code) = 0 Then
         Output "   The reference code does not include the name of the child table : "  & ref.Name
      End If
   End If
   
   For Each jn in ref.Joins
      If jn.ChildTableColumn Is Nothing or jn.ParentTableColumn Is Nothing Then
         Output "   PARENT or CHILD TABLE COLUMN is empty : " & ref.Name
      End If
   Next
Next

Output vbCrLf
