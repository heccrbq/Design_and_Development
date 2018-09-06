Dim tab, hTab, col, ref, jn

Output "TABLES :"
For Each tab In ActiveModel.Tables
   If Not right(tab.Code, 2) = "_H" Then
      If ActiveModel.FindChildByCode(tab.Code & "_H", PdPDM.cls_table) Is Nothing Then
         Output "   The snapshot table hasn't historical table : " & tab.Code
      End If
   End If 
Next

Output vbCrLf

Output "COLUMNS :"
For Each hTab In ActiveModel.Tables
   If right(hTab.Code, 2) = "_H" Then
      For Each col In hTab.Columns
         Set tab = ActiveModel.FindChildByCode(left(hTab.Code, len(hTab.Code)-2), PdPDM.cls_table)         
         If tab.FindChildByCode(col.Code, PdPDM.cls_column) Is Nothing and col.Code <> "MD_DATETIME_FROM" Then
            Output "   The table " & hTab.Name & " has wrong column : " & col.Code
         End If
      Next
   End If 
Next

Output vbCrLf

Output "REFERENCES :"
For Each ref In ActiveModel.References
   For Each jn In ref.Joins
      If jn.ParentTableColumn Is Nothing Then
         Output "   The reference " & ref.Name & " has empty PARENT TABLE COLUMN : " & jn.ParentTableColumn
      ElseIf jn.ChildTableColumn is Nothing Then
         Output "   The reference " & ref.Name & " has empty CHILD TABLE COLUMN : " & col.ChildTableColumn
      End If
   Next
Next

Output vbCrLf
