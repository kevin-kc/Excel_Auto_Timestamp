# Excel_Auto_Timestamp

** see the document "How the auto timestamp works (picture explaination)" for a visual explaination. **

Summary: Timestamp function which will place a timestamp on a cell whenever another cell in that row is changed. The changes are based on a hash created by the string values of each cell in each row respectively. It is done this way instead of just placing a timestamp whenever a cell is modified because this will allow the user to undo a change. Consider changing a cell's value from 0 to 0, the value in the hash has not changed so the timestamp will not be updated, but excel detects this as a modification so without the hash the timestamp would update. Also consider trying to undo a change like this so excel gives you back your old timestamp value, well because the timestamp column was updated the undo with undo the timestamp update, but this changes a cell in the row so the timestamp is once again updated, the regression is solved with the string hash of the row's values.


This function places a timestamp whenever the string hash that is created for each row changes.

    Function activeTimeStamp(r As Variant, lookupRange As Range) As String

        Dim s As String 'hash string
        Dim c As Long 'column counter
        s = ""

        For c = 2 To lookupRange.columns.Count
                s = s & CStr(Excel.Cells(r, c).value)
        Next

        If s = Hash.item(r) Then
            Exit Function
        Else
            activeTimeStamp = Now()
            Hash.item(r) = s
        End If
    
    End Function
    
This function creates a string hash for each row in the sheet. can be modified to be compacted with SHA if needed.

    Option Explicit
    Dim Hash As New Dictionary
    
    Function createHash(SheetNum As Integer, TableName As String)
    
        Dim sheet As Worksheet 'worksheet
        Set sheet = Excel.Worksheets(SheetNum)

        Dim r As Long, c As Long 'row,col
        Dim s As String 'cell value
        Dim targetTable As ListObject 'table object
        Dim targetRange As Range 'table data
        Set targetTable = sheet.ListObjects(TableName)
        Set targetRange = targetTable.Range

        r = targetRange.rows.Count
        For r = 2 To targetRange.rows.Count
            For c = 2 To targetRange.columns.Count
                    s = s & CStr(sheet.Cells(r, c).value)
            Next
            Call Hash.Add(r, s)
        Next

    End Function
