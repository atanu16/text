Sub UpdateBalanceSheetQuerySource()
    Dim qry As WorkbookQuery
    Dim newSource As String
    newSource = "C:\" ' Set the new source path
    
    ' Loop through all queries and find the BalanceSheet query
    For Each qry In ThisWorkbook.Queries
        If qry.Name = "BalanceSheet" Then
            ' Modify the source step in the BalanceSheet query
            qry.Formula = Replace(qry.Formula, "OldPathHere", newSource)
            Exit For
        End If
    Next qry
End Sub