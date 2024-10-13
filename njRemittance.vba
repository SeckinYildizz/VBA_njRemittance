Sub NJ_Remittance()
' Format the remittance report that is extracted as a text and pasted into a sheet to a more managable format.
    
    Dim searchString, aRow, aData, bData As String
    Dim cell As Range
    Dim i, num As Integer
    
    'Create a new sheet and type column header accordingly
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "Formatted Remittance"
    With Sheets("Formatted Remittance")
        .Range("a1:m1").Value = Array("Receipent ID", "Receipent Name", "Date of Service From", _
        "Date of Service To", "Units", "Procedure", "Description", "Amount Billed", _
        "Amount Allowed", "Total Deduct", "Amount Paid", "Control Number", "Error Codes")
    End With
    
    'Create an error handler
On Error GoTo Error_Handler
    
    'Set the essential variables
    searchString = "APPROVED ORIGINAL CLAIMS"
    i = 1
    num = 1
    
    'Some of the IDs start with 0 so that they have to be converted to text
    Range("a:a").NumberFormat = "@"
    
    ' Check all cells in the used range and find the cells that have value equal to the string
    For Each cell In Sheets(1).UsedRange
        
        ' Check if the cell contains the search string
        If Trim(cell.Value) = searchString Then
            Do While Not Len(Trim(cell.Offset(i, 0))) = 0 Or InStr(cell.Offset(i, 0).Value, searchString) > 0
                ' If the string is found, get its value and extract the ptp id
                aRow = Trim(cell.Offset(i, 0).Value)
                aData = Left(aRow, InStr(aRow, " ") - 1)
                Sheets("Formatted Remittance").Range("a" & num + 1).Value = aData
                                
                'Remove ptp id from the string
                aRow = Right(aRow, Len(aRow) - Len(aData))
                aRow = Trim(aRow)
                
                'Extract the ptp name and remove it from the main string
                aData = Left(aRow, InStr(aRow, " ") - 1)
                aRow = Right(aRow, Len(aRow) - Len(aData))
                aRow = Trim(aRow)
                bData = Left(aRow, InStr(aRow, " ") - 1)
                aRow = Right(aRow, Len(aRow) - Len(bData))
                aRow = Trim(aRow)
                Sheets("Formatted Remittance").Range("b" & num + 1).Value = aData & ", " & bData
                
                'Extract the service date start and remove it from the main string
                aData = Left(aRow, InStr(aRow, " ") - 1)
                Sheets("Formatted Remittance").Range("c" & num + 1).Value = aData
                aRow = Right(aRow, Len(aRow) - Len(aData))
                aRow = Trim(aRow)
                
                'Extract the service date end and remove it from the main string
                aData = Left(aRow, InStr(aRow, " ") - 1)
                Sheets("Formatted Remittance").Range("d" & num + 1).Value = aData
                aRow = Right(aRow, Len(aRow) - Len(aData))
                aRow = Trim(aRow)
                
                'Extract the units and remove it from the main string
                aData = Left(aRow, InStr(aRow, " ") - 1)
                Sheets("Formatted Remittance").Range("e" & num + 1).Value = aData
                aRow = Right(aRow, Len(aRow) - Len(aData))
                aRow = Trim(aRow)
                
                'Extract the control number and remove it from the main string
                aData = Right(aRow, Len(aRow) - InStrRev(aRow, " ") + 1)
                Sheets("Formatted Remittance").Range("l" & num + 1).Value = aData
                aRow = Left(aRow, Len(aRow) - Len(aData))
                aRow = Trim(aRow)
                
                'Extract the amount paid and remove it from the main string
                aData = Right(aRow, Len(aRow) - InStrRev(aRow, " ") + 1)
                Sheets("Formatted Remittance").Range("k" & num + 1).Value = aData
                aRow = Left(aRow, Len(aRow) - Len(aData))
                aRow = Trim(aRow)
                
                'Extract the total deduct and remove it from the main string
                aData = Right(aRow, Len(aRow) - InStrRev(aRow, " ") + 1)
                Sheets("Formatted Remittance").Range("j" & num + 1).Value = aData
                aRow = Left(aRow, Len(aRow) - Len(aData))
                aRow = Trim(aRow)
                
                'Extract the amount allowed and remove it from the main string
                aData = Right(aRow, Len(aRow) - InStrRev(aRow, " ") + 1)
                Sheets("Formatted Remittance").Range("i" & num + 1).Value = aData
                aRow = Left(aRow, Len(aRow) - Len(aData))
                aRow = Trim(aRow)
                
                'Extract the amount billed and remove it from the main string
                aData = Right(aRow, Len(aRow) - InStrRev(aRow, " ") + 1)
                Sheets("Formatted Remittance").Range("h" & num + 1).Value = aData
                aRow = Left(aRow, Len(aRow) - Len(aData))
                aRow = Trim(aRow)
                
                'Extract the description and remove it from the main string
                aData = Right(aRow, Len(aRow) - InStrRev(aRow, " ") + 1)
                Sheets("Formatted Remittance").Range("g" & num + 1).Value = aData
                aRow = Left(aRow, Len(aRow) - Len(aData))
                aRow = Trim(aRow)
                
                'Extract the procedure
                Sheets("Formatted Remittance").Range("f" & num + 1).Value = aRow
                
                'Get the error code
                aData = cell.Offset(i + 1, 0).Value
                Sheets("Formatted Remittance").Range("m" & num + 1).Value = aData
                i = i + 2
                num = num + 1
            Loop
            
            'Reset i for the next iteration
            i = 1
            
        End If
        
    Next cell
        
Error_Handler:
    'Final touch
    With Sheets("Formatted Remittance")
        .Range("l:l").NumberFormat = "0"
        .Columns("a:m").AutoFit
    End With
     
End Sub


