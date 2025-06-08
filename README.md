# MatchDepartmentToCustomerLocation

## USE CASE
Giving an Excel Workbook containing 3 Sheets named "Department", "Customer" and "Output". 
![image](https://github.com/user-attachments/assets/c9912b71-1b71-460c-a19a-0f742176eede)

See sample data below:

## DEPARTMENT
![image](https://github.com/user-attachments/assets/da584066-b611-482f-af0a-c31621658b46)

## CUSTOMER
![image](https://github.com/user-attachments/assets/48cca8c0-e5e0-4409-80b5-dbaade277c9f)

## BUSINESS REQUIREMENT
1. For each row in the Customer sheet, search and match every department name in the Department sheet.
2. If there's an exact match, then update the Output sheet with CustomerID, Customer_Name, Department Name and DepartmentID
3. Automate this process using VBA code.

## OUTPUT
![image](https://github.com/user-attachments/assets/30b838cf-c6ab-435c-a039-64f5f4f57318)

## VBA Module
```
Sub MatchDepartmentToCustomerLocation()

    Dim wsDepartment As Worksheet
    Dim wsCustomer As Worksheet
    Dim wsOutput As Worksheet
    
    Dim lastRowDepartment As Long
    Dim lastRowCustomer As Long
    Dim outputRow As Long
    
    Dim DepartmentName As String
    Dim DepartmentID As String
    Dim customerID As String
    Dim customerLocation As String
    
    Dim i As Long
    Dim j As Long
    
    ' Set references to sheets
    Set wsDepartment = ThisWorkbook.Sheets("Department")
    Set wsCustomer = ThisWorkbook.Sheets("Customer")
    Set wsOutput = ThisWorkbook.Sheets("Output")
    
    ' Find last rows
    lastRowDepartment = wsDepartment.Cells(wsDepartment.Rows.Count, 1).End(xlUp).Row
    lastRowCustomer = wsCustomer.Cells(wsCustomer.Rows.Count, 1).End(xlUp).Row
    
    ' Prepare output sheet: clear old data and write header
    wsOutput.Cells.ClearContents
    wsOutput.Range("A1:D1").Value = Array("Customer ID", "Customer Location", "Match", "Department ID")
    outputRow = 2
    
    ' Loop through customer
    For i = 2 To lastRowCustomer
        customerID = wsCustomer.Cells(i, 1).Value
        customerLocation = wsCustomer.Cells(i, 2).Value
        
        ' Loop through department
        For j = 2 To lastRowDepartment
            DepartmentName = wsDepartment.Cells(j, 1).Value
            DepartmentID = wsDepartment.Cells(j, 2).Value
            
            ' Check if department name is in customer location (case-insensitive)
            ' If InStr(1, customerLocation, "/ " & DepartmentName & " /", vbTextCompare) > 0 Then
            If InStr(1, customerLocation, DepartmentName, vbTextCompare) > 0 Then                ' Write output
                wsOutput.Cells(outputRow, 1).Value = customerID
                wsOutput.Cells(outputRow, 2).Value = customerLocation
                wsOutput.Cells(outputRow, 3).Value = DepartmentName
                wsOutput.Cells(outputRow, 4).Value = DepartmentID
                outputRow = outputRow + 1
            End If
        Next j
    Next i
    
    MsgBox "Processing complete. " & outputRow - 2 & " matches found.", vbInformation
    
End Sub

```
