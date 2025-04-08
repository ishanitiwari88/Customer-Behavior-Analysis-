Sub SimpleCustomerAnalysis()
    ' Create a new worksheet for the analysis
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Analysis").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = Sheets.Add
    ws.Name = "Analysis"
    
    ' Set up customer lookup
    ws.Range("A1").Value = "Customer Lookup"
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").Value = "Enter Customer ID:"
    ws.Range("B3").Value = "13593"  ' Example ID
    
    ' lookup formulas
    ws.Range("A4").Value = "Name:"
    ws.Range("B4").Formula = "=VLOOKUP(B3,ecommerce_customer_data_custom_!A:J,10,FALSE)"
    
    ws.Range("A5").Value = "Age:"
    ws.Range("B5").Formula = "=VLOOKUP(B3,ecommerce_customer_data_custom_!A:K,11,FALSE)"
    
    ws.Range("A6").Value = "Total Spent:"
    ws.Range("B6").Formula = "=SUMIFS(ecommerce_customer_data_custom_!F:F,ecommerce_customer_data_custom_!A:A,B3)"
    ws.Range("B6").NumberFormat = "$#,##0.00"
    
    ' category analysis
    ws.Range("D1").Value = "Category Analysis"
    ws.Range("D1").Font.Bold = True
    ws.Range("D3").Value = "Category"
    ws.Range("E3").Value = "Total Sales"
    ws.Range("D3:E3").Font.Bold = True
    
    ' Listing the categories
    ws.Range("D4").Value = "Electronics"
    ws.Range("D5").Value = "Home"
    ws.Range("D6").Value = "Clothing"
    ws.Range("D7").Value = "Books"
    
    ' Adding formulas
    ws.Range("E4").Formula = "=SUMIFS(ecommerce_customer_data_custom_!F:F, ecommerce_customer_data_custom_!A:A, B3, ecommerce_customer_data_custom_!C:C, D4)"
    ws.Range("E5").Formula = "=SUMIFS(ecommerce_customer_data_custom_!F:F, ecommerce_customer_data_custom_!A:A, B3, ecommerce_customer_data_custom_!C:C, D5)"
    ws.Range("E6").Formula = "=SUMIFS(ecommerce_customer_data_custom_!F:F, ecommerce_customer_data_custom_!A:A, B3, ecommerce_customer_data_custom_!C:C, D6)"
    ws.Range("E7").Formula = "=SUMIFS(ecommerce_customer_data_custom_!F:F, ecommerce_customer_data_custom_!A:A, B3, ecommerce_customer_data_custom_!C:C, D7)"
    ws.Range("E4:E7").NumberFormat = "$#,##0.00"
    
    ' Create a simple chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=350, Width:=300, Top:=50, Height:=200)
    
    With chartObj.Chart
        .SetSourceData Source:=ws.Range("D4:E7")
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "Sales by Category"
    End With
    
    ' Format worksheet
    ws.UsedRange.Columns.AutoFit
    
    ' Highlight the selected customer in main sheet
    Dim mainSheet As Worksheet
    Set mainSheet = Sheets("ecommerce_customer_data_custom_")
    
    Dim customerID As Long
    customerID = 13593  ' Example ID
    
    Dim foundCell As Range
    Dim firstAddress As String
    
    mainSheet.UsedRange.Interior.ColorIndex = xlNone
    
    With mainSheet.Range("A:A")
        Set foundCell = .Find(What:=customerID, LookIn:=xlValues)
        
        If Not foundCell Is Nothing Then
            firstAddress = foundCell.Address
            
            Do
                foundCell.EntireRow.Interior.Color = RGB(255, 255, 0)
                Set foundCell = .FindNext(foundCell)
            Loop Until foundCell.Address = firstAddress
        End If
    End With
    
    ' Activate the Analysis sheet
    ws.Activate
    MsgBox "Analysis complete!"
End Sub
