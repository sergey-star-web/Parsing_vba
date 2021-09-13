Sub ReadHTMLCode()
    Dim sURL As String
    Dim oXMLHTTP As Object
    Dim num
    Dim mycel
    Dim myrow
    Dim mycol
    Dim god
    god = 2006
    myrow = 1
    mycel = 1
    
    For i = 1 To 15 Step 1
    sURL = "https://auto.vercity.ru/statistics/sales/europe/" & god & "/russia/"        
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
    oXMLHTTP.Open "GET", sURL, False
    oXMLHTTP.send
    htmlText = oXMLHTTP.responseText
    
    While Len(htmlText) > 32000
    
    k1 = InStr(1, htmlText, "<tr data-brand=") + 16:
    c1 = Mid(htmlText, k1, 20):
    k2 = InStr(1, c1, ">") - 2:
    car = Mid(htmlText, k1, k2):
    
    myrow = myrow
    num = 1
    For j = 1 To 12 Step 1
    k1 = InStr(1, htmlText, "table_td_month table_td_collapsed") + 35:
    c1 = Mid(htmlText, k1, 20):
    k2 = InStr(1, c1, "</td>") - 1:
    jun = Mid(htmlText, k1, k2):
    Ëèñò1.Cells(myrow, 4) = jun
    Ëèñò1.Cells(myrow, 3) = num
    
    num = num + 1
    myrow = myrow + 1
    htmlText = Right(htmlText, Len(htmlText) - k1)
    Next
    
    k1 = InStr(1, htmlText, "table_td_total") + 16:
    c1 = Mid(htmlText, k1, 20):
    k2 = InStr(1, c1, "</td>") - 1:
    Price = Mid(htmlText, k1, k2):
    
    mycel = mycel
    For j = 1 To 12 Step 1
    Ëèñò1.Cells(mycel, 1) = god
    Ëèñò1.Cells(mycel, 2) = car
    Ëèñò1.Cells(mycel, 5) = Price
    
    mycel = mycel + 1
    Next
    
       htmlText = Right(htmlText, Len(htmlText) - k1 - k2)
    
   Wend
    god = god + 1
    Next
    Close #1
End Sub

Sub intxt()
Dim gid
gid = 1
Open ThisWorkbook.Path & "/rezult.txt" For Output As #1
For i = 1 To 9000 Step 1
        Print #1, Лист1.Cells(gid, 1), Лист1.Cells(gid, 2), Лист1.Cells(gid, 3), Лист1.Cells(gid, 4), Лист1.Cells(gid, 5)
        gid = gid + 1
        Next
    Close #1
    
End Sub
