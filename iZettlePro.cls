VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iZettlePro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

' iZettlePro Auto Report Downloader

Private AuthToken As String
Private ReportURL As String
Private csv As String
Private downloaded_data As Collection

Public Function get_access_token(ipos_username As String, ipos_password As String)
    Dim Json As Object
    Dim header As Dictionary
    
    authcode = EncodeBase64("iposbackoffice:")
        
    url = "https://oauth.intelligentposapi.com/token"
    postparam = "grant_type=password&scope=ALL%3AINTERNAL" & _
            "&username=" & ipos_username & _
            "&password=" & ipos_password
    
    Set header = New Dictionary
    header.Add "Authorization", "Basic " & authcode
    header.Add "Content-Type", "application/x-www-form-urlencoded"
        
    Set Json = json_parser(XmlHttpWrapper(url, "POST", postparam, header))
    AuthToken = Json("access_token")
    get_access_token = Json("access_token")
End Function

Public Function get_report_url(start_date As Date, end_date As Date)
    Dim header As Dictionary
    Dim Json As Object
    
    report_type = "full-transaction"
    param = report_type & _
            "?startDate=" & Format(start_date, "mm/dd/yyyy") & _
            "&endDate=" & Format(end_date, "mm/dd/yyyy") & _
            "&export=csv"
    url = "https://backoffice.intelligentposapi.com/v2/report/" & param
    
    Set header = New Dictionary
    header.Add "Authorization", "Bearer " & AuthToken
    
    Set Json = json_parser(XmlHttpWrapper(url, "GET", "", header))
    ReportURL = Json("url")
    get_report_url = Json("url")
End Function

Public Sub download_report()
    csv = XmlHttpWrapper(ReportURL, "GET")
End Sub

Public Sub insert_db(table_name As String)
    Dim db As DAO.Database
    Set db = CurrentDb
    
    Dim test As Double
    downloaded_data.Remove (1)  ' Delete 1st row
    For Each d In downloaded_data
        guery = build_query(d, table_name)
        
        ' Debug Output
        Forms![Downloader]![List2].AddItem Replace(Mid(guery, InStr(guery, "VALUES(") + 7), ",", "_")
        ' Listbox auto scroll bottom
        Forms![Downloader]![List2].Selected(Forms![Downloader]![List2].ListCount - 1) = True
        DoEvents
        
        db.Execute (guery)
    Next d
    db.Close: Set db = Nothing
End Sub

Private Function build_query(d As Variant, table_name As String)
        Dim trans_time As String
        Dim paidoff_date As String
        
        trans_time = convert_empty_date2null(d(12))
        paidoff_date = convert_empty_date2null(d(13))

        build_query = "INSERT INTO " & table_name & "(" & _
        "Terminal, OrderNumber, OrderId, OrderLineId," & _
        "ParentOrderLineId, DepartmentId, UserName, UserId," & _
        "TransactionType, ProductName, ProductId, TransactionTime," & _
        "PaidOffDate, CategoryName, CategoryId, GrossAmount," & _
        "TaxAmount, NetAmount, PaymentType, [Value]" & _
        ")" & _
        " VALUES(" & d(1) & "," & d(2) & _
        ",""" & d(3) & """" & ",""" & d(4) & """" & _
        ",""" & d(5) & """" & ",""" & d(6) & """" & _
        ",""" & d(7) & """" & ",""" & d(8) & """" & _
        ",""" & d(9) & """" & ",""" & d(10) & """" & _
        ",""" & d(11) & """" & _
        "," & trans_time & "," & paidoff_date & _
        ",""" & d(14) & """" & ",""" & d(15) & """" & _
        "," & Val(d(16)) & "," & Val(d(17)) & _
        "," & Val(d(18)) & _
        ",""" & d(19) & """" & _
        "," & Val(d(20)) & _
        ")"
End Function

Private Function convert_empty_date2null(date_ As Variant)
        If date_ <> "" Then
            convert_empty_date2null = "#" & date_ & "#"
        Else
            convert_empty_date2null = "null"
        End If
End Function

Private Function XmlHttpWrapper(ByVal url As String, method As String, _
        Optional ByVal postparam As String = "", _
        Optional ByVal headers As Dictionary)
        
    Dim XmlHttp As XMLHTTP60
    Set XmlHttp = New XMLHTTP60
    
    XmlHttp.Open method, url
    
    If Not (headers Is Nothing) Then
        For Each header_key In headers
            XmlHttp.setRequestHeader header_key, headers.Item(header_key)
        Next header_key
    End If
    
    If method = "GET" Then
        XmlHttp.send
    Else
        XmlHttp.send postparam
    End If
    
    Do While XmlHttp.ReadyState < 4
        DoEvents
    Loop
    
    response = XmlHttp.responseText
    Set XmlHttp = Nothing
    XmlHttpWrapper = response
End Function

Public Sub csv2array()
    Dim csvlines As Variant
    Dim data As Collection
    Set data = New Collection
    
    csvlines = Split(csv, vbLf)
    For Each line_ In csvlines
        Set linedata = New Collection
        If line_ = "" Then
            Exit For
        End If
        cells = Split(line_, ",")
        For Each cell In cells
            cell = Replace(cell, """", "")
            linedata.Add cell
        Next cell
        data.Add linedata
    Next line_
    
    Set downloaded_data = data
End Sub

Private Function json_parser(json_Text As String) As Object
    Dim obj As Object
    Set obj = JsonConverter.ParseJson(json_Text)

    Set json_parser = obj
End Function

Private Function EncodeBase64(text As String) As String
  Dim arrData() As Byte
  arrData = StrConv(text, vbFromUnicode)

  Dim objXML As MSXML2.DOMDocument60
  Dim objNode As MSXML2.IXMLDOMElement

  Set objXML = New MSXML2.DOMDocument60
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = objNode.text

  Set objNode = Nothing
  Set objXML = Nothing
End Function


