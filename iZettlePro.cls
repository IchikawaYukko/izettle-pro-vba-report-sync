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

Public Sub get_access_token(ipos_username As String, ipos_password As String)
    Dim json As Object
    Dim header As Dictionary
    
    authcode = EncodeBase64("iposbackoffice:")
        
    url = "https://oauth.intelligentposapi.com/token"
    postparam = "grant_type=password&scope=ALL%3AINTERNAL" & _
            "&username=" & ipos_username & _
            "&password=" & ipos_password
    
    Set header = New Dictionary
    header.Add "Authorization", "Basic " & authcode
    header.Add "Content-Type", "application/x-www-form-urlencoded"
        
    Set json = json_parser(XmlHttpWrapper(url, "POST", postparam, header))
    AuthToken = json.access_token
End Sub

Public Sub get_report_url(start_date As Date, end_date As Date)
    Dim header As Dictionary
    Dim json As Object
    
    report_type = "full-transaction"
    param = report_type & _
            "?startDate=" & Format(start_date, "mm/dd/yyyy") & _
            "&endDate=" & Format(end_date, "mm/dd/yyyy") & _
            "&export=csv"
    url = "https://backoffice.intelligentposapi.com/v2/report/" & param
    
    Set header = New Dictionary
    header.Add "Authorization", "Bearer " & AuthToken
    
    Set json = json_parser(XmlHttpWrapper(url, "GET", "", header))
    ReportURL = json.url
End Sub

Public Sub download_report()
    csv = XmlHttpWrapper(ReportURL, "GET")
End Sub

Public Sub insert_db()
    Dim db As DAO.Database
    Set db = CurrentDb
    db.Execute ("INSERT INTO")
    db.Close: db = Nothing
End Sub

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
        cells = Split(line_, ",")
        For Each cell In cells
            cell = Replace(cell, """", "")
            linedata.Add cell
        Next cell
        data.Add linedata
    Next line_
    
    Set downloaded_data = data
End Sub

Private Function json_parser(json_text As String) As Object
    Dim obj As Object
    Set obj = CreateObject("ScriptControl")
    
    obj.Language = "JScript"
    obj.AddCode "function jsonParse(s) { return eval('(' + s + ')'); }"

    Set json_parser = obj.CodeObject.jsonParse(json_text)
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


