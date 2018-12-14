VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Downloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()
    Dim iztl
    Dim db As DAO.Recordset
    Dim table As String
    Dim most_newest_date, today As Date
    Dim download_from As Date
    Dim download_to As Date
    
    Set iztl = New iZettlePro
    shop_open_date = DateSerial(2018, 10, 5)
    table = "FullTransactionCafeV2"
    
    Set db = CurrentDb.OpenRecordset("SELECT Max(TransactionTime) FROM " & table & ";")
    most_newest_date = db.Fields(0)
    If IsNull(most_newest_date) Then
        most_newest_date = shop_open_date
    End If
    download_from = Round(most_newest_date)
    download_to = Date - 1  ' Yesterday
    
    If download_from = Date Then
        MsgBox "All data are up-to-date. No downloads necessary.", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    box = MsgBox("This will download " & download_from & " to " & download_to, vbOKCancel + vbInformation, "Download FullTransaction")
    If box = vbCancel Then
        Exit Sub
    End If
    
    Forms![Downloader]![Label_status].Caption = "Authorizing with iZettlePro server"
    DoEvents
    
    Forms![Downloader]![List2].AddItem "Access Token:" & iztl.get_access_token(Forms![Downloader]![UserName], Forms![Downloader]![password])
    Forms![Downloader]![List2].AddItem "Report URL:" & iztl.get_report_url(download_from, download_to)
    DoEvents
    
    Forms![Downloader]![Label_status].Caption = "Downloading from Amazon AWS....."
    DoEvents
    
    With iztl
        .download_report
        .csv2array
    End With
    
    Forms![Downloader]![Label_status].Caption = "Inserting records to " & table & " ....."
    DoEvents
    iztl.insert_db table
    
    Forms![Downloader]![Label_status].Caption = "Download complete."
End Sub
