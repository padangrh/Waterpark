Attribute VB_Name = "Module1"
Public fMainForm As FrmMain
Public con As New ADODB.Connection
Public rsSupplier As ADODB.Recordset
Public status
Public username As String
Public Setting_Object As Object

' **********************************************
' Posiflex usbpd.dll DLL
' **********************************************
Public Declare Function WritePD _
    Lib "usbpd.dll" _
    (ByVal data As String, ByVal Length As Long) _
As Long

Public Declare Function WritePD80 _
    Lib "usbpd.dll" Alias "WritePD" _
    (ByRef data As Any, ByVal Length As Long) _
As Long

Public Declare Function PdState _
    Lib "usbpd.dll" _
    () _
As Long

Public Declare Function OpenUSBpd _
    Lib "usbpd.dll" _
    () _
As Long

Public Declare Function CloseUSBpd _
    Lib "usbpd.dll" _
    () _
As Long

Declare Sub Sleep Lib "kernel32" _
   (ByVal dwMilliseconds As Long)
   
Public Sub connect()
    
    
'    con.ConnectionString = "Provider=MSDASQL.1;Password=yuyu;Persist Security Info=True;User ID=root;Data Source=foodcourt1"
    'con.ConnectionString = "Provider=MSDASQL.1;Password=yuyu;Persist Security Info=True;User ID=root;Data Source=iceskating"
    con.ConnectionString = "Provider=MSDASQL.1;Password=" & Setting_Object.item("DB_Pw") & ";Persist Security Info=True;User ID=" & Setting_Object.item("DB_Id") & ";Data Source=" & Setting_Object.item("DB_Name")
    con.Open
End Sub

Public Function getSupplier(kode As String) As Boolean
    If rsSupplier Is Nothing Then
        Set rsSupplier = con.Execute("select * from tbsuplier")
    'Else
        'rsSupplier.MoveFirst
    End If
    
    Dim found As Boolean
    found = False
    If Not rsSupplier.EOF Then
        rsSupplier.MoveFirst
        Do While Not rsSupplier.EOF
            If kode = rsSupplier!kdsuplier Then
                found = True
                Exit Do
            End If
            rsSupplier.MoveNext
        Loop
    End If
    getSupplier = found
End Function

Public Function priceToNum(price As String) As Long
    price = Replace(price, ".", "")
    price = Replace(price, ",", "")
    priceToNum = Val(price)
End Function

Public Function isMaster() As Boolean
    isMaster = (status = "Master")
End Function

Public Function isSPV() As Boolean
    isSPV = (status = "Supervisor")
End Function

Public Function isInTBAktif(kodeRFID As String) As Boolean
    isInTBAktif = False
    Dim rsAktif As ADODB.Recordset
    Set rsAktif = con.Execute("select * from tbaktif where rfid = '" & kodeRFID & "'")
    If Not rsAktif.EOF Then isInTBAktif = True
    Set rsAktif = Nothing
End Function

Public Sub backupAktif(kodeRFID As String, reason As String)
    Dim rsAktif As ADODB.Recordset
    Set rsAktif = con.Execute("select * from tbaktif where rfid = '" & kodeRFID & "'")
    reason = rsAktif!keterangan & " - " & reason
    con.Execute ("insert into tbnonaktif values ('" & rsAktif!rfid & "','" & Format(rsAktif!tanggal, "yyyy-mm-dd") & "','" & rsAktif!jam & "','" & rsAktif!status & "','" & reason & "','" & username & "')")
    Set rsAktif = Nothing
End Sub

Sub Main()
    Set fMainForm = New FrmMain
    fMainForm.Show
End Sub


