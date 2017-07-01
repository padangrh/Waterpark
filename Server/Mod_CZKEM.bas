Attribute VB_Name = "Mod_CZKEM"
Public Sub pushC1(rfid As String)
    Dim rsReader As ADODB.Recordset
    Dim bcrtUser As Boolean
    'confirmAllC1
    'checkInactiveC1
    Set rsReader = con.Execute("select * from tbreader where rfid = '" & rfid & "'")
    If StatusC1_1 Then
        frmMain.CZKEM1.CardNumber(0) = CLng(rfid)
        bcrtUser = frmMain.CZKEM1.SetUserInfo(1, rsReader!id, "Tamu", "0751", 0, 1)
    End If
    If StatusC1_2 Then
        frmMain.CZKEM2.CardNumber(0) = CLng(rfid)
        bcrtUser = frmMain.CZKEM2.SetUserInfo(1, rsReader!id, "Tamu", "0751", 0, 1)
    End If
    If StatusC1_3 Then
        frmMain.CZKEM3.CardNumber(0) = CLng(rfid)
        bcrtUser = frmMain.CZKEM3.SetUserInfo(1, rsReader!id, "Tamu", "0751", 0, 1)
    End If
End Sub

Public Sub deleteC1(ReaderID As String)
    Dim rsReader As ADODB.Recordset
    Dim bdltUser As Boolean
    'confirmAllC1
    'checkInactiveC1
    Set rsReader = con.Execute("select * from tbreader where rfid = '" & ReaderID & "'")
    If Not rsReader.EOF Then
        If StatusC1_1 Then
            bdltUser = frmMain.CZKEM1.DeleteEnrollData(1, rsReader!id, 1, 12)
        End If
        If StatusC1_2 Then
            bdltUser = frmMain.CZKEM2.DeleteEnrollData(1, rsReader!id, 1, 12)
        End If
        If StatusC1_3 Then
            bdltUser = frmMain.CZKEM3.DeleteEnrollData(1, rsReader!id, 1, 12)
        End If
    End If
End Sub


Public Function confirmC1(ip As String) As Boolean
    Dim yy As Boolean
    yy = False
'    FrmMain.Winsock1.LocalPort = 0
    If frmMain.Winsock1.State = sckClosed Then
        frmMain.Winsock1.Protocol = sckTCPProtocol
        frmMain.Winsock1.connect ip, 80
        frmMain.Timer1.Enabled = True
        Do
            DoEvents
            If frmMain.Winsock1.State = 7 Then
                frmMain.Timer1.Enabled = False
                yy = True
            End If
        Loop While frmMain.Timer1.Enabled
        frmMain.Winsock1.Close
    ElseIf frmMain.Winsock1.State = 7 Then
        yy = True
    End If
    confirmC1 = yy
End Function

Public Sub confirmAllC1()
    Dim xx As Boolean
    xx = False
    xx = confirmC1(Setting_Object("C1_1"))
    If xx = False Then
        frmMain.cmdC1_1.BackColor = &HFF&
        StatusC1_1 = xx
    End If
    xx = confirmC1(Setting_Object("C1_2"))
    If xx = False Then
        frmMain.cmdC1_2.BackColor = &HFF&
        StatusC1_2 = xx
    End If
    xx = confirmC1(Setting_Object("C1_3"))
    If xx = False Then
        frmMain.cmdC1_3.BackColor = &HFF&
        StatusC1_3 = xx
    End If
    DoEvents
End Sub

Public Sub checkInactiveC1()
    Dim xx As Boolean
    xx = False
    
    If StatusC1_1 = False Then
        StatusC1_1 = confirmC1(Setting_Object("C1_1"))
        If StatusC1_1 Then
            Dim C1_1Con As Boolean
            CZKEM1.BASE64 = 1
            C1_1Con = False
            C1_1Con = CZKEM1.Connect_Net(Setting_Object("C1_1"), 4370)
            If C1_1Con Then CZKEM1.Beep 150
            refillC1 1
            cmdC1_1.BackColor = &HFF00&
        End If
    End If
    
    If StatusC1_2 = False Then
        StatusC1_2 = confirmC1(Setting_Object("C1_2"))
        If StatusC1_2 Then
            Dim C1_2Con As Boolean
            CZKEM2.BASE64 = 1
            C1_2Con = False
            C1_2Con = CZKEM2.Connect_Net(Setting_Object("C1_2"), 4370)
            If C1_2Con Then CZKEM2.Beep 150
            refillC1 2
            cmdC1_2.BackColor = &HFF00&
        End If
    End If
    
    If StatusC1_3 = False Then
        StatusC1_3 = confirmC1(Setting_Object("C1_3"))
        If StatusC1_3 Then
            Dim C1_3Con As Boolean
            CZKEM3.BASE64 = 1
            C1_3Con = False
            C1_3Con = CZKEM3.Connect_Net(Setting_Object("C1_3"), 4370)
            If C1_3Con Then CZKEM3.Beep 150
            refillC1 3
            cmdC1_3.BackColor = &HFF00&
        End If
    End If
    DoEvents
End Sub

Public Sub refillC1(C1_ID As Integer)
    Dim i As Long
    Dim tempCZKEM As CZKEM
    Dim rsReader As ADODB.Recordset
    Dim dwEnrollNmber As Long
    Dim name As String
    Dim pwd As String
    Dim privilege As Long
    Dim sEnabled As Boolean
    
    If C1_ID = 1 Then
        Set tempCZKEM = frmMain.CZKEM1
    ElseIf C1_ID = 2 Then
        Set tempCZKEM = frmMain.CZKEM2
    ElseIf C1_ID = 3 Then
        Set tempCZKEM = frmMain.CZKEM3
    End If
    
    ' Sambung besok
    
    Set rsReader = con.Execute("select * from tbreader")

    If tempCZKEM.ReadAllUserID(1) Then
        dwEnrollNmber = 0
        Do While tempCZKEM.GetAllUserInfo(CLng(1), dwEnrollNmber, name, pwd, privilege, sEnabled)
            If dwEnrollNmber <> 65534 Then
                bdltUser = tempCZKEM.DeleteEnrollData(1, dwEnrollNmber, 1, 12)
            End If
        Loop
    End If
        
    Do While Not rsReader.EOF
        tempCZKEM.CardNumber(0) = CLng(rsReader!rfid)
        bcrtUser = tempCZKEM.SetUserInfo(1, rsReader!id, "Tamu", "0751", 0, 1)
        rsReader.MoveNext
    Loop
    
    Set tempCZKEM = Nothing
    Set rsReader = Nothing
End Sub

