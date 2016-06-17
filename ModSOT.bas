Attribute VB_Name = "ModSOT"
 '****************************************************************
'Create. Hasanuddin Magna
'Description:
'This module (a work in progress) currently either
'All Connection in Database Server,Create Cursor Temporary, any fuction utility
Public Const MAX_NAME_LENGTH = 50   ' name maximum length
'Global Const MINIMUM_ID_LENGTH = 3   'ID minimum length
'Global Const MAX_ID_LENGTH = 25   'ID maximum length
Public Const MINIMUM_PASSWORD_LENGTH = 6   'password minimum length
Public Const MAX_PASSWORD_LENGTH = 15   'password maximum length
'Public adminID  'for admin only
Public i, n, msg As Integer
Public msg1 As String
Public cUserName As String
Public cTraining As Boolean
Public cIML As String
Public cAccess As String
Public cTemp As String
Public cMaster As String
Public cNmForm As String
Public LoadForm As New Collection
Public fDOTPoint As Boolean

' Variabel Global untuk Connection

Public strSOT As String
Public strMMpro As String
Public cmd As ADODB.Command
Public cnSOT As ADODB.Connection
Public cnSOTLocal As ADODB.Connection
Public cnMMpro As ADODB.Connection

' Untuk SAP Connection
Public conn As Object
Public SAPConn As Object
Public SAPLogon As Object
Public PrmTGL As String
Dim training As Boolean
Dim fReadUserProfile As Object
' *******************   '
'   Main Procedure      '
' *******************   '
Sub Main()
    training = False   'Untuk Production
    'training = True     'Untuk Testing / Training
    'for developing and testing
    If training = True Then
        cTraining = False
    Else
        cTraining = False
    End If
    CreateConnection
    mnuSOT.Show
    mnuSOT.Enabled = False
    frmLogin1.Show
End Sub

' ============== '
' SAP Connection '
' ============== '
Function NewLogin(frm As Form, Optional isPrompt As Boolean = False) As Boolean
    mnuSOT.stBar.Panels(1).Text = "Connecting to SAP... "
    
    Set SAPLogon = CreateObject("SAP.LogonControl.1")
    'Set SAPConn = frm.SAPLogon.NewConnection
    Set SAPConn = frmDO.SAPLogon.NewConnection
    SAPConn.Language = "E"
    
    'Set conn = frmDO.SAPLogon.NewConnection
    'ConnectKe = "IKD"
    ConnectKe = "IKP"
    If ConnectKe = "IKT" Then
        SAPConn.Client = "501" ' client IKT
        SAPConn.System = "IKT"
        SAPConn.SystemNumber = "2"
        SAPConn.Destination = "IKT"
        SAPConn.ApplicationServer = "172.16.58.5" 'IP IKT
        SAPConn.User = ""
        SAPConn.Password = ""
        'SAPConn.ApplicationServer = "ikp3.app.co.id"
        
  ElseIf ConnectKe = "IKP" Then ' client IKP
         SAPConn.Client = "888"
         SAPConn.System = "IKP"
         SAPConn.SystemNumber = "00"
         SAPConn.Destination = "IKP"
         'SAPConn.ApplicationServer = "ikpbaru.app.co.id" ' IP IKP
         SAPConn.ApplicationServer = "pr3sikp.app.co.id" ' IP IKP
         SAPConn.User = ""
         SAPConn.Password = ""
         
'IP IKP 172.16.59.1
'        SAPConn.MessageServer = "ikpbaru.app.co.id"
'        SAPConn.GroupName = "APP_PUBLIC"
        
    ElseIf ConnectKe = "DRP" Then ' client DRP
        SAPConn.Client = "888"
        SAPConn.System = "DRP"
        SAPConn.SystemNumber = "00"
        SAPConn.Destination = "DRP"
        SAPConn.ApplicationServer = "ikp1.app.co.id" ' IP DRP
    ElseIf ConnectKe = "IKD" Then ' client IKP
        SAPConn.Client = "050"
        SAPConn.System = "IKD"
        SAPConn.SystemNumber = 0
        SAPConn.Destination = "IKD"
        SAPConn.ApplicationServer = "ikd.app.co.id"  'IP IKD
        SAPConn.User = ""
        SAPConn.Password = ""
    End If

    If isPrompt Then
       SAPConn.User = ""
       SAPConn.Password = ""
    End If
  
    'If SAPConn.Logon(0, True) <> True Then
    If SAPConn.Logon(0, Not isPrompt) <> True Then
        If SAPConn.IsConnected = 8 Then
            msg = MsgBox("Can not logon to SAP", vbCritical, "Logon Error")
        ElseIf SAPConn.IsConnected = 2 Then
            mnuSOT.stBar.Panels(1).Text = "Canceling Logon to SAP"
    
        End If
        NewLogin = False
        Exit Function
    End If
    
    mnuSOT.stBar.Panels(1).Text = "Connected to " + SAPConn.System + "(" + SAPConn.Client + ")"
   ' frmDO.SAPFunc.Connection = SAPConn
   ' frmDO.SAPTrans.Connection = SAPConn
   'remark by ilham 07 dec 2004 18:00
    'fDOTPoint = ReadUserProfile(frm)
    NewLogin = True
End Function
Function ReadUserProfile(frm As Form) As Boolean
Dim wfunct1 As Object
Dim wfunct2 As Object
    cUser = SAPConn.User
    Set wfunct1 = CreateObject("SAP.Functions")
    Set wfunct1.Connection = SAPConn
    Set fReadUserProfile = wfunct1.Add("Y_READ_USR01")
    
    If fReadUserProfile Is Nothing Then
       Set fReadUserProfile = frmDO.SAPFunc.Add("Y_READ_USR01")
    End If


    'If fReadUserProfile.Call <> True Then
    '   MsgBox "Fail on RFC Call 'Y_READ_USR01' " & Chr(13) & wfunct2.Exception
    'End If
    
    fReadUserProfile.Exports.Item("USER").Value = cUser
    
    mnuSOT.stBar.Panels(1).Text = "Read " + cUser + "'s User Profile"
     If fReadUserProfile.Call Then
    ' If wfunct2.Call Then
        'Set XTABLE = wfunct2.Tables.Item("wtable")
        Set XTABLE = fReadUserProfile.Tables.Item("wtable")
        If XTABLE.Value(1, "DCPFM") = "X" Then                                               ' X berarti gunakan point
            ReadUserProfile = True
        Else
            ReadUserProfile = False
        End If
    Else
        If fReadUserProfile.Exception = "NOT_FOUND" Then
            ' Update SOT with correct Status
            msg1 = "User " + cUser + " doesn't exist in SAP"
            msg1 = MsgBox(msg1, vbInformation, "SOT Message Info")
        Else
            msg1 = fReadUserProfile.Exception + " when read user profile"
            msg1 = MsgBox(msg1, vbInformation, "SOT Message Info")
        End If
    End If
End Function
' ============================ '
' Functionality for PostgreeSQL '
' ============================ '
Sub CreateConnection()
    ' Inisialisasi Object ADO
   
    Set cmd = New ADODB.Command
    Set cnSOT = New ADODB.Connection
    Set cnSOTLocal = New ADODB.Connection
   
    ''strSOT = "DRIVER={PostgreSQL Unicode};SERVER=localhost;port=5432;DATABASE=SOT;UID=postgres;PWD=lupapasswd"
    strSOT = "DRIVER={PostgreSQL Unicode};SERVER=172.16.123.17;port=5432;DATABASE=SOT;UID=postgres;PWD="

   
        cnSOT.CursorLocation = adUseServer
        cnSOT.Mode = adModeReadWrite
        cnSOT.Open strSOT
    
    ' Bikin connection local untuk SOT
        cnSOTLocal.CursorLocation = adUseClient
        cnSOTLocal.Mode = adModeReadWrite
        cnSOTLocal.Open strSOT
        cmd.Prepared = True
End Sub
Sub DropTable(cn As ADODB.Connection, ByVal strtable As String)
    'Akan ngedrop table jika Table Exist
    On Error Resume Next
    cn.Execute ("drop table " + Trim(strtable))
    'cn.Execute ("delete from " + Trim(strtable))
End Sub
Sub DeleteIsiTable(cn As ADODB.Connection, ByVal strtable As String)
    'Akan ngedrop table jika Table Exist
    On Error Resume Next
    cn.Execute ("delete from TEMPG2")
End Sub
Sub zap(cn As ADODB.Connection, ByVal strtable As String)
    On Error Resume Next
    'cn.Execute ("delete from " + strtable)
    cn.Execute ("delete from TEMPG2")
End Sub
Sub CreateCurCont(cn As ADODB.Connection, ByVal strtable As String)
    DropTable cn, strtable
    cnSOT.Execute ("create table " + Trim(strtable) + " (" + _
    "delino char(10) NULL," + _
    "CONTNO char(25) NULL," + _
    "NewCont char(25) NULL," + _
    "VENUM char(10) NULL)")
End Sub
Sub CreateCurGridH(cn As ADODB.Connection, ByVal strtable As String)
    DropTable cn, strtable
    cn.Execute ("create table " + Trim(strtable) + " ( " + _
    "NoItems numeric(10,4) null," + _
    "NoPO char(20) null ," + _
    "NoSJ char(30) NULL , " + _
    "satuan char (15) NULL , " + _
    "nmbarang char (100) NULL, " + _
    "QTYSJ char(20) NULL ," + _
    "QTYactual char(20) NULL)")
End Sub
Sub CreateCurGridH1(cn As ADODB.Connection, ByVal strtable As String)
DropTable cn, strtable
    cn.Execute ("create table " + Trim(strtable) + " ( " + _
    "NoItems numeric(10,4) null," + _
    "satuan char (15) NULL , " + _
    "nmbarang char (50) NULL, " + _
    "keterangan char (50) NULL," + _
    "QTY char(20) NULL)")
End Sub
Sub CreateCurGrid(cn As ADODB.Connection, ByVal strtable As String)
    DropTable cn, strtable
    cn.Execute ("create table " + Trim(strtable) + " (" + _
    "delino char(10) NULL," + _
    "plancont char(25) NULL," + _
    "matnr char(8) NULL," + _
    "charg char(10) NULL," + _
    "qtydn decimal(10,4) NULL," + _
    "uom char(5) NULL," + _
    "qtydnkg decimal(10,4) NULL," + _
    "no_container char(25) NULL," + _
    "VENUM char(10) NULL)")
End Sub
Sub CreateCurLocal(cn As ADODB.Connection, ByVal strtable As String)
     DropTable cn, strtable
     cn.Execute ("create table " + Trim(strtable) + " (" + _
                    "VBELN char(10) NULL," + "ERNAM char(12) NULL," + _
                    "ERDAT char(10) NULL," + "BZIRK char(6) NULL," + _
                    "VSTEL char(4) NULL," + "VSTELT char(20) NULL," + _
                    "VKORG char(8) NULL," + "WADAT char(10) NULL," + _
                    "LDDAT char(10) NULL," + _
                    "TDDAT char(10) NULL," + "LFDAT char(10) NULL," + _
                    "ROUTE char(6) NULL," + "KUNNR char(10) NULL," + _
                    "KUNAG char(10) NULL," + "KDGRP char(2) NULL," + _
                    "BTGEW char(15) NULL," + "NTGEW char(15) NULL," + _
                    "GEWEI char(3) NULL," + "VOLUM char(15) NULL," + _
                    "VOLEH char(3) NULL," + "ANZPK char(5) NULL," + _
                    "AEDAT char(15) NULL," + "TRATY char(4) NULL," + _
                    "VTEXT char(20) NULL," + "TRAID char(20) NULL," + _
                    "KUNNRT char(50) NULL," + "KUNAGT char(50) NULL)")
End Sub

Sub CreateCurExport(cn As ADODB.Connection, ByVal strtable As String)
    DropTable cn, strtable
    cn.Execute ("create table " + Trim(strtable) + " (" + _
        "VBELN char(10) NULL," + "ERNAM char(12) NULL," + _
        "ERDAT char(10) NULL," + "BZIRK char(6) NULL," + _
        "VSTEL char(4) NULL," + "VSTELT char(20) NULL," + _
        "VKORG char(8) NULL," + "WADAT char(10) NULL," + _
        "LDDAT char(10) NULL," + _
        "TDDAT char(10) NULL," + "LFDAT char(10) NULL," + _
        "ROUTE char(6) NULL," + "KUNNR char(10) NULL," + _
        "KUNAG char(10) NULL," + "KDGRP char(2) NULL," + _
        "BTGEW char(15) NULL," + "NTGEW char(15) NULL," + _
        "GEWEI char(3) NULL," + "VOLUM char(15) NULL," + _
        "VOLEH char(3) NULL," + "ANZPK char(5) NULL," + _
        "AEDAT char(15) NULL," + "TRATY char(4) NULL," + _
        "VTEXT char(20) NULL," + "TRAID char(20) NULL," + _
        "KUNNRT char(30) NULL," + "KUNAGT char(30) NULL)")
End Sub
' ============================ '
' Functionality of Form Access '
' ============================ '
Function CheckRequired(frm As Form) As Boolean
    Dim Ctrl As Control
    For Each Ctrl In frm.Controls
        If ((TypeOf Ctrl Is TextBox) Or (TypeOf Ctrl Is SSOleDBCombo)) And Left(Ctrl.Tag, 1) = "?" Then
            ''frmPenerimaan.txtKodeSupplier.Text = "object_ini_dioffin_dulu"
            ''frmPenerimaan.txtSatuan.Text = "object_ini_dioffin_dulu"
            If Trim(Ctrl.Text) = "" Or Trim(Ctrl.Text) = "?" Then
                msg = MsgBox("Required entry not made." + vbCr + "Please input " + _
                Mid(Ctrl.Tag, 3), vbExclamation, "SOT Message Error")
                Ctrl.SetFocus
                CheckRequired = False
                Exit Function
            End If
        End If
    Next Ctrl
    CheckRequired = True
End Function
Function CheckRequiredOpt(frm As Form) As Boolean
    Dim Ctrl As Control
    For Each Ctrl In frm.Controls
        If (TypeOf Ctrl Is OptionButton) Then
            If Trim(Ctrl.Value) = False Then
                msg = MsgBox("Required entry no made." + vbCr + "Please tick for option delivery. ")
                Ctrl.SetFocus
                CheckRequiredOpt = False
                Exit Function
            End If
        End If
    Next Ctrl
    CheckRequiredOpt = True
End Function
Sub Sorot(obj As Object)
    obj.SelStart = 0
    obj.SelLength = Len(obj)
End Sub
Function f(ByVal str As String) As String
    f = "'" + str + "'"
End Function
Function NS(ByVal obj As Variant) As String
    NS = IIf(IsNull(obj), "", obj)
End Function
Function NM(ByVal obj As Variant) As String
    NM = IIf(IsNull(obj), "-", obj)
End Function
Function NN(ByVal obj As Variant) As Variant
    NN = IIf(IsNull(obj), 0, obj)
End Function
Function v(ByVal str As String) As String
    v = "'" + str + "',"
End Function
Function s(ByVal str As String) As String
    s = "'" + str + "'"
End Function

Function g(ByVal str As String) As String
    g = "'" + str + "',"
End Function

Sub grdLeaveCell(grdMe As MSFlexGrid, cntlme As Control)
'
grdMe.Text = cntlme.Text
cntlme.Text = ""
cntlme.Visible = False
'
End Sub

Public Sub GrdEnterCell(grdMe As MSFlexGrid, cntlme As Control)
'
Dim intCntl As Integer
'
With grdMe
    cntlme.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth, .CellHeight
    cntlme.Visible = True
    cntlme.SetFocus
End With
'
End Sub

Public Function fnomorkendaraan(fstr1 As String) As String
'
fstr3 = ""
fnomorkendaraan = ""
For i = 1 To 11
    fstr3 = Mid(fstr1, i, 1)
    If (fstr3 <> " ") Then
        If (fstr3 <> "_") Then
            fnomorkendaraan = fnomorkendaraan + fstr3
        End If
    End If
Next i
'
End Function

Public Sub goto_Cb(ByVal cb As ComboBox, parValue As String)
    If Trim(parValue) = "" Or IsNull(parValue) Then
        cb.ListIndex = -1
        Exit Sub
    End If
    For i = 0 To (cb.ListCount - 1)
        cb.ListIndex = i
        If Trim(cb.Text) = Trim(parValue) Then
            Exit Sub
        End If
        'If getCb_codeValue(cb) = Trim(parValue) Then
        '    Exit Sub
        'End If
    Next
End Sub


' Copy from coding Pindo By yady on 2009-05-19

'Added by Sahrul on 29-07-2008
Function GetBarang(ByVal stCode As String) As String
 Dim rs As New ADODB.Recordset
 Dim stSQL As String
 On Error GoTo ErrSub
 '
 stSQL = "select keterangan as nama from kodebarang (nolock) " & _
        "where kode='" & Trim(stCode) & "'"
 rs.Open stSQL, cnSOT, adOpenForwardOnly, adLockReadOnly
 If Not rs.EOF Then
   GetBarang = Trim(rs!nama)
   rs.Close
 End If
 Set rs = Nothing
 Exit Function
 
ErrSub:
  MsgBox Err.Description, vbCritical, "Error..."
End Function
'Added by Sahrul on 29-07-2008
Function GetExpedisi(ByVal stCode As String) As String
 Dim rs As New ADODB.Recordset
 Dim stSQL As String
 On Error GoTo ErrSub
 '
 stSQL = "select keterangan as nama from kodeexpedisi (nolock) " & _
        "where kode='" & Trim(stCode) & "'"
 rs.Open stSQL, cnSOT, adOpenForwardOnly, adLockReadOnly
 If Not rs.EOF Then
   GetExpedisi = Trim(rs!nama)
   rs.Close
 End If
 Set rs = Nothing
 Exit Function
 
ErrSub:
  MsgBox Err.Description, vbCritical, "Error..."
End Function

'Added by Sahrul on 19-08-2008
Function GetBeratMobil(ByVal stNoPol As String, ByVal stTypeTruck As String) As Double
 Dim rs As New ADODB.Recordset
 Dim stSQL As String
 
 On Error GoTo ErrSub
 '
 stSQL = "select (BrtKosong1 + BrtKosong2) as beratkosong from tbrtmobil where nomobil = '" & Trim(stNoPol) & "' and JnKend = '" & Trim(stTypeTruck) & "' "
 rs.Open stSQL, cnSOT, adOpenForwardOnly, adLockReadOnly
 If Not rs.EOF Then
   GetBeratMobil = rs!GrossWeight
   rs.Close
 End If
 Set rs = Nothing
 Exit Function
 
ErrSub:
  MsgBox Err.Description, vbCritical, "Error..."
End Function


