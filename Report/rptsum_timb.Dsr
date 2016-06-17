VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} rptsum_timb 
   Caption         =   "LAPORAN BERAT TIMBANG - PENERIMAAN"
   ClientHeight    =   10950
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11280
   Icon            =   "rptsum_timb.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   19897
   _ExtentY        =   19315
   SectionData     =   "rptsum_timb.dsx":000C
End
Attribute VB_Name = "rptsum_timb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iRow As Integer

Private Sub ActiveReport_ReportStart()
txtStartDate.Text = frmPrintBerat_Timb.txtDateIn
txtDateEnd.Text = frmPrintBerat_Timb.txtDateEnd
    If frmPrintBerat_Timb.Option1 = True Then
        lblModul.Caption = "Pengiriman Barang"
        lblgudang.Caption = "GUDANG MUAT"
        lbliso.Caption = "FO/8/08/112   REV.0"
    Else
        lblModul.Caption = "Penerimaan Barang"
        lblgudang.Caption = "GUDANG BONGKAR"
        lbliso.Caption = "FO/8/08/113   REV.0"
    End If
    
txtusername.Text = cUserName
iRow = 0
End Sub

Private Sub Detail_Format()

    If iRow Mod 2 = 0 Then
        Detail.BackColor = vbWhite
    Else
        Detail.BackColor = &HC0C0FF
    End If
    iRow = iRow + 1
    lblRank.Caption = str(iRow)
End Sub

Private Sub PageHeader_Format()
 lblDate = Format$(Now)
End Sub
