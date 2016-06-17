VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} rpt_DCO_in_out_mobil 
   Caption         =   "DCO - LAPORAN KELUAR MASUK MOBIL - PENGIRIMAN"
   ClientHeight    =   8595
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   19897
   _ExtentY        =   15161
   SectionData     =   "rpt_DCO_in_out_mobil.dsx":0000
End
Attribute VB_Name = "rpt_DCO_in_out_mobil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iRow As Integer
Private Sub ActiveReport_ReportStart()
txtStartDate.Text = frmRPT_DCO_harian.txtDateIn
txtDateEnd.Text = frmRPT_DCO_harian.txtDateEnd
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
