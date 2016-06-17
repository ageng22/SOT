VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} rptNPBS 
   Caption         =   "NOTA PENERIMAAN BARANG SEMENTARA"
   ClientHeight    =   8595
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11280
   Icon            =   "rptNPBS.dsx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   19897
   _ExtentY        =   15161
   SectionData     =   "rptNPBS.dsx":000C
End
Attribute VB_Name = "rptNPBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iRow As Integer

Private Sub ActiveReport_ReportStart()
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
Private Sub ph_npbs_Format()
 lblDate = Format$(Now)
End Sub

