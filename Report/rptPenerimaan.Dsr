VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} rptPenerimaan 
   BorderStyle     =   0  'None
   ClientHeight    =   10935
   ClientLeft      =   0
   ClientTop       =   -180
   ClientWidth     =   15240
   Icon            =   "rptPenerimaan.dsx":0000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   _ExtentX        =   26882
   _ExtentY        =   19288
   SectionData     =   "rptPenerimaan.dsx":000C
End
Attribute VB_Name = "rptPenerimaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iRow As Integer
Private Sub ActiveReport_ReportStart()
txtStartDate.Text = frmRPT_TIMB_harian.txtDateIn
txtDateEnd.Text = frmRPT_TIMB_harian.txtDateEnd
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

Private Sub gf_TIMB_terima_Format()
sumRANK.Caption = lblRank
End Sub

Private Sub ph_TIMB_terima_Format()
lblDate = Format$(Now)
End Sub

