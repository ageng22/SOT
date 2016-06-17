VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} rpt_DCO_kirim2 
   Caption         =   "-TIMBANGAN"
   ClientHeight    =   10950
   ClientLeft      =   165
   ClientTop       =   630
   ClientWidth     =   20250
   Icon            =   "rpt_DCO_kirim2.dsx":0000
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   35719
   _ExtentY        =   19315
   SectionData     =   "rpt_DCO_kirim2.dsx":000C
End
Attribute VB_Name = "rpt_DCO_kirim2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iRow As Integer
Private Sub ActiveReport_ReportStart()
txtStartDate.Text = frmRPT_DCO_harian2.txtDateIn
txtDateEnd.Text = frmRPT_DCO_harian2.txtDateEnd
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

Private Sub GroupFooter1_Format()
sumRANK.Caption = lblRank
End Sub
Private Sub ReportHeader_Format()
lblDate = Format$(Now)
End Sub

