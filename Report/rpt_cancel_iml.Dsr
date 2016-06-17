VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} rpt_cancel_iml 
   Caption         =   "-TIMBANGAN"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "rpt_cancel_iml.dsx":0000
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   20955
   _ExtentY        =   15161
   SectionData     =   "rpt_cancel_iml.dsx":000C
End
Attribute VB_Name = "rpt_cancel_iml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iRow As Integer
Private Sub ActiveReport_ReportStart()
txtStartDate.Text = frm_cancel_iml.txtDateIn
txtDateEnd.Text = frm_cancel_iml.txtDateEnd
txtusername.Text = cUserName
txtJudul.Text = frm_cancel_iml.txtJudul
Labelno.Caption = frm_cancel_iml.txtno
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
