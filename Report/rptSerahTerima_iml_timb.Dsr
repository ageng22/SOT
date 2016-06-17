VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} rptSerahTerima_iml_timb 
   Caption         =   "LAPORAN SERAH TERIMA SPBMT"
   ClientHeight    =   8595
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11280
   Icon            =   "rptSerahTerima_iml_timb.dsx":0000
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   19897
   _ExtentY        =   15161
   SectionData     =   "rptSerahTerima_iml_timb.dsx":000C
End
Attribute VB_Name = "rptSerahTerima_iml_timb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iRow As Integer
Dim CStr1 As String

Private Sub ActiveReport_ReportStart()
txtStartDate.Text = frmSerahTerima_NPBMT.txtDateIn
txtDateEnd.Text = frmSerahTerima_NPBMT.txtDateEnd
lblLokasiGudang.Caption = frmSerahTerima_NPBMT.txtGudangtujuan
txtusername.Text = cUserName
iRow = 0
CStr1 = ""
End Sub

Private Sub Detail_Format()

If iRow Mod 2 = 0 Then
    Detail.BackColor = vbWhite
Else
    Detail.BackColor = &HC0C0FF
End If
    
If CStr1 <> txtGDMUAT.Text Then
    iRow = iRow + 1
    lblRank.Caption = str(iRow)
    CStr1 = txtGDMUAT.Text
Else
    lblRank.Caption = ""
    txtGDMUAT.Text = ""
    Field11.Text = ""
    txtTotKendaraan.Text = ""
End If

End Sub

Private Sub PageHeader_Format()
 lblDate = Format$(Now)
End Sub
