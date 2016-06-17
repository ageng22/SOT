VERSION 5.00
Begin {82282820-C017-11D0-A87C-00A0C90F29FC} rptNPBT 
   Caption         =   "Nota Penerimaan Barang Masuk Timbang ( NPBMT )"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   20250
   Icon            =   "rptNPBT.dsx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35719
   _ExtentY        =   19315
   SectionData     =   "rptNPBT.dsx":000C
End
Attribute VB_Name = "rptNPBT"
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

Private Sub ph_iml_Format()
 lblDate = Format$(Now)
 End Sub
