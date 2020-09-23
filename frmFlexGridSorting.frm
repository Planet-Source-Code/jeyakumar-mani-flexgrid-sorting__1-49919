VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmFlexGridSorting 
   Caption         =   "Flexgrid Sorting"
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2775
      Left            =   -15
      TabIndex        =   0
      Top             =   135
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
End
Attribute VB_Name = "frmFlexGridSorting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
MSHFlexGrid1.Rows = 7
MSHFlexGrid1.Cols = 2
MSHFlexGrid1.TextMatrix(0, 0) = "Name"
MSHFlexGrid1.TextMatrix(0, 1) = "AGE"

MSHFlexGrid1.TextMatrix(1, 0) = "Jaikumar"
MSHFlexGrid1.TextMatrix(1, 1) = "28"
MSHFlexGrid1.TextMatrix(2, 0) = "Suresh"
MSHFlexGrid1.TextMatrix(2, 1) = "27"
MSHFlexGrid1.TextMatrix(3, 0) = "Arun"
MSHFlexGrid1.TextMatrix(3, 1) = "26"
MSHFlexGrid1.TextMatrix(4, 0) = "Rajesh"
MSHFlexGrid1.TextMatrix(4, 1) = "31"
MSHFlexGrid1.TextMatrix(5, 0) = "TP"
MSHFlexGrid1.TextMatrix(5, 1) = "23"
MSHFlexGrid1.TextMatrix(6, 0) = "Udhay"
MSHFlexGrid1.TextMatrix(6, 1) = "27"
End Sub

Private Sub MSHFlexGrid1_Click()

If MSHFlexGrid1.TopRow = 1 And MSHFlexGrid1.RowSel = MSHFlexGrid1.Rows - 1 Then
    MsgBox "You have selected the column " & MSHFlexGrid1.Col
    If MSHFlexGrid1.Col = 1 Then
        MSHFlexGrid1.Sort = flexSortNumericDescending
    End If
    If MSHFlexGrid1.Col = 0 Then
        MSHFlexGrid1.Sort = flexSortStringAscending
    End If
End If

End Sub
