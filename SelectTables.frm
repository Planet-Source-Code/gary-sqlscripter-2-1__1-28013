VERSION 5.00
Begin VB.Form frmSelectTables 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Tables to Create Insert Scripts for"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Invert"
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Unselect All"
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select All"
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ListBox lstTables 
      Height          =   2985
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSelectTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
DefInt A-Z

Private Sub CancelButton_Click()
    lstTables.Clear
    Me.Hide
End Sub

Private Sub cmdSelect_Click(Index As Integer)
    Dim i As Integer
    
    Select Case Index
        Case 0 ' select all
            For i = 0 To lstTables.ListCount - 1
                lstTables.Selected(i) = True
            Next
        Case 1 ' unselect
            For i = 0 To lstTables.ListCount - 1
                lstTables.Selected(i) = False
            Next
        Case 2 ' invert
            For i = 0 To lstTables.ListCount - 1
                lstTables.Selected(i) = Not lstTables.Selected(i)
            Next
    End Select
End Sub

Private Sub OKButton_Click()
    Me.Hide
End Sub
