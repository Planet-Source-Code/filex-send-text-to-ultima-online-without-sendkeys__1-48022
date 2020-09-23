VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Send To UO"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   2835
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "testing"
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send It"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Text To Say:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form for sending text to Ultima Online's Window
' Code by Jason (jason@filex.org)

Private Sub Command1_Click()
Call SendUOText(Text1.text)
End Sub

Private Sub Form_Load()

End Sub
