VERSION 5.00
Begin VB.UserControl ocxValor 
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   ScaleHeight     =   510
   ScaleWidth      =   2685
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   60
      Width           =   2475
   End
End
Attribute VB_Name = "ocxValor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Property Get Text() As String
    Text = Text1.Text
End Property





Private Sub Text1_GotFocus()
    Text1.Text = ChkVal(Text1.Text, 0, 0)
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = ChkVal(Text1.Text, KeyAscii, CDecMoeda)
End Sub

Private Sub Text1_LostFocus()
    Text1.Text = ConvMoeda(Text1.Text)
End Sub

Private Sub UserControl_Initialize()
    Text1.Top = 0
    Text1.Left = 0
End Sub

Private Sub UserControl_Resize()
    Text1.Width = UserControl.Width
    Text1.Height = UserControl.Height
End Sub

