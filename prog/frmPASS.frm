VERSION 5.00
Begin VB.Form frmPASS 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'å≈íË¿ﬁ≤±€∏ﬁ
   ClientHeight    =   960
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   3600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ÇlÇr ÉSÉVÉbÉN"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'µ∞≈∞ Ã´∞—ÇÃíÜâõ
   Begin VB.TextBox Text1 
      Height          =   372
      IMEMode         =   3  'µÃå≈íË
      Index           =   0
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   1812
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "∑¨›æŸÅFEsc"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   10.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   4
      Left            =   2220
      TabIndex        =   2
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'ìßñæ
      Caption         =   "ÉpÉXÉèÅ[Éh"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1272
   End
End
Attribute VB_Name = "frmPASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Act_Flg As Integer                  'Activate∫›ƒ€∞ŸÃ◊∏ﬁ

Private Sub Form_Activate()

    If Act_Flg = True Then Exit Sub

    Act_Flg = True              'Activate∫›ƒ€∞ŸÃ◊∏ﬁ

    Text1(0).Text = ""
    Text1(0).SelStart = 0
    Text1(0).SelLength = Len(Text1(0).Text)
    Text1(0).SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Act_Flg = False         'Activate∫›ƒ€∞ŸÃ◊∏ﬁ
        Me.Visible = False
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

    Act_Flg = False             'Activate∫›ƒ€∞ŸÃ◊∏ﬁ

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim Str_Wk  As String

    If KeyCode <> vbKeyReturn Then Exit Sub


    Str_Wk = StrConv(Text1(0).Text, vbUpperCase)

    If Str_Wk <> frmMENU.L_Pass(0).Caption Then
        Beep
        Text1(0).SelStart = 0
        Text1(0).SelLength = Len(Text1(0).Text)
        Text1(0).SetFocus
        Exit Sub
    End If

    frmMENU.L_Pass(1).Caption = frmMENU.L_Pass(0).Caption
    Act_Flg = False         'Activate∫›ƒ€∞ŸÃ◊∏ﬁ
    Me.Visible = False

End Sub
