VERSION 5.00
Begin VB.Form F1060401 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�I�ԃo�[�R�[�h���"
   ClientHeight    =   6315
   ClientLeft      =   2025
   ClientTop       =   2655
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   11295
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   3
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   3960
      MaxLength       =   2
      TabIndex        =   2
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   3240
      MaxLength       =   2
      TabIndex        =   1
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������~"
      Height          =   375
      Left            =   9000
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   0
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "��  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�m  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�|"
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   21
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�|"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   20
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�|"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   19
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Bar"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   28.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�I��"
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   16
      Top             =   2280
      Width           =   615
   End
End
Attribute VB_Name = "F1060401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PRT_CAN As Boolean                  '����r���L�����Z���v��
Dim NormalFont As New StdFont           '����t�H���g
Dim Code39Font As New StdFont           '����t�H���g

Dim S_Tana As String
Dim E_Tana As String

Dim Text_Max As Integer                 '��ʍ��ڕʍő���ޯ��
Dim Command_Max As Integer
Private Sub Clear_Field()
Dim i As Integer

    For i = 0 To Text_Max
        Text(i).Text = ""
    Next i
End Sub


Private Function Print_Proc() As Integer

Dim sts As Integer
Dim flg As Boolean
Dim com As Integer

    Print_Proc = False


    PRT_CAN = False
    flg = False
    Call UniCode_Conv(K0_TANA.Soko_No, Text(0).Text)
    Call UniCode_Conv(K0_TANA.Retu, Text(1).Text)
    Call UniCode_Conv(K0_TANA.Ren, Text(2).Text)
    Call UniCode_Conv(K0_TANA.Dan, Text(3).Text)
    
    
        
        sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�I�}�X�^")
                Beep
                MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
                Unload Me
        End Select
                                            '���׈��
                                            '����t�H���g�ݒ�
'        Printer.Print
'        Printer.Print
'        Printer.Print
        Printer.Print
        Set Printer.Font = Code39Font
        Printer.Print Tab(4);
        Printer.Print "*/" + StrConv(TANAREC.Soko_No, vbUnicode) + StrConv(TANAREC.Retu, vbUnicode) + StrConv(TANAREC.Ren, vbUnicode) + StrConv(TANAREC.Dan, vbUnicode) + "*"
        Printer.Print
        
        Set Printer.Font = NormalFont
        Printer.Print Tab(6);
        Printer.Print "*/" + StrConv(TANAREC.Soko_No, vbUnicode) + StrConv(TANAREC.Retu, vbUnicode) + StrConv(TANAREC.Ren, vbUnicode) + StrConv(TANAREC.Dan, vbUnicode) + "*"
        flg = True
        
    
        Printer.Print
        Printer.Print
        Printer.Print
        
        Printer.Print
        Printer.Print
        Printer.Print


    PRT_CAN = False

End Function

Private Sub Command_Click(Index As Integer)

Dim yn As Integer
Dim RetBuf As String
Dim i As Integer
    
    Select Case Index
        Case 0                              '���
                                            '�G���[�`�F�b�N
            If Len(Text(0).Text) = 0 Then
                S_Tana = Space(8)
            Else
                If Len(Text(0).Text) <> 0 Then
                    S_Tana = Text(0).Text
                    For i = 1 To 3
                        If Not IsNumeric(Text(i).Text) Then
                            Beep
                            MsgBox "���͂������ڂ̓G���[�ł��B", vbOKOnly + vbExclamation
                            Text(0).SelStart = 0
                            Text(0).SelLength = Len(Text(0).Text)
                            Text(0).SetFocus
                            Exit Sub
                        Else
                            S_Tana = S_Tana & Format(CInt(Text(i).Text), "00")
                            Text(i).Text = Format(CInt(Text(i).Text), "00")
                        End If
                    Next i
                Else
                    S_Tana = Space(8)
                End If

            End If

            Beep
            yn = MsgBox("�m�肵�܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                If Print_Proc() Then
                    Unload Me
                End If
'                Call Clear_Field
            End If
            Text(0).SelStart = 0
            Text(0).SelLength = Len(RTrim(Text(0).Text))
            Text(0).SetFocus
        Case 8                              '���
            Unload Me
        Case 11                             '�I��
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub

Private Sub Command1_Click()
    PRT_CAN = True
End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
    
    Text_Max = 3                '��ʍ��ڕʍő���ޯ��
    Command_Max = 11

    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        End
    End If
    
                                '�I�}�X�^�n�o�d�m
    If TANA_Open(0) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                '����t�H���g�ݒ�
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With
    Set Printer.Font = Code39Font
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F1060401.FontName
        .Size = 14
    End With
'    Set Printer.Font = NormalFont
                                
                                '��ʏ����ݒ�
    Call Clear_Field
    
    Text(0).SelStart = 0
    Text(0).SelLength = Len(RTrim(Text(0).Text))
    Text(0).SetFocus
    
    End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '�I�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
    
    sts = BTRV(BtOpReset, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "�I�}�X�^")
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
    End If

    End
End Sub


Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim RetBuf As String
Dim i As Integer

    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            If Index <> 0 Then
                If Not IsNumeric(Text(Index).Text) Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text(Index).SelStart = 0
                    Text(Index).SelLength = Len(Text(Index).Text)
                    Text(Index).SetFocus
                    Exit Sub
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                End If
            End If
            For i = Index + 1 To Text_Max
                If Text(i).Enabled Then
                    Text(i).SelStart = 0
                    Text(i).SelLength = Len(RTrim(Text(i).Text))
                    Text(i).SetFocus
                    Exit For
                End If
            Next i
        Case vbKeyUp
            For i = Index - 1 To 0 Step -1
                If Text(i).Enabled Then
                    Text(i).SelStart = 0
                    Text(i).SelLength = Len(RTrim(Text(i).Text))
                    Text(i).SetFocus
                    Exit For
                End If
            Next i
        Case vbKeyF1
            Command(0).Value = True
        Case vbKeyF9
            Command(8).Value = True
        Case vbKeyF12
            Command(11).Value = True
    End Select
End Sub

