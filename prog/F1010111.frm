VERSION 5.00
Begin VB.Form F1010111 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�q�Ƀ}�X�^�����e�i���X(�q�Ɂ��I�}�X�^)"
   ClientHeight    =   6300
   ClientLeft      =   1920
   ClientTop       =   2430
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
   ScaleHeight     =   6300
   ScaleWidth      =   11295
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���i���ݒ�"
      Height          =   1695
      Left            =   6480
      TabIndex        =   41
      Top             =   2040
      Width           =   3615
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���i���p�ȊO"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   43
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "���i���p�q��"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   42
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   7320
      MaxLength       =   3
      TabIndex        =   39
      Top             =   1440
      Width           =   495
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   2
      Left            =   6000
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   12
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   11
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   10
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   8
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1560
      Width           =   375
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   4
      Left            =   9120
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   5
      Left            =   8880
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   3480
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   0
      Left            =   1560
      MaxLength       =   16
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   3
      Left            =   6000
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   3
      Top             =   840
      Width           =   1935
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
      TabIndex        =   24
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
      Index           =   10
      Left            =   9480
      TabIndex        =   23
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
      Index           =   9
      Left            =   8640
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�f�[�^"
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
      TabIndex        =   21
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
      Index           =   7
      Left            =   6480
      TabIndex        =   20
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
      Index           =   6
      Left            =   5640
      TabIndex        =   19
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
      Index           =   5
      Left            =   4800
      TabIndex        =   18
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
      Index           =   4
      Left            =   3960
      TabIndex        =   17
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
      Index           =   3
      Left            =   2640
      TabIndex        =   16
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
      Index           =   2
      Left            =   1800
      TabIndex        =   15
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�X  �V"
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�O�`�P�O�O�j"
      Height          =   255
      Index           =   14
      Left            =   7920
      TabIndex        =   40
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������"
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   38
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ڋ敪"
      Height          =   255
      Index           =   13
      Left            =   4920
      TabIndex        =   37
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   255
      Index           =   12
      Left            =   5280
      TabIndex        =   36
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i"
      Height          =   255
      Index           =   11
      Left            =   4200
      TabIndex        =   35
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   255
      Index           =   10
      Left            =   5280
      TabIndex        =   34
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�A"
      Height          =   255
      Index           =   9
      Left            =   4200
      TabIndex        =   33
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   255
      Index           =   8
      Left            =   5280
      TabIndex        =   32
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   7
      Left            =   4200
      TabIndex        =   31
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   255
      Index           =   6
      Left            =   8280
      TabIndex        =   30
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�g�p��"
      Height          =   255
      Index           =   5
      Left            =   7800
      TabIndex        =   29
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�q�ɕ���"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   28
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�q�ɖ���"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   27
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�q�ɇ�"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   26
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ƕ��敪"
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   25
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "F1010111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Text_Max    As Integer                 '��ʍ��ڕʍő���ޯ��
Dim Combo_Max   As Integer
Dim Command_Max As Integer
Dim Soko_Csv    As String

Dim To_Ido_Soko As String * 2
Dim To_Ido_Yoin As String * 2

Dim Zaiko_Flg   As Boolean

Dim Ws_No       As String * 3


Private Function Data_Proc() As Integer

Dim FileNo          As Integer
Dim fileName        As String
Dim Ret             As Integer

Dim com             As Integer
Dim sts             As Integer

Dim c               As String * 128

    Call Input_Lock

    FileNo = FreeFile
    fileName = Soko_Csv
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & Right(Trim(fileName), Len(Trim(fileName)) - Ret)

    On Error GoTo Error_Proc

    Open (fileName) For Output As FileNo
    
    Write #FileNo, "���ƕ��敪", "�q�ɇ�", "�q�ɖ���", "�q�ɕ���", "�q�ɋ敪", "�����O", "�g�p��", "���ډ�", "��͈�", "�A�͈�", "�i�͈�", "�����_", "���i���q��"

    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�q�Ƀ}�X�^")
                Exit Function
        End Select
    
        Write #FileNo, StrConv(SOKOREC.JGYOBU, vbUnicode),
        
        
        If GetIni("SOKO_NO", StrConv(SOKOREC.Soko_No, vbUnicode), "SYS", c) Then
            Write #FileNo, StrConv(SOKOREC.Soko_No, vbUnicode),
        Else
            Write #FileNo, Trim(c),
        End If
        Write #FileNo, StrConv(SOKOREC.SOKO_NAME, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.SOKO_BUN, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.SOKO_KBN, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.NAIGAI, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.KAHI_KBN, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.KONS_KBN, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.RETU_START, vbUnicode) & "�`" & StrConv(SOKOREC.RETU_END, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.REN_START, vbUnicode) & "�`" & StrConv(SOKOREC.REN_END, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.DAN_START, vbUnicode) & "�`" & StrConv(SOKOREC.DAN_END, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.ORDER_POINT, vbUnicode),
        Write #FileNo, StrConv(SOKOREC.GOODS_ON_F, vbUnicode)
    
    
        com = BtOpGetNext
    Loop

    Close #FileNo
    
    Call Input_UnLock
    
    Beep
    MsgBox "�u" & fileName & "�v�͐���ɏo�͂���܂����B"


    Exit Function
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "���g�p���ł��B"
        Data_Proc = False
    Else
        MsgBox "Err.Number= " & Err.Number
        Data_Proc = True
    End If


    Call Input_UnLock



End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i As Integer

    F1010111.MousePointer = vbHourglass

    Call Ctrl_Lock(F1010111)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1010111)


    F1010111.MousePointer = vbDefault

End Sub
                                    '�S�q�Ƀ}�X�^�̓ǂݍ���
Private Function List_Proc()
Dim sts As Integer
Dim com As Integer
    
    List_Proc = False
    
    Combo(0).Clear
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�q�Ƀ}�X�^")
                List_Proc = True
                Exit Function
        End Select
        
        Combo(0).AddItem (StrConv(SOKOREC.Soko_No, vbUnicode))
        com = BtOpGetNext
    Loop
    
End Function
                                    '��ʏ�����Ԃ�ݒ肷��
Private Sub Clear_Field(Mode As Integer)
Dim i As Integer

    If (Mode = 0) Then
        Combo(0).Text = ""
    End If
    Combo(1).Text = SOKO_BUN0$
    Combo(2).Text = KONS_KBN0$
    Combo(3).Text = "�i�Ȃ��j"
    Combo(4).Text = NAIGAI0$
    Combo(5).Text = KAHI_KBN0$
    
    For i = 0 To 7
        Text(i).Text = ""
    Next i
                
    Combo(2).Enabled = True
    Combo(3).Enabled = True
    Combo(4).Enabled = True
    For i = 1 To 6
        Text(i).Enabled = True
    Next i

    Option1(0).Value = False
    Option1(1).Value = True
    

End Sub

'                                       ���͍��ڂ̃G���[�`�F�b�N
Private Function Err_Chk() As Integer
            
Dim RetBuf As String
Dim i As Integer


    Err_Chk = False
    If Len(Combo(0).Text) = 0 Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Combo(0).SelStart = 0
        Combo(0).SelLength = Len(Combo(0).Text)
        Combo(0).SetFocus
        Err_Chk = True
        Exit Function
    End If
            
    If Combo(2).Text = KONS_KBN1$ Then
        For i = 0 To UBound(JGYOBU_T)
            If Combo(3).Text = RTrim(JGYOBU_T(i).NAME) Then
                If JGYOBU_T(i).Code = "0" Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Combo(2).SetFocus
                    Err_Chk = True
                    Exit Function
                End If
            End If
        Next i
    
        If Combo(4).Text = NAIGAI0$ Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B"
            Combo(2).SetFocus
            Err_Chk = True
            Exit Function
        End If
    End If
                    
            
            
            
    For i = 1 To 6
        If Not IsNumeric(Text(i).Text) Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B"
            Err_Chk = True
            Text(i).SelStart = 0
            Text(i).SelLength = Len(Text(i).Text)
            Text(i).SetFocus
            Exit Function
        Else
            Text(i).Text = Format(CInt(Text(i).Text), "00")
        End If
    Next i
    
    For i = 1 To 5 Step 2
        If Text(i).Text > Text(i + 1).Text Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B"
            Text(i).SelStart = 0
            Text(i).SelLength = Len(Text(i).Text)
            Text(i).SetFocus
            Err_Chk = True
            Exit Function
        End If
    Next i
    
    
    If Text(7).Text = "" Then
        Text(7).Text = "   "
    End If
    
    If Not IsNumeric(Text(7).Text) Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Err_Chk = True
        Text(7).SelStart = 0
        Text(7).SelLength = Len(Text(7).Text)
        Text(7).SetFocus
        Exit Function
    Else
        Text(7).Text = Format(CInt(Text(7).Text), "#0")
    End If


End Function

Private Function Item_Dsp() As Integer
Dim sts As Integer
Dim i As Integer

    Item_Dsp = False
    Call UniCode_Conv(K0_SOKO.Soko_No, Combo(0).Text)
    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    Select Case sts
        Case BtNoErr
            For i = 0 To UBound(JGYOBU_T)
                If JGYOBU_T(i).Code = StrConv(SOKOREC.JGYOBU, vbUnicode) Then
                    Combo(3).Text = RTrim(JGYOBU_T(i).NAME)
                    Exit For
                End If
                                                    '��O�����i�Ȃ��͂��j
                If JGYOBU_T(i).Code = " " Then
                    Combo(3).Text = "�i�Ȃ��j"
                    Exit For
                End If
            Next i
            Select Case StrConv(SOKOREC.SOKO_BUN, vbUnicode)
                Case BUN_JITU$
                    Combo(1).Text = SOKO_BUN0$
                Case BUN_KASO$
                    Combo(1).Text = SOKO_BUN1$
'                Case bun_AUTO$%
'                    Combo(1).Text = SOKO_bun2$
            End Select
            Select Case StrConv(SOKOREC.NAIGAI, vbUnicode)
                Case NAIGAI_NON$
                    Combo(4).Text = NAIGAI0$
                Case NAIGAI_NAI$
                    Combo(4).Text = NAIGAI1$
                Case NAIGAI_GAI$
                    Combo(4).Text = NAIGAI2$
            End Select
            Select Case StrConv(SOKOREC.KAHI_KBN, vbUnicode)
                Case KAHI_KBN_OK$
                    Combo(5).Text = KAHI_KBN0$
                Case KAHI_KBN_NG$
                    Combo(5).Text = KAHI_KBN1$
            End Select
            Select Case StrConv(SOKOREC.KONS_KBN, vbUnicode)
                Case KONS_KBN_OK$
                    Combo(2).Text = KONS_KBN0$
                Case KONS_KBN_NG$
                    Combo(2).Text = KONS_KBN1$
            End Select
        
            Text(0).Text = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
            Text(1).Text = StrConv(SOKOREC.RETU_START, vbUnicode)
            Text(2).Text = StrConv(SOKOREC.RETU_END, vbUnicode)
            Text(3).Text = StrConv(SOKOREC.REN_START, vbUnicode)
            Text(4).Text = StrConv(SOKOREC.REN_END, vbUnicode)
            Text(5).Text = StrConv(SOKOREC.DAN_START, vbUnicode)
            Text(6).Text = StrConv(SOKOREC.DAN_END, vbUnicode)
                        
                        
            If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
            End If
            
            Text(7).Text = Format(CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0")
                        
            If StrConv(SOKOREC.GOODS_ON_F, vbUnicode) = GOODS_ON Then
                Option1(0).Value = True
                Option1(1).Value = False
            Else
                Option1(0).Value = False
                Option1(1).Value = True
            End If
                        
            If Combo(1).Text = SOKO_BUN1$ Then
                Combo(2).Enabled = False
                Combo(3).Enabled = False
                Combo(4).Enabled = False
                For i = 1 To 6
                    Text(i).Enabled = False
                Next i
            
                Text(7).Enabled = False
            
                Frame1.Enabled = True
            Else
                Combo(2).Enabled = True
                Combo(3).Enabled = True
                Combo(4).Enabled = True
                For i = 1 To 6
                    Text(i).Enabled = True
                Next i
            
            
                Text(7).Enabled = True
            
                Frame1.Enabled = False
            
            
            End If
            
            If Combo(2).Text = KONS_KBN0$ Then
                Combo(3).Enabled = False
                Combo(4).Enabled = False
            Else
                Combo(3).Enabled = True
                Combo(4).Enabled = True
            End If
            
        Case BtErrKeyNotFound
            Call Clear_Field(1)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
            Item_Dsp = True
    End Select

End Function

Private Function Update_Proc() As Integer
                                            '�q�Ƀ}�X�^�̒ǉ��^����

Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
Dim i               As Integer

Dim OLD_Retu_Start  As String * 2
Dim OLD_Retu_End    As String * 2

Dim OLD_Ren_Start   As String * 2
Dim OLD_Ren_End     As String * 2

Dim OLD_Dan_Start   As String * 2
Dim OLD_Dan_End     As String * 2

Dim Retu            As Integer
Dim Ren             As Integer
Dim Dan             As Integer

Dim Upd_com         As Integer



    Update_Proc = True

    Call Input_Lock

    Call UniCode_Conv(K0_SOKO.Soko_No, Combo(0).Text)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SOKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Call Input_UnLock
                    Call Clear_Field(0)
                    Exit Function
                End If
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�q�Ƀ}�X�^")
                Exit Function
        End Select
    
        DoEvents
    
    Loop
                                            
    '���@��^�A�^�i�͈͂�ۑ�����
    If com = BtOpUpdate Then
    
        OLD_Retu_Start = StrConv(SOKOREC.RETU_START, vbUnicode)
        OLD_Retu_End = StrConv(SOKOREC.RETU_END, vbUnicode)
                                            
        OLD_Ren_Start = StrConv(SOKOREC.REN_START, vbUnicode)
        OLD_Ren_End = StrConv(SOKOREC.REN_END, vbUnicode)
                                            
        OLD_Dan_Start = StrConv(SOKOREC.DAN_START, vbUnicode)
        OLD_Dan_End = StrConv(SOKOREC.DAN_END, vbUnicode)
                                            
                                            
    End If
                                            
                                            
                                            
                                            '���R�[�h���e�ҏW
    For i = 0 To UBound(JGYOBU_T)
        If RTrim(JGYOBU_T(i).NAME) = Combo(3).Text Then
            Call UniCode_Conv(SOKOREC.JGYOBU, JGYOBU_T(i).Code)
            Exit For
        End If
                                             '��O�����i�Ȃ��͂��j
        If JGYOBU_T(i).Code = " " Then
            Call UniCode_Conv(SOKOREC.JGYOBU, "0")
            Exit For
        End If
    Next i
    
    If i > UBound(JGYOBU_T) Then
        Call UniCode_Conv(SOKOREC.JGYOBU, "0")
    End If
    
    Call UniCode_Conv(SOKOREC.Soko_No, Combo(0).Text)
    Call UniCode_Conv(SOKOREC.SOKO_NAME, Text(0).Text)

    Select Case RTrim(Combo(1).Text)
        Case RTrim(SOKO_BUN0$)
            Call UniCode_Conv(SOKOREC.SOKO_BUN, BUN_JITU$)
        Case RTrim(SOKO_BUN1$)
            Call UniCode_Conv(SOKOREC.SOKO_BUN, BUN_KASO$)
'        Case SOKO_KBN2$
'            Call UniCode_Conv(SOKOREC.SOKO_bun, bun_AUTO$)
    End Select
    
    Call UniCode_Conv(SOKOREC.SOKO_KBN, "")
    
    Select Case Combo(4).Text
        Case NAIGAI0$
            Call UniCode_Conv(SOKOREC.NAIGAI, NAIGAI_NON$)
        Case NAIGAI1$
            Call UniCode_Conv(SOKOREC.NAIGAI, NAIGAI_NAI$)
        Case NAIGAI2$
            Call UniCode_Conv(SOKOREC.NAIGAI, NAIGAI_GAI$)
    End Select
    Select Case Combo(5).Text
        Case KAHI_KBN0$
            Call UniCode_Conv(SOKOREC.KAHI_KBN, KAHI_KBN_OK$)
        Case KAHI_KBN1$
            Call UniCode_Conv(SOKOREC.KAHI_KBN, KAHI_KBN_NG$)
    End Select
    Select Case Combo(2).Text
        Case KONS_KBN0$
            Call UniCode_Conv(SOKOREC.KONS_KBN, KONS_KBN_OK$)
        Case KONS_KBN1$
            Call UniCode_Conv(SOKOREC.KONS_KBN, KONS_KBN_NG$)
    End Select

'    If (StrConv(SOKOREC.SOKO_KBN, vbUnicode) = KBN_KASO$) Then
'        Call UniCode_Conv(SOKOREC.RETU_START, "00")
'        Call UniCode_Conv(SOKOREC.RETU_END, "00")
'        Call UniCode_Conv(SOKOREC.REN_START, "00")
'        Call UniCode_Conv(SOKOREC.REN_END, "00")
'        Call UniCode_Conv(SOKOREC.DAN_START, "00")
'        Call UniCode_Conv(SOKOREC.DAN_END, "00")
'    Else
        Call UniCode_Conv(SOKOREC.RETU_START, Text(1).Text)
        Call UniCode_Conv(SOKOREC.RETU_END, Text(2).Text)
        Call UniCode_Conv(SOKOREC.REN_START, Text(3).Text)
        Call UniCode_Conv(SOKOREC.REN_END, Text(4).Text)
        Call UniCode_Conv(SOKOREC.DAN_START, Text(5).Text)
        Call UniCode_Conv(SOKOREC.DAN_END, Text(6).Text)
        Call UniCode_Conv(SOKOREC.FILLER, "")
'   End If
    
    If (StrConv(SOKOREC.SOKO_KBN, vbUnicode) = BUN_KASO$) Then
        Call UniCode_Conv(SOKOREC.ORDER_POINT, "")
    Else
        Call UniCode_Conv(SOKOREC.ORDER_POINT, Format(CInt(Text(7).Text), "000"))
    End If


    If Option1(0).Value = True Then
        Call UniCode_Conv(SOKOREC.GOODS_ON_F, GOODS_ON)
    Else
        Call UniCode_Conv(SOKOREC.GOODS_ON_F, GOODS_OFF)
    End If

    Do
        sts = BTRV(com, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SOKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Call Input_UnLock
                    Call Clear_Field(0)
                    Exit Function
                End If
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "�q�Ƀ}�X�^")
                Exit Function
        End Select
    Loop
    
    '�I�}�X�^�̒ǉ�����
    If com = BtOpInsert Then
        '�V�K�ǉ����͑S���P�[�V�����ǉ�
        For Retu = CInt(Text(1).Text) To CInt(Text(2).Text)
            For Ren = CInt(Text(3).Text) To CInt(Text(4).Text)
                For Dan = CInt(Text(5).Text) To CInt(Text(6).Text)
                
                
                    Call UniCode_Conv(K0_TANA.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Format(Retu, "00"))
                    Call UniCode_Conv(K0_TANA.Ren, Format(Ren, "00"))
                    Call UniCode_Conv(K0_TANA.Dan, Format(Dan, "00"))
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                Upd_com = BtOpInsert
                                Exit Do
                                    '����͖���
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
                                    '�I�f�[�^�X�V�^�ǉ�
'                           Call UniCode_Conv(TANAREC.JGYOBU, StrConv(SOKOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(TANAREC.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(TANAREC.Retu, Format$(Retu, "00"))
                    Call UniCode_Conv(TANAREC.Ren, Format$(Ren, "00"))
                    Call UniCode_Conv(TANAREC.Dan, Format$(Dan, "00"))
                    Call UniCode_Conv(TANAREC.KAHI_KBN, StrConv(SOKOREC.KAHI_KBN, vbUnicode))
'                          Call UniCode_Conv(TANAREC.NAIGAI, StrConv(SOKOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(TANAREC.TANA_COND, "0")
'                         Call UniCode_Conv(TANAREC.KONS_KBN, StrConv(SOKOREC.KONS_KBN, vbUnicode))
                    
                    
                    Call UniCode_Conv(TANAREC.ZAIKO_SHOGO_FLG, ZAIKO_SHOGO_FLG_OK) '�݌ɏƍ��@�u�O�v   2004.02
                    
                    Call UniCode_Conv(TANAREC.FILLER, "")
                    Do
                        sts = BTRV(Upd_com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
                                    
                Next Dan
            Next Ren
        Next Retu
                                        
                                        
    Else
       '-----------------------    �X�V���͑������̂ݒǉ�  ---------------------------
        
        '��̏����i�J�n�ʒu�j---------------------------------------------
        For Retu = CInt(OLD_Retu_Start) - 1 To CInt(Text(1).Text) Step -1
        
            For Ren = CInt(OLD_Ren_Start) To CInt(OLD_Ren_End)
                
                For Dan = CInt(OLD_Dan_Start) To CInt(OLD_Dan_End)
                
                
                    Call UniCode_Conv(K0_TANA.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Format(Retu, "00"))
                    Call UniCode_Conv(K0_TANA.Ren, Format(Ren, "00"))
                    Call UniCode_Conv(K0_TANA.Dan, Format(Dan, "00"))
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                Upd_com = BtOpInsert
                                Exit Do
                                    '����͖���
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
                                    '�I�f�[�^�X�V�^�ǉ�
'                           Call UniCode_Conv(TANAREC.JGYOBU, StrConv(SOKOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(TANAREC.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(TANAREC.Retu, Format$(Retu, "00"))
                    Call UniCode_Conv(TANAREC.Ren, Format$(Ren, "00"))
                    Call UniCode_Conv(TANAREC.Dan, Format$(Dan, "00"))
                    Call UniCode_Conv(TANAREC.KAHI_KBN, StrConv(SOKOREC.KAHI_KBN, vbUnicode))
'                          Call UniCode_Conv(TANAREC.NAIGAI, StrConv(SOKOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(TANAREC.TANA_COND, "0")
'                         Call UniCode_Conv(TANAREC.KONS_KBN, StrConv(SOKOREC.KONS_KBN, vbUnicode))
                    
                    
                    Call UniCode_Conv(TANAREC.ZAIKO_SHOGO_FLG, ZAIKO_SHOGO_FLG_OK) '�݌ɏƍ��@�u�O�v   2004.02
                    
                    Call UniCode_Conv(TANAREC.FILLER, "")
                    Do
                        sts = BTRV(Upd_com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
                
                
                
                
                Next Dan
        
            Next Ren
        
        
        Next Retu
        
        
        
        '��̏����i�I���ʒu�j---------------------------------------------
        For Retu = CInt(OLD_Retu_End) + 1 To CInt(Text(2).Text)
        
            For Ren = CInt(OLD_Ren_Start) To CInt(OLD_Ren_End)
                
                For Dan = CInt(OLD_Dan_Start) To CInt(OLD_Dan_End)
                
                
                    Call UniCode_Conv(K0_TANA.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Format(Retu, "00"))
                    Call UniCode_Conv(K0_TANA.Ren, Format(Ren, "00"))
                    Call UniCode_Conv(K0_TANA.Dan, Format(Dan, "00"))
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                Upd_com = BtOpInsert
                                Exit Do
                                    '����͖���
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
                                    '�I�f�[�^�X�V�^�ǉ�
'                           Call UniCode_Conv(TANAREC.JGYOBU, StrConv(SOKOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(TANAREC.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(TANAREC.Retu, Format$(Retu, "00"))
                    Call UniCode_Conv(TANAREC.Ren, Format$(Ren, "00"))
                    Call UniCode_Conv(TANAREC.Dan, Format$(Dan, "00"))
                    Call UniCode_Conv(TANAREC.KAHI_KBN, StrConv(SOKOREC.KAHI_KBN, vbUnicode))
'                          Call UniCode_Conv(TANAREC.NAIGAI, StrConv(SOKOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(TANAREC.TANA_COND, "0")
'                         Call UniCode_Conv(TANAREC.KONS_KBN, StrConv(SOKOREC.KONS_KBN, vbUnicode))
                    
                    
                    Call UniCode_Conv(TANAREC.ZAIKO_SHOGO_FLG, ZAIKO_SHOGO_FLG_OK) '�݌ɏƍ��@�u�O�v   2004.02
                    
                    Call UniCode_Conv(TANAREC.FILLER, "")
                    Do
                        sts = BTRV(Upd_com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
                
                
                
                
                Next Dan
        
            Next Ren
        
        
        Next Retu
                                        
                                        
                                        
        '�A�̏���(�J�n�ʒu) ---------------------------------------------
        For Retu = CInt(Text(1).Text) To CInt(Text(2).Text)
            For Ren = CInt(OLD_Ren_Start) - 1 To CInt(Text(3).Text) Step -1
                For Dan = CInt(OLD_Dan_Start) To CInt(OLD_Dan_End)
            
            
                    Call UniCode_Conv(K0_TANA.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Format(Retu, "00"))
                    Call UniCode_Conv(K0_TANA.Ren, Format(Ren, "00"))
                    Call UniCode_Conv(K0_TANA.Dan, Format(Dan, "00"))
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                Upd_com = BtOpInsert
                                Exit Do
                                    '����͖���
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
                                    '�I�f�[�^�X�V�^�ǉ�
    '                           Call UniCode_Conv(TANAREC.JGYOBU, StrConv(SOKOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(TANAREC.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(TANAREC.Retu, Format$(Retu, "00"))
                    Call UniCode_Conv(TANAREC.Ren, Format$(Ren, "00"))
                    Call UniCode_Conv(TANAREC.Dan, Format$(Dan, "00"))
                    Call UniCode_Conv(TANAREC.KAHI_KBN, StrConv(SOKOREC.KAHI_KBN, vbUnicode))
    '                          Call UniCode_Conv(TANAREC.NAIGAI, StrConv(SOKOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(TANAREC.TANA_COND, "0")
    '                         Call UniCode_Conv(TANAREC.KONS_KBN, StrConv(SOKOREC.KONS_KBN, vbUnicode))
                    
                    
                    Call UniCode_Conv(TANAREC.ZAIKO_SHOGO_FLG, ZAIKO_SHOGO_FLG_OK) '�݌ɏƍ��@�u�O�v   2004.02
                    
                    Call UniCode_Conv(TANAREC.FILLER, "")
                    Do
                        sts = BTRV(Upd_com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
            
            
            
            
                Next Dan
            Next Ren
        Next Retu
                                        
                                        
        '�A�̏���(�I���ʒu) ---------------------------------------------
        For Retu = CInt(Text(1).Text) To CInt(Text(2).Text)
            For Ren = CInt(OLD_Ren_End) + 1 To CInt(Text(4).Text)
                For Dan = CInt(OLD_Dan_Start) To CInt(OLD_Dan_End)
            
            
                    Call UniCode_Conv(K0_TANA.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Format(Retu, "00"))
                    Call UniCode_Conv(K0_TANA.Ren, Format(Ren, "00"))
                    Call UniCode_Conv(K0_TANA.Dan, Format(Dan, "00"))
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                Upd_com = BtOpInsert
                                Exit Do
                                    '����͖���
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
                                    '�I�f�[�^�X�V�^�ǉ�
    '                           Call UniCode_Conv(TANAREC.JGYOBU, StrConv(SOKOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(TANAREC.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(TANAREC.Retu, Format$(Retu, "00"))
                    Call UniCode_Conv(TANAREC.Ren, Format$(Ren, "00"))
                    Call UniCode_Conv(TANAREC.Dan, Format$(Dan, "00"))
                    Call UniCode_Conv(TANAREC.KAHI_KBN, StrConv(SOKOREC.KAHI_KBN, vbUnicode))
    '                          Call UniCode_Conv(TANAREC.NAIGAI, StrConv(SOKOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(TANAREC.TANA_COND, "0")
    '                         Call UniCode_Conv(TANAREC.KONS_KBN, StrConv(SOKOREC.KONS_KBN, vbUnicode))
                    
                    
                    Call UniCode_Conv(TANAREC.ZAIKO_SHOGO_FLG, ZAIKO_SHOGO_FLG_OK) '�݌ɏƍ��@�u�O�v   2004.02
                    
                    Call UniCode_Conv(TANAREC.FILLER, "")
                    Do
                        sts = BTRV(Upd_com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
            
            
            
            
                Next Dan
            Next Ren
        Next Retu
                                        
                                        
                                        
        '�i�̏���(�J�n�ʒu) ---------------------------------------------
        For Retu = CInt(Text(1).Text) To CInt(Text(2).Text)
            For Ren = CInt(Text(3).Text) To CInt(Text(4).Text)
                For Dan = CInt(OLD_Dan_End) - 1 To CInt(Text(5).Text) Step -1
            
            
                    Call UniCode_Conv(K0_TANA.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Format(Retu, "00"))
                    Call UniCode_Conv(K0_TANA.Ren, Format(Ren, "00"))
                    Call UniCode_Conv(K0_TANA.Dan, Format(Dan, "00"))
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                Upd_com = BtOpInsert
                                Exit Do
                                    '����͖���
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
                                    '�I�f�[�^�X�V�^�ǉ�
    '                           Call UniCode_Conv(TANAREC.JGYOBU, StrConv(SOKOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(TANAREC.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(TANAREC.Retu, Format$(Retu, "00"))
                    Call UniCode_Conv(TANAREC.Ren, Format$(Ren, "00"))
                    Call UniCode_Conv(TANAREC.Dan, Format$(Dan, "00"))
                    Call UniCode_Conv(TANAREC.KAHI_KBN, StrConv(SOKOREC.KAHI_KBN, vbUnicode))
    '                          Call UniCode_Conv(TANAREC.NAIGAI, StrConv(SOKOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(TANAREC.TANA_COND, "0")
    '                         Call UniCode_Conv(TANAREC.KONS_KBN, StrConv(SOKOREC.KONS_KBN, vbUnicode))
                    
                    
                    Call UniCode_Conv(TANAREC.ZAIKO_SHOGO_FLG, ZAIKO_SHOGO_FLG_OK) '�݌ɏƍ��@�u�O�v   2004.02
                    
                    Call UniCode_Conv(TANAREC.FILLER, "")
                    Do
                        sts = BTRV(Upd_com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
            
            
            
            
                Next Dan
            Next Ren
        Next Retu
                                        
        '�i�̏���(�I���ʒu) ---------------------------------------------
        For Retu = CInt(Text(1).Text) To CInt(Text(2).Text)
            For Ren = CInt(Text(3).Text) To CInt(Text(4).Text)
                For Dan = CInt(OLD_Dan_End) + 1 To CInt(Text(6).Text)
            
            
                    Call UniCode_Conv(K0_TANA.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Format(Retu, "00"))
                    Call UniCode_Conv(K0_TANA.Ren, Format(Ren, "00"))
                    Call UniCode_Conv(K0_TANA.Dan, Format(Dan, "00"))
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                Upd_com = BtOpInsert
                                Exit Do
                                    '����͖���
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
                                    '�I�f�[�^�X�V�^�ǉ�
    '                           Call UniCode_Conv(TANAREC.JGYOBU, StrConv(SOKOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(TANAREC.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(TANAREC.Retu, Format$(Retu, "00"))
                    Call UniCode_Conv(TANAREC.Ren, Format$(Ren, "00"))
                    Call UniCode_Conv(TANAREC.Dan, Format$(Dan, "00"))
                    Call UniCode_Conv(TANAREC.KAHI_KBN, StrConv(SOKOREC.KAHI_KBN, vbUnicode))
    '                          Call UniCode_Conv(TANAREC.NAIGAI, StrConv(SOKOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(TANAREC.TANA_COND, "0")
    '                         Call UniCode_Conv(TANAREC.KONS_KBN, StrConv(SOKOREC.KONS_KBN, vbUnicode))
                    
                    
                    Call UniCode_Conv(TANAREC.ZAIKO_SHOGO_FLG, ZAIKO_SHOGO_FLG_OK) '�݌ɏƍ��@�u�O�v   2004.02
                    
                    Call UniCode_Conv(TANAREC.FILLER, "")
                    Do
                        sts = BTRV(Upd_com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
            
            
            
            
                Next Dan
            Next Ren
        Next Retu
                                        
                                        
       '-----------------------    �X�V���͌��Z���̂ݍ폜  ---------------------------
                                        
        '��̏����i�J�n�ʒu�j---------------------------------------------
        
        
        For Retu = CInt(OLD_Retu_Start) To CInt(Text(1).Text) - 1
        
            For Ren = CInt(OLD_Ren_Start) To CInt(OLD_Ren_End)
                
                For Dan = CInt(OLD_Dan_Start) To CInt(OLD_Dan_End)
                    
                    Call UniCode_Conv(K0_TANA.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Format(Retu, "00"))
                    Call UniCode_Conv(K0_TANA.Ren, Format(Ren, "00"))
                    Call UniCode_Conv(K0_TANA.Dan, Format(Dan, "00"))
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                Upd_com = BtOpInsert
                                Exit Do
                                    '����͖���
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
                                    
                    If Upd_com = BtOpUpdate Then
                                    
                                    
'                        If Zaiko_Check_Proc(Zaiko_Flg) Then
'                            Unload Me
'                        End If
                                    '�I�f�[�^�폜
                        
                        
                        Do
                            sts = BTRV(BtOpDelete, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                    If ans = vbCancel Then
                                        Call Input_UnLock
                                        Exit Function
                                    End If
                                Case Else
                                    Call Input_UnLock
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                    Exit Function
                            End Select
                        
                        Loop
                    
                    End If
                
                
                
                Next Dan
        
            Next Ren
        
        
        Next Retu
        
        
        '��̏����i�I���ʒu�j---------------------------------------------
        
        For Retu = CInt(OLD_Retu_End) To CInt(Text(2).Text) + 1 Step -1
        
            For Ren = CInt(OLD_Ren_Start) To CInt(OLD_Ren_End)
                
                For Dan = CInt(OLD_Dan_Start) To CInt(OLD_Dan_End)
                
                
                    Call UniCode_Conv(K0_TANA.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Format(Retu, "00"))
                    Call UniCode_Conv(K0_TANA.Ren, Format(Ren, "00"))
                    Call UniCode_Conv(K0_TANA.Dan, Format(Dan, "00"))
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                Upd_com = BtOpInsert
                                Exit Do
                                    '����͖���
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
                    
                    If Upd_com = BtOpUpdate Then
                                    
'                        If Zaiko_Check_Proc(Zaiko_Flg) Then
'                            Unload Me
'                        End If
                                    '�I�f�[�^�폜
                        Do
                            sts = BTRV(BtOpDelete, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                    If ans = vbCancel Then
                                        Call Input_UnLock
                                        Exit Function
                                    End If
                                Case Else
                                    Call Input_UnLock
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                    Exit Function
                            End Select
                        
                        Loop
                    
                    End If
                
                
                
                
                Next Dan
        
            Next Ren
        
        
        Next Retu
                                        
                                        
        '�A�̏���(�J�n�ʒu) ---------------------------------------------
        For Retu = CInt(Text(1).Text) To CInt(Text(2).Text)
            
            For Ren = CInt(OLD_Ren_Start) To CInt(Text(3).Text) - 1
                
                For Dan = CInt(OLD_Dan_Start) To CInt(OLD_Dan_End)
            
            
                    Call UniCode_Conv(K0_TANA.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Format(Retu, "00"))
                    Call UniCode_Conv(K0_TANA.Ren, Format(Ren, "00"))
                    Call UniCode_Conv(K0_TANA.Dan, Format(Dan, "00"))
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                Upd_com = BtOpInsert
                                Exit Do
                                    '����͖���
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
            
                    If Upd_com = BtOpUpdate Then
                                    
                                    
'                        If Zaiko_Check_Proc(Zaiko_Flg) Then
'                            Unload Me
'                        End If
                                    '�I�f�[�^�폜
                        Do
                            sts = BTRV(BtOpDelete, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                    If ans = vbCancel Then
                                        Call Input_UnLock
                                        Exit Function
                                    End If
                                Case Else
                                    Call Input_UnLock
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                    Exit Function
                            End Select
                        
                        Loop
                    
                    End If
            
            
            
                Next Dan
            Next Ren
        Next Retu
                                        
                                        
        '�A�̏���(�I���ʒu) ---------------------------------------------
        For Retu = CInt(Text(1).Text) To CInt(Text(2).Text)
            
            For Ren = CInt(OLD_Ren_End) To CInt(Text(4).Text) + 1 Step -1
                
                For Dan = CInt(OLD_Dan_Start) To CInt(OLD_Dan_End)
            
            
                    Call UniCode_Conv(K0_TANA.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Format(Retu, "00"))
                    Call UniCode_Conv(K0_TANA.Ren, Format(Ren, "00"))
                    Call UniCode_Conv(K0_TANA.Dan, Format(Dan, "00"))
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                Upd_com = BtOpInsert
                                Exit Do
                                    '����͖���
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
            
                    If Upd_com = BtOpUpdate Then
                                    
                                    
'                        If Zaiko_Check_Proc(Zaiko_Flg) Then
'                            Unload Me
'                        End If
                                    '�I�f�[�^�폜
                        Do
                            sts = BTRV(BtOpDelete, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                    If ans = vbCancel Then
                                        Call Input_UnLock
                                        Exit Function
                                    End If
                                Case Else
                                    Call Input_UnLock
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                    Exit Function
                            End Select
                        
                        Loop
                    
                    End If
            
            
            
                Next Dan
            Next Ren
        Next Retu
                                        
                                        
                                        
        '�i�̏���(�J�n�ʒu) ---------------------------------------------
        For Retu = CInt(Text(1).Text) To CInt(Text(2).Text)
            For Ren = CInt(Text(3).Text) To CInt(Text(4).Text)
                
                For Dan = CInt(OLD_Dan_Start) To CInt(Text(5).Text) - 1
            
            
                    Call UniCode_Conv(K0_TANA.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Format(Retu, "00"))
                    Call UniCode_Conv(K0_TANA.Ren, Format(Ren, "00"))
                    Call UniCode_Conv(K0_TANA.Dan, Format(Dan, "00"))
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                Upd_com = BtOpInsert
                                Exit Do
                                    '����͖���
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
                
                
            
                    If Upd_com = BtOpUpdate Then
                                    
                                    
'                        If Zaiko_Check_Proc(Zaiko_Flg) Then
'                            Unload Me
'                        End If
                                    '�I�f�[�^�폜
                        Do
                            sts = BTRV(BtOpDelete, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                    If ans = vbCancel Then
                                        Call Input_UnLock
                                        Exit Function
                                    End If
                                Case Else
                                    Call Input_UnLock
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                    Exit Function
                            End Select
                        
                        Loop
                    
                    End If
            
            
            
                
                
                Next Dan
            Next Ren
        Next Retu
                                        
        '�i�̏���(�I���ʒu) ---------------------------------------------
        For Retu = CInt(Text(1).Text) To CInt(Text(2).Text)
            For Ren = CInt(Text(3).Text) To CInt(Text(4).Text)
                
                For Dan = CInt(OLD_Dan_End) To CInt(Text(6).Text) + 1 Step -1
            
            
                    Call UniCode_Conv(K0_TANA.Soko_No, Combo(0).Text)
                    Call UniCode_Conv(K0_TANA.Retu, Format(Retu, "00"))
                    Call UniCode_Conv(K0_TANA.Ren, Format(Ren, "00"))
                    Call UniCode_Conv(K0_TANA.Dan, Format(Dan, "00"))
                    Do
                        sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                        Select Case sts
                            Case BtNoErr
                                Upd_com = BtOpUpdate
                                Exit Do
                            Case BtErrKeyNotFound
                                Upd_com = BtOpInsert
                                Exit Do
                                    '����͖���
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Call Input_UnLock
                                    Exit Function
                                End If
                            Case Else
                                Call Input_UnLock
                                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                Exit Function
                        End Select
                    Loop
                    If Upd_com = BtOpUpdate Then
                                    
                                    
'                        If Zaiko_Check_Proc(Zaiko_Flg) Then
'                            Unload Me
'                        End If
                                    '�I�f�[�^�폜
                        Do
                            sts = BTRV(BtOpDelete, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                    If ans = vbCancel Then
                                        Call Input_UnLock
                                        Exit Function
                                    End If
                                Case Else
                                    Call Input_UnLock
                                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                    Exit Function
                            End Select
                        
                        Loop
                    
                    End If
            
            
            
            
                Next Dan
            Next Ren
        Next Retu
                                        
                                        
                                        
    End If
                                        
                                        

 '   If Zaiko_Flg Then
 '
 '       MsgBox "�폜���ꂽ�I�ɍ݌ɂ����݂��܂����B���z�q��[" & To_Ido_Soko & "]���m�F���Ă��������B"
 '
 '
 '   End If

                                        
    Call Input_UnLock
                                        
                                        
                                        
                                        '���X�g�{�b�N�X�ǉ�
    If com = BtOpInsert Then
        Combo(0).AddItem Combo(0).Text
    End If
                                        '��ʃN���A�[
    Call Clear_Field(0)

    Update_Proc = False


End Function
Private Function Delete_Proc() As Integer
                            '�q�Ƀ}�X�^�̍폜
Dim sts     As Integer
Dim ans     As Integer
Dim flg     As Boolean
Dim i       As Integer
Dim com     As Integer

    Delete_Proc = True

    Call Input_Lock


    Call UniCode_Conv(K0_SOKO.Soko_No, Combo(0).Text)
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                flg = True
                Exit Do
            Case BtErrKeyNotFound
                flg = False
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SOKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Call Input_UnLock
                    Call Clear_Field(0)
                    Exit Function
                End If
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�q�Ƀ}�X�^")
                Exit Function
        End Select
    Loop

    If flg Then
        
        Call UniCode_Conv(K0_TANA.Soko_No, StrConv(SOKOREC.Soko_No, vbUnicode))
        Call UniCode_Conv(K0_TANA.Retu, "")
        Call UniCode_Conv(K0_TANA.Ren, "")
        Call UniCode_Conv(K0_TANA.Dan, "")
                
        com = BtOpGetGreater
                
        Do
            DoEvents
            Do
                sts = BTRV(com + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                Select Case sts
                    Case BtNoErr
                        
                        If StrConv(TANAREC.Soko_No, vbUnicode) <> StrConv(SOKOREC.Soko_No, vbUnicode) Then
                            sts = BtErrEOF
                        End If
                        
                        Exit Do
                    Case BtErrEOF
                        Exit Do
                            '����͖���
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Call Input_UnLock
                            Exit Function
                        End If
                    Case Else
                        Call Input_UnLock
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                        Exit Function
                End Select
            Loop
                    
            If sts = BtErrEOF Then
                Exit Do
            End If
                                    
                                    
'           If Zaiko_Check_Proc(Zaiko_Flg) Then
'               Unload Me
'           End If
                                    '�I�f�[�^�폜
                        
                        
            Do
                sts = BTRV(BtOpDelete, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Call Input_UnLock
                            Exit Function
                        End If
                    Case Else
                        Call Input_UnLock
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                        Exit Function
                End Select
            
            Loop
        
        Loop
        
        
        
        
        
        Do
            sts = BTRV(BtOpDelete, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<SOKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        sts = BTRV(BtOpUnlock, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                        Call Input_UnLock
                        Call Clear_Field(0)
                        Exit Function
                    End If
                Case Else
                    Call Input_UnLock
                    Call File_Error(sts, BtOpDelete, "�q�Ƀ}�X�^")
                    Exit Function
            End Select
        Loop
    End If
                                                    
                                        
                                        
                                        
                                        '���X�g�{�b�N�X�폜
    For i = 0 To Combo(0).ListCount - 1
        If Combo(0).Text = Combo(0).List(i) Then
            Combo(0).RemoveItem i
            Exit For
        End If
    Next i
                                        '��ʃN���A�[
    Call Clear_Field(0)

    Call Input_UnLock

    Delete_Proc = False


End Function

Private Sub Combo_DblClick(Index As Integer)

    If (Index = 0) Then
        If Item_Dsp() Then
            Unload Me
        End If
                
        Text(0).SetFocus
    End If

End Sub

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim wk As String

    Select Case KeyCode
        Case vbKeyReturn
            Select Case Index
                Case 0
                    wk = Trim(Combo(Index).Text)
                    If Len(wk) <> 2 Then
                        Beep
                        MsgBox "�q�ɇ��͂Q�����͂ł��B"
                        Combo(Index).SetFocus
                        Exit Sub
                    End If
                    If Item_Dsp() Then
                        Unload Me
                    End If
                        
                Case 1
                    If Combo(Index).Text = SOKO_BUN1$ Then
                        Combo(2).Text = KONS_KBN0$
'                        For i = 0 To UBound(JGYOBU_T)
'                            If RTrim(JGYOBU_T(i).CODE) = "0" Then
'                                Combo(3).Text = RTrim(JGYOBU_T(i).NAME)
'                                Exit For
'                            End If
'                        Next i
                        Combo(3).Text = "�i�Ȃ��j"
                        Combo(4).Text = NAIGAI0$
                        Combo(2).Enabled = False
                        Combo(3).Enabled = False
                        Combo(4).Enabled = False
                        For i = 1 To 6
                            Text(i).Text = "01"
                            Text(i).Enabled = False
                        Next i
                    
                    
                        Text(7).Text = "000"
                        Text(7).Enabled = False
                        
                        Frame1.Enabled = True
                    Else
                        Combo(2).Enabled = True
                        For i = 1 To 6
                            Text(i).Enabled = True
                        Next i
                    
                        Text(7).Enabled = True
                        Frame1.Enabled = False
                    End If
                
                Case 2
                
                    If Combo(Index).Text = KONS_KBN0$ Then
'                        For i = 0 To UBound(JGYOBU_T)
'                            If RTrim(JGYOBU_T(i).CODE) = "0" Then
'                                Combo(3).Text = RTrim(JGYOBU_T(i).NAME)
'                                Exit For
'                            End If
'                        Next i
                        Combo(3).Text = "�i�Ȃ��j"
                        Combo(4).Text = NAIGAI0$
                        Combo(3).Enabled = False
                        Combo(4).Enabled = False
                    Else
                        Combo(3).Enabled = True
                        Combo(4).Enabled = True
                
                    End If
            End Select
            If Index = 4 Then
                Text(0).SetFocus
            Else
                For i = Index + 1 To 4
                    If Combo(i).Enabled Then
                        Combo(i).SetFocus
                        Exit For
                    End If
                Next i
                If i > 4 Then
                    Text(0).SetFocus
                End If
            End If
    End Select
End Sub


Private Sub Combo_LostFocus(Index As Integer)

'    If (Index = 0) Then
'        If Item_Dsp() Then
'            Unload Me
'        End If
                
'        Text(0).SelStart = ZERO
'        Text(0).SelLength = Len(RTrim(Text(0).Text))
'        Text(0).SetFocus
'    End If


End Sub

Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case Index
        Case 0
                                            '�G���[�`�F�b�N
            If Err_Chk() Then
                Exit Sub
            End If
            
            Beep
            yn = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
            Combo(0).SetFocus
        Case 3
            Beep
            yn = MsgBox("�폜���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                If Delete_Proc() Then
                   Unload Me
                End If
            End If
            Combo(0).SetFocus
        Case 8
            Beep
            yn = MsgBox("�f�[�^�o�͂��܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                If Data_Proc() Then
                    Unload Me
                End If
            End If
        Case 11
            Unload Me
        Case Else
            Beep
    End Select
End Sub


Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i           As Integer
Dim c           As String * 128
Dim sts         As Integer
Dim sBuffer     As String * 255
Dim com         As String

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
    
    Text_Max = 7                '��ʍ��ڕʍő���ޯ��
    Combo_Max = 5
    Command_Max = 11

    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                
                                
                                
                                '�ړ���q�ɔԍ���荞��
    If GetIni(StrConv(App.EXEName, vbProperCase), "IDO_SOKO", "SYS", c) Then
        Beep
        MsgBox "�ړ���q�ɔԍ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    To_Ido_Soko = RTrim(c)
                                '�ړ��v����荞��
    If GetIni(StrConv(App.EXEName, vbProperCase), "YOIN", "SYS", c) Then
        Beep
        MsgBox "�ړ��v���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    To_Ido_Yoin = RTrim(c)


'�[���ԍ���荞��
                                'ܰ��ð��ݔԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    Ws_No = RTrim(com)
                                
                                
                                '�b�r�u�t�@�C������荞��
    If GetIni("FILE", "SOKO_CSV", "SYS", c) Then
        Beep
        MsgBox "�q�Ƀ}�X�^�f�[�^�o�͗p�t�@�C��[SOKO_CSV]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    Soko_Csv = Trim(c)
                                
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    Combo(3).AddItem "�i�Ȃ��j"
    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).Code = " " Then
            Exit For
        End If
        Combo(3).AddItem RTrim(JGYOBU_T(i).NAME)
    Next i

                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�I�}�X�^�n�o�d�m
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^(�_�~�[)�n�o�d�m
    If wITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�v���}�X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^(�_�~�[)�n�o�d�m
    If wZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����f�[�^�n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�q�ɕ��ސݒ�
    Combo(1).AddItem SOKO_BUN0$
    Combo(1).AddItem SOKO_BUN1$
                                '���ډ�
    Combo(2).AddItem KONS_KBN0$
    Combo(2).AddItem KONS_KBN1$
                                '�����O�ݒ�
    Combo(4).AddItem NAIGAI0$
    Combo(4).AddItem NAIGAI1$
    Combo(4).AddItem NAIGAI2$
                                '�g�p��
    Combo(5).AddItem KAHI_KBN0$
    Combo(5).AddItem KAHI_KBN1$
    
    If List_Proc() Then
        Unload Me
    End If
    Beep
    MsgBox "���̏����̓V�X�e���S�̂ɉe������̂ŏ\�����ӂ��đ��삵�Ă��������B"
                                '��ʏ����ݒ�
    Call Clear_Field(0)
    
    Combo(0).SetFocus
    
    End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^")
        End If
    End If
    
                                            '�I�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^(�_�~�[)�b�k�n�r�d
    sts = BTRV(BtOpClose, wITEM_POS, wITEMREC, Len(wITEMREC), K0_wITEM, Len(K0_wITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�݌Ƀf�[�^(�_�~�[)�b�k�n�r�d
    sts = BTRV(BtOpClose, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�݌Ɉړ����f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ����f�[�^")
        End If
    End If

    
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "�q�Ƀ}�X�^")
    End If
    Set F1010111 = Nothing

    End
End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If

End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim RetBuf As String
Dim i As Integer


    Select Case KeyCode
        Case vbKeyReturn
            If (Index > 0) Then
                If Not IsNumeric(Text(Index).Text) Then
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Exit Sub
                Else
                    Text(Index).Text = Format(CInt(Text(Index).Text), "00")
                End If
            End If
            If Index < Text_Max Then
                For i = Index + 1 To Text_Max
                    If Text(i).Enabled Then
                        Text(i).SetFocus
                        Exit For
                    End If
                Next i
            End If
    End Select
End Sub
Private Function Zaiko_Check_Proc(Zaiko_Flg) As Integer
'---------------------------------- �폜�ΏےI�̍݌�����
Dim sts         As Integer
Dim ans         As Integer


Dim JGYOBU      As String * 1
Dim NAIGAI      As String * 1
Dim HIN_GAI     As String * 13
Dim NYUKA_DT    As String * 8
Dim LOCATION    As String * 8
Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long

    Zaiko_Check_Proc = True


    
    Do
        
        DoEvents
        
        Call UniCode_Conv(K0_ZAIKO.Soko_No, StrConv(TANAREC.Soko_No, vbUnicode))
        Call UniCode_Conv(K0_ZAIKO.Retu, StrConv(TANAREC.Retu, vbUnicode))
        Call UniCode_Conv(K0_ZAIKO.Ren, StrConv(TANAREC.Ren, vbUnicode))
        Call UniCode_Conv(K0_ZAIKO.Dan, StrConv(TANAREC.Dan, vbUnicode))
        Call UniCode_Conv(K0_ZAIKO.JGYOBU, "")
        Call UniCode_Conv(K0_ZAIKO.NAIGAI, "")
        Call UniCode_Conv(K0_ZAIKO.HIN_GAI, "")
        Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")
        Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")
        
        
        sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetGreater, "�݌Ƀf�[�^")
                Exit Function
        End Select
    
        If StrConv(ZAIKOREC.Soko_No, vbUnicode) <> StrConv(TANAREC.Soko_No, vbUnicode) Or _
            StrConv(ZAIKOREC.Retu, vbUnicode) <> StrConv(TANAREC.Retu, vbUnicode) Or _
            StrConv(ZAIKOREC.Ren, vbUnicode) <> StrConv(TANAREC.Ren, vbUnicode) Or _
            StrConv(ZAIKOREC.Dan, vbUnicode) <> StrConv(TANAREC.Dan, vbUnicode) Then
            
            Exit Do
        End If
    
    
        If StrConv(ZAIKOREC.LOCK_F, vbUnicode) = LOCK_ON And _
            (Trim(StrConv(ZAIKOREC.WEL_ID, vbUnicode)) <> Ws_No Or _
            Trim(StrConv(ZAIKOREC.PRG_ID, vbUnicode)) <> App.EXEName) Then
            
            
            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
            If ans = vbCancel Then
                Exit Function
            End If
        Else
                                            '�g�����U�N�V�����J�n
            sts = BTRV(BtOpBeginConcurrentTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            If sts <> BtNoErr Then
                Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
                Exit Function
            End If
    
            JGYOBU = StrConv(ZAIKOREC.JGYOBU, vbUnicode)
            NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
            HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
            NYUKA_DT = StrConv(ZAIKOREC.NYUKA_DT, vbUnicode)
            LOCATION = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
            Sumi_Qty = 0
            Mi_Qty = 0
            If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_ON Then
                Sumi_Qty = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            Else
                Mi_Qty = CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
            End If

            sts = Zaiko_Lock_Proc(StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode), _
                            StrConv(ZAIKOREC.JGYOBU, vbUnicode), _
                            StrConv(ZAIKOREC.NAIGAI, vbUnicode), _
                            StrConv(ZAIKOREC.HIN_GAI, vbUnicode), _
                            Ws_No)
            Select Case sts
                Case False
                Case True, SYS_CANCEL
                    GoTo Abort_Tran
                Case SYS_ERR
                    GoTo Abort_Tran
            End Select
    


            sts = IDO_Update_Proc(JGYOBU, _
                                    NAIGAI, _
                                    HIN_GAI, _
                                    NYUKA_DT, _
                                    LOCATION, _
                                    (To_Ido_Soko & "01" & "01" & "01"), _
                                    To_Ido_Yoin, _
                                    Sumi_Qty, _
                                    Mi_Qty, _
                                    Ws_No, _
                                    Ws_No, , _
                                    "�I�����e�i���X")
            Select Case sts
                Case False
                Case Else
                    GoTo Abort_Tran
            End Select
    
    
    
            sts = BTRV(BtOpEndTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
            If sts <> BtNoErr Then
                Call File_Error(sts, BtOpEndTransaction, "")
                GoTo Abort_Tran
            End If
    
            Zaiko_Flg = True
    
        End If
    Loop




    Zaiko_Check_Proc = False
    Exit Function

Abort_Tran:
    
    sts = BTRV(BtOpAbortTransaction, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

End Function


