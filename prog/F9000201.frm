VERSION 5.00
Begin VB.Form F9000201 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�I�f�[�^�o�^"
   ClientHeight    =   4935
   ClientLeft      =   2130
   ClientTop       =   2715
   ClientWidth     =   10935
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
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   10935
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.ListBox List1 
      Height          =   1740
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   22
      Top             =   1680
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   5040
      MaxLength       =   25
      TabIndex        =   20
      Top             =   2760
      Width           =   1092
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   5040
      MaxLength       =   25
      TabIndex        =   17
      Top             =   2160
      Width           =   1092
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Enabled         =   0   'False
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   5040
      MaxLength       =   25
      TabIndex        =   1
      Top             =   1680
      Width           =   1092
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   5040
      MaxLength       =   13
      TabIndex        =   0
      Top             =   1200
      Width           =   1092
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
      Left            =   9840
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4200
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
      Left            =   9000
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4200
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
      Left            =   8160
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4200
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
      Index           =   8
      Left            =   7320
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4200
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
      Left            =   6240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4200
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
      Left            =   5400
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4200
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
      Left            =   4560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4200
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
      Left            =   3720
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4200
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
      Index           =   3
      Left            =   2640
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4200
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4200
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�o �^"
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���I��"
      Height          =   252
      Index           =   2
      Left            =   4080
      TabIndex        =   21
      Top             =   2760
      Width           =   852
   End
   Begin VB.Label Label 
      Alignment       =   2  '��������
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ݒ�t�@�C���F"
      ForeColor       =   &H00FF0000&
      Height          =   252
      Index           =   1
      Left            =   2280
      TabIndex        =   19
      Top             =   480
      Width           =   6612
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�G���[����"
      Height          =   252
      Index           =   0
      Left            =   3720
      TabIndex        =   18
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�͌���"
      Height          =   252
      Index           =   15
      Left            =   3720
      TabIndex        =   16
      Top             =   1680
      Width           =   972
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   120
      TabIndex        =   15
      Top             =   4560
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���͌���"
      Height          =   252
      Index           =   14
      Left            =   3720
      TabIndex        =   14
      Top             =   1200
      Width           =   972
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F9000201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const In_File = "c:\SDC_fukuroi\WORK\TANA.CSV"

Dim W01 As String       '�I��
Dim W02 As String       '�z�X�g�I�ԁ^���l�i�_�~�[�j

Dim OK_DATE As String
Private Const ptxMAX% = 2
Private Const ptxIN% = 0
Private Const ptxOUT% = 1
Private Const ptxERR% = 2

Dim W_Cnt_in As Long
Dim W_Cnt_Out As Long
Dim W_Cnt_Err As Long
                                    '��ʏ�����Ԃ�ݒ肷��
Private Sub Clear_Field()
Dim i As Integer
    
    For i = 0 To ptxMAX
        Text(i).Text = ""
    Next i
    
End Sub
                                            '�I�̒ǉ�
Private Function Out_Proc() As Integer
Dim sts As Integer
Dim com As Integer
Dim yn As Integer
Dim Qty As Long
Dim W_No As String

    Out_Proc = False
    W_Cnt_in = 0
    W_Cnt_Out = 0
    W_Cnt_Err = 0
    On Error Resume Next
            '���ږ�Rec�@Dummy�ǂ�
    Input #1, W01, W02
    Do While Not EOF(1)
        Input #1, W01, W02
        If W01 = "" Then Exit Do
        W_Cnt_in = W_Cnt_in + 1
        Text(0) = W_Cnt_in
        DoEvents
        
        If TANA_OUT Then
            Beep
            Out_Proc = SYS_ERR
            Exit Function
        End If
    Loop
    If SOKO_UPDT Then
        Exit Function
    End If
    
    Out_Proc = False
End Function
Private Function SOKO_UPDT()
Dim com         As Integer
Dim yn          As Integer
Dim sts         As Integer
Dim W_Soko      As String
Dim W_Retu_S    As String
Dim W_Retu_E    As String
Dim W_Ren_S     As String
Dim W_Ren_E     As String
Dim W_Dan_S     As String
Dim W_Dan_E     As String
    SOKO_UPDT = True
    W_Soko = ""
    W_Cnt_in = 0
    com = BtOpGetFirst
    Do
        Do
            sts = BTRV(com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound, BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    yn = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If yn = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�I�f�[�^")
                    TANA_OUT = SYS_ERR
                    Exit Function
            End Select
        Loop
        If sts <> BtNoErr Then Exit Do
        W_Cnt_in = W_Cnt_in + 1
        Text(3) = W_Cnt_in
        DoEvents
        If W_Soko = StrConv(TANAREC.Soko_No, vbUnicode) Then
            W_Retu_E = StrConv(TANAREC.Retu, vbUnicode)
            W_Ren_E = StrConv(TANAREC.Ren, vbUnicode)
            W_Dan_E = StrConv(TANAREC.Dan, vbUnicode)
        Else
            If W_Soko <> "" Then
                
                Call UniCode_Conv(K0_SOKO.Soko_No, W_Soko)
                
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
                        Call UniCode_Conv(SOKOREC.RETU_START, W_Retu_S)
                        Call UniCode_Conv(SOKOREC.RETU_END, W_Retu_E)
                        Call UniCode_Conv(SOKOREC.REN_START, W_Ren_S)
                        Call UniCode_Conv(SOKOREC.REN_END, W_Ren_E)
                        Call UniCode_Conv(SOKOREC.DAN_START, W_Dan_S)
                        Call UniCode_Conv(SOKOREC.DAN_END, W_Dan_E)
                        sts = BTRV(BtOpUpdate, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                        If sts <> BtNoErr Then
                            MsgBox "�q�ɍX�V�G���[�I STS=" & sts, vbExclamation
                            Exit Function
                        End If
                    Case BtErrKeyNotFound, BtErrEOF
                        MsgBox "�q�ɖ����G���[�I SOKO=<" & W_Soko & ">", vbExclamation
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�q��Ͻ�")
                        Exit Function
                End Select
                
                
                
                W_Soko = StrConv(TANAREC.Soko_No, vbUnicode)
                W_Retu_S = StrConv(TANAREC.Retu, vbUnicode)
                W_Ren_S = StrConv(TANAREC.Ren, vbUnicode)
                W_Dan_S = StrConv(TANAREC.Dan, vbUnicode)
                W_Retu_E = StrConv(TANAREC.Retu, vbUnicode)
                W_Ren_E = StrConv(TANAREC.Ren, vbUnicode)
                W_Dan_E = StrConv(TANAREC.Dan, vbUnicode)
            End If
        End If
        
        com = BtOpGetNext
    Loop
    If W_Soko <> "" Then
        Call UniCode_Conv(K0_SOKO.Soko_No, W_Soko)
                
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                If W_Retu_E > StrConv(SOKOREC.RETU_END, vbUnicode) Then
                    Call UniCode_Conv(SOKOREC.RETU_END, W_Retu_E)
                End If
    '            Call UniCode_Conv(SOKOREC.REN_START, W_Ren_S)
                If W_Ren_E > StrConv(SOKOREC.REN_END, vbUnicode) Then
                    Call UniCode_Conv(SOKOREC.REN_END, W_Ren_E)
                End If
    '            Call UniCode_Conv(SOKOREC.DAN_START, W_Dan_S)
                If W_Dan_E > StrConv(SOKOREC.DAN_END, vbUnicode) Then
                    Call UniCode_Conv(SOKOREC.DAN_END, W_Dan_E)
                End If
                sts = BTRV(BtOpUpdate, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                If sts <> BtNoErr Then
                    MsgBox "�q�ɍX�V�G���[�I STS=" & sts, vbExclamation
                    Exit Function
                End If
            Case BtErrKeyNotFound, BtErrEOF
                MsgBox "�q�ɖ����G���[�I SOKO=<" & W_Soko & ">", vbExclamation
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�q��Ͻ�")
                Exit Function
        End Select
                
                        
    End If
    
    SOKO_UPDT = False
End Function

Private Function Err_Chk()
Dim com     As Integer
Dim yn      As Integer
Dim sts     As Integer
Dim W_Soko  As String
Dim W_Retu  As String
Dim W_Ren   As String
Dim W_Dan   As String
Dim W_SOKO_ERR As String

    Err_Chk = True
    If Len(W01) <> 8 Then
        Exit Function
    End If
    W_Soko = Mid(W01, 1, 2)
    W_Retu = Mid(W01, 3, 2)
    W_Ren = Mid(W01, 5, 2)
    W_Dan = Mid(W01, 7, 2)
                                                '�I�f�[�^�ҏW
    Call UniCode_Conv(K0_TANA.Soko_No, W_Soko)                  '�q��
    Call UniCode_Conv(K0_TANA.Retu, W_Retu)                     '��
    Call UniCode_Conv(K0_TANA.Ren, W_Ren)                       '�A
    Call UniCode_Conv(K0_TANA.Dan, W_Dan)                       '�i
    
    Do
        sts = BTRV(BtOpGetEqual, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                yn = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If yn = vbCancel Then
                    W_Cnt_Err = W_Cnt_Err + 1
                    Text(2) = W_Cnt_Err
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�I�f�[�^")
                TANA_OUT = SYS_ERR
                Exit Function
        End Select
    Loop
    If sts = BtNoErr Then
        W_Cnt_Err = W_Cnt_Err + 1
Debug.Print W_Soko & W_Retu & W_Ren & W_Dan
        Text(2) = W_Cnt_Err
        Exit Function
    End If
    
    Call UniCode_Conv(K0_SOKO.Soko_No, W_Soko)
    Do
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound, BtErrEOF
                
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE     '���R�[�h�g�p��
                Beep
                yn = MsgBox("���[���Ńf�[�^�g�p���ł��B<SOKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If yn = vbCancel Then Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                Beep
                MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B<SOKO>"
                Exit Function
        End Select
    Loop
    If sts <> BtNoErr Then
        If W_Soko <> W_SOKO_ERR Then List1.AddItem W_Soko
        W_Cnt_Err = W_Cnt_Err + 1
        Text(2) = W_Cnt_Err
        Exit Function
        W_SOKO_ERR = W_Soko
    End If
    
    Err_Chk = False
    
End Function

Private Function TANA_OUT()
Dim com     As Integer
Dim yn      As Integer
Dim sts     As Integer
Dim W_Soko  As String
Dim W_Retu  As String
Dim W_Ren   As String
Dim W_Dan   As String

    TANA_OUT = True
    If Err_Chk Then
        TANA_OUT = False
        Exit Function
    End If
    
    W_Soko = Mid(W01, 1, 2)
    W_Retu = Mid(W01, 3, 2)
    W_Ren = Mid(W01, 5, 2)
    W_Dan = Mid(W01, 7, 2)
                                                '�I�f�[�^�ҏW
    Call UniCode_Conv(TANAREC.Soko_No, W_Soko)                  '�q��
    Call UniCode_Conv(TANAREC.Retu, W_Retu)                     '��
    Call UniCode_Conv(TANAREC.Ren, W_Ren)                       '�A
    Call UniCode_Conv(TANAREC.Dan, W_Dan)                       '�i
    Call UniCode_Conv(TANAREC.TANA_COND, "0")                   '�I���
    Call UniCode_Conv(TANAREC.KAHI_KBN, "0")                    '�g�p��
    
    Call UniCode_Conv(TANAREC.ZAIKO_SHOGO_FLG, "0")             '�݌ɏƍ� "0":�ƍ��Ώ� 2006.03.20
    
    
    Call UniCode_Conv(TANAREC.FILLER, "")
    
    Do
        sts = BTRV(BtOpInsert, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                yn = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If yn = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpInsert, "�I�f�[�^")
                TANA_OUT = SYS_ERR
                Exit Function
        End Select
    Loop
    
    W_Cnt_Out = W_Cnt_Out + 1
    Text(1) = W_Cnt_Out
    
    TANA_OUT = False
End Function
Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case Index
        Case 0                       '�X�V
            yn = MsgBox("�o�^���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbNo Then
                Command(11).SetFocus
                Exit Sub
            End If
            OK_DATE = Date
            sts = Out_Proc()
            Select Case sts
                Case False, True
                    
                Case SYS_ERR
                    Beep
                    MsgBox "�V�X�e���G���[�����I", vbExclamation
                    Unload Me
            End Select
            
            
            MsgBox "�o�^�I���I"
            Command(11).SetFocus
            Exit Sub
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
Dim i As Integer
Dim c As String * 128
Dim sts As Integer
    
    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
    
    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                '�q�Ƀf�[�^�t�@�C���n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�I�f�[�^�t�@�C���n�o�d�m
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Label(1) = Trim(Label(1).Caption) & In_File
    List1.Clear
    On Error GoTo Err_Exit
            '�ݒ�f�[�^�n�o�d�m(TXT)
    Open In_File For Input As #1 Len = 512
    Command(0).SetFocus
    Exit Sub
Err_Exit:
    MsgBox "�ݒ�p�f�[�^�t�@�C���L��܂���I", vbExclamation
    Unload Me
    End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

                                            '�q�Ƀf�[�^�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�q�Ƀf�[�^�t�@�C��")
        End If
    End If
                                            '�I�f�[�^�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�f�[�^�t�@�C��")
        End If
    End If
    
    sts = BTRV(BtOpReset, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "�I�f�[�^�t�@�C��")
    End If
    Close #1
    Set F9000201 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F9000201.Caption = "�݌ɓo�^�i" + RTrim(JGYOBU_T(Index).NAME) + "�j"
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

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
Dim sts As Integer
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Text(0).SetFocus
    
End Sub
