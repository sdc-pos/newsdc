VERSION 5.00
Begin VB.Form F1011051 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "�I�ԕʌ���������N��������"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
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
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4575
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3900
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "�����f�[�^(�݌v)"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   3
      Left            =   540
      TabIndex        =   6
      Top             =   2160
      Width           =   1920
   End
   Begin VB.Label L_CNT 
      Alignment       =   1  '�E����
      BorderStyle     =   1  '����
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   2
      Left            =   2640
      TabIndex        =   5
      Top             =   2100
      Width           =   1215
   End
   Begin VB.Label L_CNT 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   1
      Left            =   3120
      TabIndex        =   4
      Top             =   1620
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "�����������N"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   2
      Left            =   540
      TabIndex        =   3
      Top             =   1620
      Width           =   1440
   End
   Begin VB.Label L_CNT 
      Alignment       =   2  '��������
      BorderStyle     =   1  '����
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   0
      Left            =   3120
      TabIndex        =   2
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "������������"
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   1
      Left            =   540
      TabIndex        =   1
      Top             =   1080
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  '����
      Caption         =   "�I�ʌ����}�X�^�f�[�^������"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   3990
   End
End
Attribute VB_Name = "F1011051"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const plabPacking_No% = 0       '������
Private Const plabRank% = 1             '�����N
Private Const plaData_Create% = 2       '�f�[�^��������

Private Const RANK_CHR$ = "A-1,A-2,B-1,B-2,C-1,C-2,D,E"
Dim T_Rank              As Variant      '�����N�e�[�u��

Dim T_SokoChr(26)       As String       '�q�ɕ����ˑq�ɇ� �Ǒ�ð���
Dim wPUT_CNT            As Long         '�f�[�^��������(�\���p����)

Private Sub Command_Click(Index As Integer)

    If Data_Create_Proc Then
        MsgBox "�ُ픭���ׁ̈A�����I������܂����B", vbOKOnly
    End If

    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If

End Sub

Private Sub Form_Load()
Dim c       As String * 128
Dim sts     As Integer
Dim yn      As Integer


    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If

    '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)

                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�I�}�X�^�n�o�d�m
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If


    '���s�m�F
    yn = vbYes
    Beep
    yn = MsgBox("�I�ʌ����}�X�^�f�[�^�����A���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
    If yn = vbYes Then
        Show
        DoEvents
        Command(0).Value = True     '�����J�n
    Else
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
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
        Call File_Error(sts, BtOpReset, "")
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
    End If

    Set F1011051 = Nothing

    End

End Sub

Private Function Data_Create_Proc() As Integer
'========================================================================================
'                   �I�ʌ����}�X�^�f�[�^��������
'========================================================================================
Dim i           As Integer
Dim j           As Integer
Dim wCSV        As Variant
Dim wStr        As String
Dim FNo         As Integer
Dim c           As String * 128
Dim sts         As Integer


    Data_Create_Proc = True


    '�I�ʌ����}�X�^�n�o�d�m
    If TPACKING_ReCreate Then           '̧�ٍč쐬
        Exit Function
    End If

    If TPACKING_Open(BtOpenNomal) Then  'OPEN
        Exit Function
    End If


    '�q�ɔԍ��Ǒւ������̎擾(ini̧�ق��)
    For i = 1 To UBound(T_SokoChr)
        If GetIni("SOKO_NO", Format(i, "00"), "SYS", c) Then
        Else
            T_SokoChr(i) = RTrim(c)
        End If
    Next i


    '�����N�e�[�u��������
    T_Rank = Split(RANK_CHR, ",", -1, vbTextCompare)


    '�I�ʌ����}�X�^�ݒ�p�b�r�u�@�n�o�d�m
    If GetIni("FILE", "TPACKING_CSV", "SYS", c) Then
        Beep
        MsgBox "�I�ʌ����}�X�^�b�r�u�t�@�C�����̊l���Ɏ��s���܂����B"
        GoTo Data_Create_Proc_Exit
    End If
    wStr = RTrim(c)

    FNo = FreeFile
    Open wStr For Input As #FNo

    '�I�ʌ����}�X�^�ݒ�p�b�r�u�@�Ǎ���
    L_CNT(plaData_Create).Caption = ""              '�f�[�^�������� �N���A
    wPUT_CNT = 0

    Do While EOF(FNo) = False
        Line Input #FNo, wStr
        wCSV = Split(wStr, ",", -1, vbTextCompare)

        If Trim(wCSV(0)) = "No" Then
        Else
'        If Left(wCSV(0), 1) = "D" Then
            
            
            For i = 0 To UBound(T_Rank)
                j = i * 2 + 1
                If wCSV(j) <> "" And wCSV(j + 1) <> "" Then
                    '�I�ʌ����}�X�^�f�[�^������
                    If Data_Put_Proc(CStr(wCSV(0)), CStr(wCSV(j)), _
                                     CStr(wCSV(j + 1)), CStr(T_Rank(i))) Then
                        Close FNo
                        GoTo Data_Create_Proc_Exit
                    End If
                End If
            Next i
'        End If
        End If
    Loop

    '�I�ʌ����}�X�^�ݒ�p�b�r�u�@�b�k�n�r�d
    Close FNo


    Label1(0).Caption = ""
    DoEvents

    Data_Create_Proc = False
    Beep
    MsgBox "�f�[�^��������������ɏI�����܂����B"



Data_Create_Proc_Exit:

    '�I�ʌ����}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�ʌ����}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If


End Function

Private Function Data_Put_Proc(pPACKNo As String, pSRANGE As String, pERANGE As String, pRANK As String) As Integer
'========================================================================================
'                   �I�ʌ����}�X�^�f�[�^������
'========================================================================================
Dim sts         As Integer
Dim com         As Integer

Dim wSoko       As String       '�q�ɔԍ�
Dim wRetu_S     As String       '�� �J�n
Dim wRetu_E     As String       '�� �I��
Dim wRen_S      As String       '�A �J�n
Dim wRen_E      As String       '�A �I��
Dim i           As Integer



    Data_Put_Proc = True


    '�i���\��
    L_CNT(plabPacking_No).Caption = pPACKNo         '������ ������
    L_CNT(plabRank).Caption = pRANK                 '������ �����N


    '�q�ɔԍ��`�F�b�N
    wSoko = ""
    For i = 1 To UBound(T_SokoChr)
        If T_SokoChr(i) = Left(pSRANGE, 1) Then
            wSoko = Format(i, "00")
            Exit For
        End If
    Next i

    If wSoko = "" Then
        Beep
        MsgBox "�w�蕶���ɊY������q�ɂ�������܂���" & vbCrLf & vbCrLf & _
               "�������F" & pPACKNo & vbCrLf & _
               "�����N�@�F" & pRANK & vbCrLf & _
               "�͈͎w��F" & pSRANGE & "�`" & pERANGE
        Exit Function
    End If

    If Left(pSRANGE, 1) <> Left(pERANGE, 1) Then
        Beep
        MsgBox "�q�ɂ��ׂ��͈͎͂w��ł��܂���" & vbCrLf & vbCrLf & _
               "�������F" & pPACKNo & vbCrLf & _
               "�����N�@�F" & pRANK & vbCrLf & _
               "�͈͎w��F" & pSRANGE & "�`" & pERANGE
        Exit Function
    End If


    '�͈͊J�n�`�͈͏I���`�F�b�N
    If InStr(pSRANGE, "-") > 0 And InStr(pERANGE, "-") > 0 Then
    Else
        Beep
        MsgBox "�͈͎w�肪����������܂���" & vbCrLf & vbCrLf & _
               "�������F" & pPACKNo & vbCrLf & _
               "�����N�@�F" & pRANK & vbCrLf & _
               "�͈͎w��F" & pSRANGE & "�`" & pERANGE
        Exit Function
    End If

    i = InStr(pSRANGE, "-")
    wRetu_S = Mid(pSRANGE, 2, i - 2)    '�� �J�n
    wRen_S = Mid(pSRANGE, i + 1, 2)     '�A �J�n

    i = InStr(pERANGE, "-")
    wRetu_E = Mid(pERANGE, 2, i - 2)    '�� �I��
    wRen_E = Mid(pERANGE, i + 1, 2)     '�A �I��

    If IsNumeric(wRetu_S) And IsNumeric(wRen_S) And _
       IsNumeric(wRetu_E) And IsNumeric(wRen_E) Then
    Else
        Beep
        MsgBox "�͈͎w�肪����������܂���" & vbCrLf & vbCrLf & _
               "�������F" & pPACKNo & vbCrLf & _
               "�����N�@�F" & pRANK & vbCrLf & _
               "�͈͎w��F" & pSRANGE & "�`" & pERANGE
        Exit Function
    End If

    wRetu_S = Format(Val(wRetu_S), "00")    '�� �J�n
    wRen_S = Format(Val(wRen_S), "00")      '�A �J�n
    wRetu_E = Format(Val(wRetu_E), "00")    '�� �I��
    wRen_E = Format(Val(wRen_E), "00")      '�A �I��


    '�I�}�X�^���w��͈̗͂L���f�[�^����
    Call UniCode_Conv(K0_TANA.Soko_No, wSoko)
    Call UniCode_Conv(K0_TANA.Retu, wRetu_S)
    Call UniCode_Conv(K0_TANA.Ren, wRen_S)
    Call UniCode_Conv(K0_TANA.Dan, "")
    com = BtOpGetGreaterEqual
    Do
        sts = BTRV(com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(TANAREC.Soko_No, vbUnicode) <> wSoko Or _
                   StrConv(TANAREC.Retu, vbUnicode) > wRetu_E Or _
                   StrConv(TANAREC.Ren, vbUnicode) > wRen_E Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�I�}�X�^")
                Exit Function
        End Select

        '�g�p�\�I�̂ݏ����Ώ�
        If StrConv(TANAREC.KAHI_KBN, vbUnicode) = KAHI_KBN_OK Then

            '�I�ʌ����}�X�^�f�[�^������
            Call UniCode_Conv(TPACKINGREC.Soko_No, StrConv(TANAREC.Soko_No, vbUnicode))
            Call UniCode_Conv(TPACKINGREC.Retu, StrConv(TANAREC.Retu, vbUnicode))
            Call UniCode_Conv(TPACKINGREC.Ren, StrConv(TANAREC.Ren, vbUnicode))
            Call UniCode_Conv(TPACKINGREC.PACKING_NO, pPACKNo)
            Call UniCode_Conv(TPACKINGREC.RANK, pRANK)
            Call UniCode_Conv(TPACKINGREC.FILLER, "")
            sts = BTRV(BtOpInsert, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
            If sts <> BtNoErr Then
                Call File_Error(sts, com, "�I�ʌ����}�X�^")
                Exit Function
            End If

            '�I�ʌ����}�X�^�f�[�^���������\��
            wPUT_CNT = wPUT_CNT + 1
            L_CNT(plaData_Create).Caption = Format(wPUT_CNT, "#,0")
            DoEvents

        End If

        Call UniCode_Conv(K0_TANA.Soko_No, StrConv(TANAREC.Soko_No, vbUnicode))
        Call UniCode_Conv(K0_TANA.Retu, StrConv(TANAREC.Retu, vbUnicode))
        Call UniCode_Conv(K0_TANA.Ren, StrConv(TANAREC.Ren, vbUnicode))
        Call UniCode_Conv(K0_TANA.Dan, "99")
    Loop


    Data_Put_Proc = False


End Function
