VERSION 5.00
Begin VB.Form F1200401 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�͈͓��ړ����݌Ɉꗗ�f�[�^�쐬"
   ClientHeight    =   6960
   ClientLeft      =   2025
   ClientTop       =   2550
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
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   11295
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   7800
      MaxLength       =   2
      TabIndex        =   6
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   7200
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   6360
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   5520
      MaxLength       =   2
      TabIndex        =   3
      Top             =   2520
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   4920
      MaxLength       =   2
      TabIndex        =   2
      Top             =   2520
      Width           =   375
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   4080
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   4080
      MaxLength       =   4
      TabIndex        =   1
      Top             =   2520
      Width           =   615
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
      Index           =   10
      Left            =   9480
      TabIndex        =   17
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
      Index           =   8
      Left            =   7800
      TabIndex        =   15
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
      Index           =   7
      Left            =   6480
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      Index           =   3
      Left            =   2640
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      Index           =   0
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   375
      Index           =   6
      Left            =   8160
      TabIndex        =   27
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   375
      Index           =   5
      Left            =   7560
      TabIndex        =   26
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�N"
      Height          =   375
      Index           =   4
      Left            =   6960
      TabIndex        =   25
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���`"
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   24
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   375
      Index           =   2
      Left            =   5280
      TabIndex        =   23
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�N"
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   22
      Top             =   2640
      Width           =   375
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
      Height          =   315
      Left            =   120
      TabIndex        =   21
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   255
      Index           =   33
      Left            =   2400
      TabIndex        =   20
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���t�͈�"
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   19
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1200401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxS_YY% = 0                  '�J�n�@�N
Private Const ptxS_MM% = 1                  '�J�n�@��
Private Const ptxS_DD% = 2                  '�J�n�@��

Private Const ptxE_YY% = 3                  '�I���@�N
Private Const ptxE_MM% = 4                  '�I���@��
Private Const ptxE_DD% = 5                  '�I���@��

Private Const Text_Max% = 5                 '��ʍ��ڕʍő���ޯ��

Private Const pcmbNAIGAI% = 0               '�����O

Dim PSTOCK_DATA    As String                '�͈͓��ړ����݌Ɉꗗ�f�[�^�t���p�X
Private Function OUTPUT_Proc() As Integer
'----------------------------------------------------------------------------
'                  �b�r�u�f�[�^�o�͏���
'----------------------------------------------------------------------------
    
Dim sts             As Integer
Dim com             As Integer
Dim Ret             As Integer
    

Dim FileNo          As Integer
Dim fileName        As String

Dim Save_Soko       As String * 2

Dim c               As String * 128
Dim Soko_No         As String * 2

    OUTPUT_Proc = True
'���s���̓C�x���g�擾�s��
    Call Input_Lock         '��ʍ��ڃ��b�N

    If Data_Make_Proc() Then
        Call Input_UnLock
        Exit Function
    End If


    FileNo = FreeFile
    fileName = PSTOCK_DATA
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Error_Proc
    
    Open (fileName) For Output As FileNo


    Write #FileNo, "�͈͓��ړ����݌Ɉꗗ"
    Write #FileNo, "�W���I��", "�i�ԁi�O���j", "�o�n�r�݌�", "���_�݌�", "�݌Ɂ{", "�݌Ɂ|"

    
    com = BtOpGetFirst

    Do
        DoEvents
        
        sts = BTRV(com, PSTOCK_POS, PSTOCKREC, Len(PSTOCKREC), K1_PSTOCK, Len(K1_PSTOCK), 1)

        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�͈͓��ړ����݌Ɉꗗ")
                Exit Function
        End Select
                                                '�W���I��
        
        If GetIni("SOKO_NO", Left(StrConv(PSTOCKREC.ST_Location, vbUnicode), 2), "SYS", c) Then
            Soko_No = Left(StrConv(PSTOCKREC.ST_Location, vbUnicode), 2)
        Else
            Soko_No = Trim(c)
        End If
        
        Write #FileNo, Soko_No & "-" & _
                Mid(StrConv(PSTOCKREC.ST_Location, vbUnicode), 3, 2) & "-" & _
                Mid(StrConv(PSTOCKREC.ST_Location, vbUnicode), 5, 2) & "-" & _
                Right(StrConv(PSTOCKREC.ST_Location, vbUnicode), 2),
                                                '�i�ԁi�O���j
        Write #FileNo, StrConv(PSTOCKREC.HIN_GAI, vbUnicode),
                                                '�o�n�r���݌�
        Write #FileNo, Format(CLng(StrConv(PSTOCKREC.T_Zai_Qty, vbUnicode)), "#0"),
                                                '���_�݌�
        Write #FileNo, Format(CLng(StrConv(PSTOCKREC.HS_ZAIQTY, vbUnicode)), "#0"),
                                                '���ԓ�����
        Write #FileNo, Format(CLng(StrConv(PSTOCKREC.Plus_QTY, vbUnicode)), "#0"),
                                                '���ԓ��o��
        Write #FileNo, Format(CLng(StrConv(PSTOCKREC.Minus_QTY, vbUnicode)), "#0")



        com = BtOpGetNext
    Loop

    Close #FileNo
    
    Call Input_UnLock         '��ʍ��ڃ��b�N����
    Beep
    MsgBox "�u" & fileName & "�v�͐���ɏo�͂���܂����B"

    OUTPUT_Proc = False
    
    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "���g�p���ł��B"
        Call Input_UnLock         '��ʍ��ڃ��b�N����
        OUTPUT_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        OUTPUT_Proc = True
    End If
End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1200401.MousePointer = vbHourglass

    Call Ctrl_Lock(F1200401)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1200401)


    F1200401.MousePointer = vbDefault

End Sub

Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   ���̓G���[�`�F�b�N����
'----------------------------------------------------------------------------
Dim i   As Integer
    
    Err_Chk = True

    For i = ptxS_YY To ptxE_DD
        
        If Trim(Text(i).Text) = "" Then
            If i = ptxS_YY Then
                Text(i).Text = "0000"
            End If
            If i = ptxS_MM Or i = ptxS_DD Then
                Text(i).Text = "00"
            End If
        
            If i = ptxE_YY Then
                Text(i).Text = "9999"
            End If
            If i = ptxE_MM Or i = ptxE_DD Then
                Text(i).Text = "99"
            End If
        End If
        
        If Not IsNumeric(Trim(Text(i).Text)) Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B"
            Text(i).SetFocus
            Exit Function
        Else
            
            If i <> ptxS_YY And i <> ptxE_YY Then
                Text(i).Text = Format(CInt(Trim(Text(i).Text)), "00")
            End If
        
        
        End If
    
    Next i
    
    If (Text(ptxS_YY).Text & Text(ptxS_MM).Text & Text(ptxS_DD).Text) > _
        (Text(ptxE_YY).Text & Text(ptxE_MM).Text & Text(ptxE_DD).Text) Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(ptxS_YY).SetFocus
        Exit Function
    End If
    
    Err_Chk = False

End Function

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �R���{�{�b�N�X���́i�j�����c�������j����
'----------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case pcmbNAIGAI        '�����敪
            Text(ptxS_YY).SetFocus
    End Select

End Sub

Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        Case 7                              '�f�[�^�o��
            If Err_Chk() Then
                Exit Sub
            End If
        
            Beep
            ans = MsgBox("�u�͈͓��ړ����݌Ɉꗗ�v�f�[�^�o�͂��܂����H", vbYesNo + vbQuestion, "�m�F����")
            
            If ans = vbYes Then
                
                
                If OUTPUT_Proc() Then
                    Unload Me
                End If
            End If
            Combo(pcmbNAIGAI).SetFocus
        Case 11                             '�I��
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
Dim i   As Integer
Dim c   As String * 128
Dim sts As Integer

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
    LOG_F = Trim(c)
                                '�͈͓��ړ����݌Ɉꗗ�t�@�C������荞��
    If GetIni("FILE", "PSTOCK_DATA", "SYS", c) Then
        Beep
        MsgBox "�͈͓��ړ����݌Ɉꗗ�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    PSTOCK_DATA = Trim(c)
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).Code = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).Code = Last_JGYOBU Then
            F1200401.Caption = "�͈͓��ړ����݌Ɉꗗ�f�[�^�쐬�i" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                
                                        '�u�����O�U�ւ��v�̗v��
    If GetIni("YOIN", "YOIN_FURIKAE", "SYS", c) Then
        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_FURIKAE] READ ERROR")
        MsgBox "�V�X�e���\��ϗv���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    YOIN_FURIKAE = Trim(c)
                                        '�u�����O�U�ւ��v�̗v��
    If GetIni("YOIN", "YOIN_FURIKAE_IN", "SYS", c) Then
        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_FURIKAE_IN] READ ERROR")
        MsgBox "�V�X�e���\��ϗv���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    YOIN_FURIKAE_IN = Trim(c)
                                        '�u�����O�U�ւ��v�̗v��
    If GetIni("YOIN", "YOIN_FURIKAE_OUT", "SYS", c) Then
        Call Log_Out(LOG_F, "[SYS.INI] [YOIN] [YOIN_FURIKAE_OUT] READ ERROR")
        MsgBox "�V�X�e���\��ϗv���̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    YOIN_FURIKAE_OUT = Trim(c)
                                
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌ɏW�v�f�[�^�n�o�d�m
    If SUMZ_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�͈͓��ړ����݌Ɉꗗ�f�[�^�n�o�d�m
    If PSTOCK_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Show
                                
                                '��ʏ����ݒ�
    Combo(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo(pcmbNAIGAI).ListIndex = 0
    
    Combo(pcmbNAIGAI).SetFocus
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�݌ɏW�v�f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌ɏW�v�f�[�^")
        End If
    End If
                                            '�͈͓��ړ����݌Ɉꗗ�\�b�k�n�r�d
    sts = BTRV(BtOpClose, PSTOCK_POS, PSTOCKREC, Len(PSTOCKREC), K0_PSTOCK, Len(K0_PSTOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�͈͓��ړ����݌Ɉꗗ�\")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1200401 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
    If JGYOBU_T(Index).Code = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).Code = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1200401.Caption = "�͈͓��ړ����݌Ɉꗗ�f�[�^�쐬�i" + RTrim(JGYOBU_T(Index).NAME) + ")"
    Last_JGYOBU = JGYOBU_T(Index).Code
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
Dim i   As Integer

    
    If KeyCode <> vbKeyReturn Then Exit Sub
                
        
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i

End Sub
Private Function Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                  �u�͈͓��ړ����݌Ɉꗗ�v�쐬����
'----------------------------------------------------------------------------
Dim sts                 As Integer
Dim com                 As Integer
Dim com_IDO             As Integer
Dim ans                 As Integer

Dim Sum_Plus            As Long
Dim Sum_Minus           As Long

Dim Sumi_QTY            As Long
Dim Mi_QTY              As Long


    Data_Make_Proc = True
'---------------------------------------------------------- '�S���R�[�h�폜
    com = BtOpGetFirst
    Do
        
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, PSTOCK_POS, PSTOCKREC, Len(PSTOCKREC), K0_PSTOCK, Len(K0_PSTOCK), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<PSTOCK.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�͈͓��ړ����݌Ɉꗗ")
                    Exit Function
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        Do
            
            sts = BTRV(BtOpDelete, PSTOCK_POS, PSTOCKREC, Len(PSTOCKREC), K0_PSTOCK, Len(K0_PSTOCK), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<PSTOCK.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "�͈͓��ړ����݌Ɉꗗ")
                    Exit Function
            End Select
        
        Loop
        
        com = BtOpGetNext
    
    Loop
    
'---------------------------------------------------------- '�W�v�f�[�^�쐬�J�n
    '�i�ڃ}�X�^�x�[�X�ŏ����J�n
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")
    
    com = BtOpGetGreaterEqual
    
    Do
        
        DoEvents
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                Exit Function
        End Select
        '�݌Ɉړ�������
        Call UniCode_Conv(K1_IDO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K1_IDO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K1_IDO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K1_IDO.JITU_DT, (Text(ptxS_YY).Text & Text(ptxS_MM).Text & Text(ptxS_DD).Text))
        Call UniCode_Conv(K1_IDO.JITU_TM, "")
    
        Sum_Plus = 0
        Sum_Minus = 0
    
        com_IDO = BtOpGetGreater
    
    
        Do
        
            DoEvents
            sts = BTRV(com_IDO, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)
            Select Case sts
                Case BtNoErr
                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                        StrConv(IDOREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                        Exit Do
                    End If
            
                    If Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                        Exit Do
                    End If
            
                    If StrConv(IDOREC.JITU_DT, vbUnicode) > (Text(ptxE_YY).Text & Text(ptxE_MM).Text & Text(ptxE_DD).Text) Then
                        Exit Do
                    End If
            
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com_IDO, "�݌Ɉړ���")
                    Exit Function
            End Select
        
            Select Case Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1)
                
                Case ACT_ZAITEI_IN
                    '�ݒ��{
                    Sum_Plus = Sum_Plus + (CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)))
                Case ACT_ZAITEI_OUT, ACT_SYUKA_KEI, ACT_SYUKA_HYO, ACT_SYUKA_GAI
                    '�ݒ��|�^�o��
                    Sum_Minus = Sum_Minus + (CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)))
                Case ACT_SYSTEM
                    '�V�X�e���\��
                    If StrConv(IDOREC.RIRK_ID, vbUnicode) = YOIN_FURIKAE_IN Then
                        '�����O�{
                        Sum_Plus = Sum_Plus + (CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)))
                    Else
                        If StrConv(IDOREC.RIRK_ID, vbUnicode) = YOIN_FURIKAE_OUT Then
                            '�����O�|
                            Sum_Minus = Sum_Minus + (CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)) + CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)))
                        End If
                    End If
            End Select
        
            com_IDO = BtOpGetNext
        
        Loop
    
    
        If Sum_Plus = 0 And Sum_Minus = 0 Then
        Else
            '�u�͈͓��ړ����݌Ɉꗗ�v�쐬
                                                                    '���ƕ�
            Call UniCode_Conv(PSTOCKREC.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                                                                    '�����O
            Call UniCode_Conv(PSTOCKREC.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                                                                    '�i�ԁi�O���j
            Call UniCode_Conv(PSTOCKREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                                    '�W���I��
            Call UniCode_Conv(PSTOCKREC.ST_Location, (StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                                        StrConv(ITEMREC.ST_DAN, vbUnicode)))
                                                                    '�o�n�r���݌�
            If Zaiko_Syukei_Proc(Sumi_QTY, Mi_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                Exit Function
            End If
            Call UniCode_Conv(PSTOCKREC.T_Zai_Qty, Format(Sumi_QTY + Mi_QTY, "00000000"))
                                                                    '���_�݌�
            Call UniCode_Conv(K0_SUMZ.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
            Call UniCode_Conv(K0_SUMZ.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_SUMZ.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        
            sts = BTRV(BtOpGetEqual, SUMZ_POS, SUMZREC, Len(SUMZREC), K0_SUMZ, Len(K0_SUMZ), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(SUMZREC.HS_ZAIQTY, "00000000")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�݌ɏW�v�f�[�^")
                    Exit Function
            End Select
        
            Call UniCode_Conv(PSTOCKREC.HS_ZAIQTY, StrConv(SUMZREC.HS_ZAIQTY, vbUnicode))
                                                                    '���ԓ�����
            Call UniCode_Conv(PSTOCKREC.Plus_QTY, Format(Sum_Plus, "00000000"))
                                                                    '���ԓ��o��
            Call UniCode_Conv(PSTOCKREC.Minus_QTY, Format(Sum_Minus, "00000000"))
        
            Do
                sts = BTRV(BtOpInsert, PSTOCK_POS, PSTOCKREC, Len(PSTOCKREC), K0_PSTOCK, Len(K0_PSTOCK), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<PSTOCK.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "�͈͓��ړ����݌Ɉꗗ")
                        Exit Function
                End Select
            Loop
        
        End If
    
        com = BtOpGetNext
    
    
    
    
    Loop
    
    Data_Make_Proc = False

End Function
