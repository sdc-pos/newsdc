VERSION 5.00
Begin VB.Form F1010701 
   BackColor       =   &H00FFFFFF&
   Caption         =   "[��ƊǗ��}�X�^]�S���ҕʃ��j���[�o�^"
   ClientHeight    =   6300
   ClientLeft      =   2130
   ClientTop       =   2430
   ClientWidth     =   11280
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
   ScaleHeight     =   6300
   ScaleWidth      =   11280
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      ItemData        =   "F1010701.frx":0000
      Left            =   2760
      List            =   "F1010701.frx":0007
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   16
      Top             =   600
      Width           =   3375
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      ItemData        =   "F1010701.frx":0021
      Left            =   6240
      List            =   "F1010701.frx":0028
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   15
      Top             =   600
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   4140
      ItemData        =   "F1010701.frx":0042
      Left            =   2760
      List            =   "F1010701.frx":0044
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   6015
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
      Index           =   10
      Left            =   9480
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
      Index           =   9
      Left            =   8640
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�� ��"
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
      Index           =   7
      Left            =   6480
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
      Index           =   6
      Left            =   5640
      TabIndex        =   7
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "MENU"
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�� ��"
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
      TabIndex        =   2
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label lblALL_Sel 
      BorderStyle     =   1  '����
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   17
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "���j���["
      Height          =   240
      Index           =   0
      Left            =   6240
      TabIndex        =   14
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�S����"
      Height          =   240
      Index           =   2
      Left            =   2760
      TabIndex        =   13
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "F1010701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const pcmbTANTO% = 0
Private Const pcmbMENU% = 1

Private Const Command_Max% = 11

Private Const MENU_NON$ = "**"
Private Const MENU_NON_N$ = "�Ȃ��@�@�@�@�@�@�@�@"

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i As Integer

    F1010701.MousePointer = vbHourglass

    Call Ctrl_Lock(F1010701)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1010701)

    F1010701.MousePointer = vbDefault

End Sub
Private Function List_Proc()
'----------------------------------------------------------------------------
'                   �S���ҕʃ��j���[�\��
'----------------------------------------------------------------------------

Dim sts     As Integer
Dim com     As Integer
Dim ans     As Integer
Dim Edit    As String


    List_Proc = True
    
    
    List1.Clear
    
    com = BtOpGetFirst
    Do
        '�S���҃}�X�^�ǂݍ���
        sts = BTRV(com, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�S���҃}�X�^")
                Exit Function
        End Select
        '�S���ҕʃ��j���[�ǂݍ���
        Call UniCode_Conv(K0_TMENU.TANTO_CODE, StrConv(TANTOREC.TANTO_CODE, vbUnicode))
        sts = BTRV(BtOpGetEqual, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(TMENUREC.MENU_GRP_NO, MENU_NON)
                Call UniCode_Conv(MENUREC.MENU_GRP, MENU_NON_N)
            Case Else
                Call File_Error(sts, BtOpGetGreater, "�S���ҕʃ��j���[")
                Exit Function
        End Select
        '���j���[�Ǘ��ǂݍ���
        If Trim(StrConv(TMENUREC.MENU_GRP_NO, vbUnicode)) = MENU_NON Then
        Else
            Call UniCode_Conv(K0_MENU.MENU_GRP_NO, StrConv(TMENUREC.MENU_GRP_NO, vbUnicode))
            Call UniCode_Conv(K0_MENU.MENU_LV1, "")
            Call UniCode_Conv(K0_MENU.MENU_LV2, "")
            Call UniCode_Conv(K0_MENU.MENU_LV3, "")
            
            sts = BTRV(BtOpGetGreater, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(MENUREC.MENU_GRP_NO, vbUnicode) <> StrConv(TMENUREC.MENU_GRP_NO, vbUnicode) Then
                    
                        Call UniCode_Conv(TMENUREC.MENU_GRP_NO, MENU_NON)
                        Call UniCode_Conv(MENUREC.MENU_GRP, MENU_NON_N)
                    
                    End If
                Case BtErrEOF
                    Call UniCode_Conv(TMENUREC.MENU_GRP_NO, MENU_NON)
                    Call UniCode_Conv(MENUREC.MENU_GRP, MENU_NON_N)
                Case Else
                    Call File_Error(sts, com, "���j���[�Ǘ��}�X�^")
                    Exit Function
            End Select
        
        End If
    
        Edit = StrConv(TANTOREC.TANTO_CODE, vbUnicode) & " "
        Edit = Edit & StrConv(TANTOREC.TANTO_NAME, vbUnicode) & "   "
        Edit = Edit & StrConv(MENUREC.MENU_GRP, vbUnicode) & "     "
        Edit = Edit & StrConv(TMENUREC.MENU_GRP_NO, vbUnicode)
    
        List1.AddItem Edit
    
    
        com = BtOpGetNext
    Loop
    
    List_Proc = False
    
End Function
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �S���ҕʃ��j���[�̍X�V
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer
    
Dim Edit    As String
Dim i       As Integer

    
    Update_Proc = True
        
    Call Input_Lock
        
        
    Do                      '�S���폜��B�č\�z
        DoEvents
        Do
            sts = BTRV(BtOpGetFirst + BtSNoWait, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
            Select Case sts
                Case BtNoErr
                    Do
                        sts = BTRV(BtOpDelete, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANTOMENU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpDelete, "�S���ҕʃ��j���[")
                                Exit Function
                        End Select
                    Loop
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANTOMENU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�S���ҕʃ��j���[")
                    Exit Function
            End Select
        Loop
        
        If sts = BtErrEOF Then
            Exit Do
        End If
        
    Loop
    
    For i = 0 To List1.ListCount - 1
        DoEvents
        
        Edit = List1.List(i)
        
        If MENU_NON = Right(Edit, 2) Then
        Else
            
            Call UniCode_Conv(TMENUREC.TANTO_CODE, Trim(Left(Edit, 5)))
            Call UniCode_Conv(TMENUREC.MENU_GRP_NO, Trim(Right(Edit, 2)))
            Call UniCode_Conv(TMENUREC.FILLER, "")
    
            Do
                sts = BTRV(BtOpInsert, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANTOMENU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "�S���ҕʃ��j���[")
                        Exit Function
                End Select
            Loop
        End If
    Next i
        
    If List_Proc() Then
        Exit Function
    End If
    
    Call Input_UnLock
    
    Update_Proc = False
End Function

Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case Index
        Case 0
            Beep
            yn = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
        
        Case 1
            
            Call List_Update_Proc
        
        Case 4
        
            F1010702.Show vbModal
            If Form_RTN Then
                Unload Me
            End If
        
        Case 8
            sts = All_Tanto_Chk_Proc()
            Select Case sts
                Case False  '���ݖ���
                
                    Beep
                    yn = MsgBox("�S�S���ҋ��ʃ��j���[��o�^���܂����H�i�S���Ҍʂ͖����ɂȂ�܂��j", vbYesNo + vbQuestion, "�m�F����")
                    If yn = vbYes Then
                        If ALL_Update_Proc(0) Then
                            Unload Me
                        End If
                    End If
                
                Case True   '���ݗL��
                    Beep
                    yn = MsgBox("�S�S���ҋ��ʃ��j���[���폜���܂����H�i�S���Ҍʂ�o�^���ĉ������j", vbYesNo + vbQuestion, "�m�F����")
                    If yn = vbYes Then
                        If ALL_Update_Proc(1) Then
                            Unload Me
                        End If
                    End If
                
                Case SYS_ERR
                    Unload Me
            End Select
        Case 11
            Unload Me
    End Select

End Sub


Private Sub Form_Activate()

Dim com                 As Integer
Dim sts                 As Integer
Dim Edit                As String
Dim Sv_MENU_GRP_No      As String * 2
                                        
                                        
                                        
                                        '���j���[�ݒ�
    Combo(pcmbMENU).Clear
    Combo(pcmbMENU).AddItem MENU_NON_N & " " & MENU_NON
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���j���[�Ǘ��}�X�^")
                Unload Me
        End Select
        
        If com = BtOpGetFirst Then
            
            Edit = StrConv(MENUREC.MENU_GRP, vbUnicode) & " " & StrConv(MENUREC.MENU_GRP_NO, vbUnicode)
            Combo(pcmbMENU).AddItem Edit
            Sv_MENU_GRP_No = StrConv(MENUREC.MENU_GRP_NO, vbUnicode)
        
        End If
        
        
        If Sv_MENU_GRP_No <> StrConv(MENUREC.MENU_GRP_NO, vbUnicode) Then
            Edit = StrConv(MENUREC.MENU_GRP, vbUnicode) & " " & StrConv(MENUREC.MENU_GRP_NO, vbUnicode)
            Combo(pcmbMENU).AddItem Edit
            Sv_MENU_GRP_No = StrConv(MENUREC.MENU_GRP_NO, vbUnicode)
        End If
        
        com = BtOpGetNext
    
    Loop
    Combo(pcmbMENU).ListIndex = 0
    
    If List_Proc() Then
        Unload Me
    End If
        
    If List1.ListCount = 0 Then
        Combo(pcmbTANTO).SetFocus
    Else
        List1.ListIndex = 0
        List1.SetFocus
    End If

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
Dim com         As Integer

Dim Sv_MENU_GRP As String * 10

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B"
        End
    End If

    
    F1010702.cboJIGYOBU.Clear
    For i = 0 To UBound(JGYOBU_T)
        If Trim(JGYOBU_T(i).CODE) = "" Then
            Exit For
        End If
        F1010702.cboJIGYOBU.AddItem JGYOBU_T(i).NAME & " " & JGYOBU_T(i).CODE
    Next i
    F1010702.cboJIGYOBU.ListIndex = 0
    If F1010702.cboJIGYOBU.ListCount = 1 Then
        F1010702.cboJIGYOBU.Enabled = False
    End If
    
    
    '�����O���ݒ�
    i = 0
    Do
        i = i + 1
        If GetIni(App.EXEName, "NAIGAI" & Format(i, "0"), "SYS", c) Then
            Exit Do
        End If
        ReDim Preserve NAIGAI_CODE(i - 1)
        NAIGAI_CODE(i - 1) = Trim(c)
    
    Loop
    If i = 1 Then
        Beep
        MsgBox "�����O�̊l���Ɏ��s���܂����B"
        End
    End If
                                
                                
    F1010702.cboNAIGAI.Clear
    For i = 0 To UBound(NAIGAI_CODE)
        
        Select Case NAIGAI_CODE(i)
            Case NAIGAI_NAI
                F1010702.cboNAIGAI.AddItem NAIGAI1 & " " & NAIGAI_CODE(i)
        
            Case NAIGAI_GAI
                F1010702.cboNAIGAI.AddItem NAIGAI2 & " " & NAIGAI_CODE(i)
        End Select
                    
    Next i
    F1010702.cboNAIGAI.ListIndex = 0
    If F1010702.cboNAIGAI.ListCount = 1 Then
        F1010702.cboNAIGAI.Enabled = False
    End If
                                
                                
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���j���[�}�X�^�n�o�d�m
    If MENU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���j���[�}�X�^�i�ꎞ�j�n�o�d�m
    If tmpMENU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���ҕʃ��j���[�n�o�d�m
    If TMENU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�v���}�X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '������Ǘ��}�X�^�n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Show
    Form_RTN = False
    Load F1010702
    If Form_RTN Then
        Unload Me
    End If
                                        
                                        '���ʃ��j���[�L������
    Call UniCode_Conv(K0_TMENU.TANTO_CODE, ALL_TANTO_CODE)
    sts = BTRV(BtOpGetEqual, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
    Select Case sts
        
        Case BtNoErr
            lblALL_Sel.Caption = "���ʃ��j���[�w��"
        Case BtErrKeyNotFound
            lblALL_Sel.Caption = "�ʃ��j���[�w��"
        Case Else
            Call File_Error(sts, com, "���j���[�Ǘ��}�X�^")
            Unload Me
    
    End Select
                                        
                                        '�S���Ґݒ�
    Combo(pcmbTANTO).Clear
    com = BtOpGetFirst
    Do
        sts = BTRV(com, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�S���҃}�X�^")
                Unload Me
        End Select
        Combo(pcmbTANTO).AddItem StrConv(TANTOREC.TANTO_CODE, vbUnicode) & " " & StrConv(TANTOREC.TANTO_NAME, vbUnicode)
        com = BtOpGetNext
    Loop
    
    If Combo(pcmbTANTO).ListCount <> 0 Then
        Combo(pcmbTANTO).ListIndex = 0
    End If
                                        '���j���[�ݒ�
    Combo(pcmbMENU).Clear
    Combo(pcmbMENU).AddItem MENU_NON
    com = BtOpGetFirst
    Do
        sts = BTRV(com, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���j���[�Ǘ��}�X�^")
                Unload Me
        End Select
        If com = BtOpGetFirst Then
            Combo(pcmbMENU).AddItem StrConv(MENUREC.MENU_GRP, vbUnicode)
            Sv_MENU_GRP = Trim(StrConv(MENUREC.MENU_GRP, vbUnicode))
        End If
        
        
        If Trim(Sv_MENU_GRP) <> Trim(StrConv(MENUREC.MENU_GRP, vbUnicode)) Then
            Combo(pcmbMENU).AddItem StrConv(MENUREC.MENU_GRP, vbUnicode)
        End If
        
        Sv_MENU_GRP = Trim(StrConv(MENUREC.MENU_GRP, vbUnicode))
        com = BtOpGetNext
    
    Loop
    If Combo(pcmbMENU).ListCount <> 0 Then
        Combo(pcmbMENU).ListIndex = 0
    End If
    
    If List_Proc() Then
        Unload Me
    End If
        
    If List1.ListCount = 0 Then
        Combo(pcmbTANTO).SetFocus
    Else
        List1.ListIndex = 0
        List1.SetFocus
    End If


End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
                                            
                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^")
        End If
    End If
                                            '���j���[�Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MENU_POS, MENUREC, Len(MENUREC), K0_MENU, Len(K0_MENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���j���[�Ǘ��}�X�^")
        End If
    End If
                                            '���j���[�Ǘ��}�X�^�i�ꎞ�t�@�C���j�b�k�n�r�d
    sts = BTRV(BtOpClose, tmpMENU_POS, tmpMENUREC, Len(tmpMENUREC), K0_tmpMENU, Len(K0_tmpMENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���j���[�Ǘ��}�X�^�i�ꎞ�t�@�C���j")
        End If
    End If
                                            '������Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")
        End If
    End If
                                            '�S���ҕʃ��j���[�b�k�n�r�d
    sts = BTRV(BtOpClose, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���ҕʃ��j���[")
        End If
    End If
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "���j���[�Ǘ��}�X�^")
    End If
    Set F1010701 = Nothing
    Set F1010702 = Nothing
    End
End Sub

Private Sub List1_GotFocus()
    
    If List1.ListCount > 0 Then
        If List1.ListIndex <= 0 Then
            List1.ListIndex = 0
        End If
    End If

End Sub

Private Sub List_Update_Proc()
'----------------------------------------------------------------------------
'                   ���X�g�{�b�N�X�X�V
'----------------------------------------------------------------------------
Dim i       As Integer
Dim Edit    As String


    For i = 0 To List1.ListCount - 1
        
        If Trim(Left(Combo(pcmbTANTO).Text, 5)) = Trim(Left(List1.List(i), 5)) Then
            List1.RemoveItem i
        End If
    
    Next i
    
    Edit = Combo(pcmbTANTO).Text & "   "
    Edit = Edit & Combo(pcmbMENU).Text
         
         
    List1.AddItem Edit

End Sub

Private Function All_Tanto_Chk_Proc() As Integer
'----------------------------------------------------------------------------
'                   �S�S���ҋ��ʃ��j���[�쐬�^�J���`�F�b�N
'----------------------------------------------------------------------------

Dim sts     As Integer
    
    
    Call UniCode_Conv(K0_TMENU.TANTO_CODE, ALL_TANTO_CODE)
    sts = BTRV(BtOpGetEqual, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
    Select Case sts
        Case BtNoErr
            All_Tanto_Chk_Proc = True
        Case BtErrKeyNotFound
            All_Tanto_Chk_Proc = False
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�S���ҕʃ��j���[")
            All_Tanto_Chk_Proc = SYS_ERR
    End Select


    
End Function

Private Function ALL_Update_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   �S�S���ҋ��ʃ��j���[�o�^�^�폜
'                   0:�ǉ��@1:�폜
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim Rec_Flg As Boolean


    ALL_Update_Proc = True

    Call UniCode_Conv(K0_TMENU.TANTO_CODE, ALL_TANTO_CODE)
    
    Do
        
        sts = BTRV(BtOpGetEqual + BtSNoWait, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
        Select Case sts
            Case BtNoErr
                Rec_Flg = True
                Exit Do
            Case BtErrKeyNotFound
                Rec_Flg = False
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANTOMENU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���j���[�Ǘ��}�X�^")
                Exit Function
        End Select
    
    Loop

    Select Case Mode
        Case 0              '�ǉ�
            
            lblALL_Sel.Caption = "���ʃ��j���[�w��"
            
            Select Case Rec_Flg
                Case True
                
                    sts = BTRV(BtOpUnlock, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
                    If sts Then
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���j���[�Ǘ��}�X�^")
                        Exit Function
                    End If
                
                Case False
                    
                    Call UniCode_Conv(TMENUREC.TANTO_CODE, ALL_TANTO_CODE)  '�S���҃R�[�h
                    Call UniCode_Conv(TMENUREC.MENU_GRP_NO, "")             '���j���[�O���[�v
                    Call UniCode_Conv(TMENUREC.FILLER, "")
                    
                    Do
                        sts = BTRV(BtOpInsert, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANTOMENU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpInsert, "���j���[�Ǘ��}�X�^")
                                Exit Function
                        End Select
                    Loop
            
            End Select
        Case 1              '�폜
            
            lblALL_Sel.Caption = "�ʃ��j���[�w��"
            
            Select Case Rec_Flg
                Case True
                
                    Do
                        sts = BTRV(BtOpDelete, TMENU_POS, TMENUREC, Len(TMENUREC), K0_TMENU, Len(K0_TMENU), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANTOMENU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpDelete, "���j���[�Ǘ��}�X�^")
                                Exit Function
                        End Select
                    Loop
                
                
                Case False
                    
            End Select
    
    End Select


    ALL_Update_Proc = False

End Function
