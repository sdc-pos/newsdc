VERSION 5.00
Begin VB.Form F1010801 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�S���҃}�X�^�����e�i���X"
   ClientHeight    =   11955
   ClientLeft      =   2130
   ClientTop       =   2430
   ClientWidth     =   17055
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
   ScaleHeight     =   11955
   ScaleWidth      =   17055
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   6840
      MaxLength       =   2
      TabIndex        =   20
      Top             =   960
      Width           =   372
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   18
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   4  '�S�p�Ђ炪��
      Index           =   1
      Left            =   3240
      MaxLength       =   10
      TabIndex        =   15
      Top             =   960
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Appearance      =   0  '�ׯ�
      Height          =   9150
      ItemData        =   "F1010801.frx":0000
      Left            =   2160
      List            =   "F1010801.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   5415
   End
   Begin VB.TextBox Text 
      Appearance      =   0  '�ׯ�
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   0
      Top             =   960
      Width           =   735
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   11160
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   11160
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   11160
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   11160
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   11160
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   11160
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   11160
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   11160
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   11160
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
      Top             =   11160
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
      Top             =   11160
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   11160
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "(�󔒁F�o�Η\��\���ΏۊO�@�ȊO�F�Ώ�)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6840
      TabIndex        =   21
      Top             =   1440
      Width           =   3732
   End
   Begin VB.Label Label 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "�敪"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   6840
      TabIndex        =   19
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   6000
      TabIndex        =   17
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "�S���Җ���"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   3240
      TabIndex        =   16
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label Label 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "�S����"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   2160
      TabIndex        =   14
      Top             =   720
      Width           =   720
   End
End
Attribute VB_Name = "F1010801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxTANTO_CODE% = 0
Private Const ptxTANTO_NAME% = 1
Private Const ptxPOST_CODE% = 2
'2011.09.06
Private Const ptxKUBUN% = 3



Private Const Text_Max% = 3                     '��ʍ��ڕʍő���ޯ��
Private Const Command_Max% = 11

Private TANTO_CSV As String

'Private Const LAST_UPDATE_DAY$ = "[F101080] 2011.09.30 10:00 [���i�����ёΉ�]" 2011.09.06
Private Const LAST_UPDATE_DAY$ = "[F101080] 2019.06.25 11:15"  '2019.06.25 ��ʃT�C�Y�g��



Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i As Integer

    F1010801.MousePointer = vbHourglass

    Call Ctrl_Lock(F1010801)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1010801)

    F1010801.MousePointer = vbDefault

End Sub
Private Function List_Proc()
'----------------------------------------------------------------------------
'                   �S���҃}�X�^�\��
'----------------------------------------------------------------------------

Dim sts     As Integer
Dim com     As Integer
Dim yn      As Integer
Dim Edit    As String


    List_Proc = True
    
    
    List1.Clear
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�S���҃}�X�^")
                Exit Function
        End Select
        
        Edit = StrConv(TANTOREC.TANTO_CODE, vbUnicode) & "    "
        Edit = Edit & StrConv(TANTOREC.TANTO_NAME, vbUnicode) & "   "
        
        '2011.09.06
'        Edit = Edit & StrConv(TANTOREC.POST_CODE, vbUnicode)
        Edit = Edit & StrConv(TANTOREC.POST_CODE, vbUnicode) & "     "
        Edit = Edit & StrConv(TANTOREC.KUBUN, vbUnicode)
        '2011.09.06
         
        List1.AddItem Edit
         
        com = BtOpGetNext
    Loop
    
    List_Proc = False
    
End Function
Private Sub Clear_Field(Optional Mode As Integer = 0)
'----------------------------------------------------------------------------
'                   ��ʓ��e�����ݒ�
'----------------------------------------------------------------------------

Dim i As Integer
    
    For i = Mode To Text_Max
        Text(i) = ""
    Next i

End Sub
Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   ���͓��e�̃`�F�b�N
'----------------------------------------------------------------------------
    Err_Chk = True
    
    If Len(Text(ptxTANTO_CODE).Text) = 0 Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(ptxTANTO_CODE).SetFocus
        Exit Function
    End If
        
    Err_Chk = False
End Function

Private Sub Item_Dsp()
'----------------------------------------------------------------------------
'                   ���ו\��
'----------------------------------------------------------------------------
    Text(ptxTANTO_CODE).Text = Trim(StrConv(TANTOREC.TANTO_CODE, vbUnicode))
    Text(ptxTANTO_NAME).Text = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
    Text(ptxPOST_CODE).Text = Trim(StrConv(TANTOREC.POST_CODE, vbUnicode))
    '2011.09.06
    Text(ptxKUBUN).Text = Trim(StrConv(TANTOREC.KUBUN, vbUnicode))

End Sub
Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �S���҃}�X�^�̒ǉ��^�C��
'----------------------------------------------------------------------------
Dim sts As Integer
Dim ans As Integer
Dim com As Integer

    Update_Proc = True
    
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text(ptxTANTO_CODE).Text)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
                com = BtOpUpdate
                Exit Do
            Case BtErrKeyNotFound
                com = BtOpInsert
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANTO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Call Clear_Field(0)
                    Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                Exit Function
        End Select
    Loop
                                            '���R�[�h���e�ҏW
    Call UniCode_Conv(TANTOREC.TANTO_CODE, Text(ptxTANTO_CODE).Text)
    Call UniCode_Conv(TANTOREC.TANTO_NAME, Text(ptxTANTO_NAME).Text)
    Call UniCode_Conv(TANTOREC.POST_CODE, Text(ptxPOST_CODE).Text)
    
    '2011.09.06
    Call UniCode_Conv(TANTOREC.KUBUN, Text(ptxKUBUN).Text)
    '2011.09.06
    
    
    Call UniCode_Conv(TANTOREC.FILLER, "")

    Do
        sts = BTRV(com, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANTO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                    Call Clear_Field(0)
                    Update_Proc = False
                    Exit Function
                End If
                        
            Case Else
                Call File_Error(sts, com, "�S���҃}�X�^")
                Exit Function
        End Select
    Loop
    
    Call List_Update_Proc(0)                '���X�g�{�b�N�X�X�V

    Call Clear_Field(0)                     '��ʃN���A�[
    
    Update_Proc = False
End Function
Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   �S���҃}�X�^�̍폜
'----------------------------------------------------------------------------
Dim sts As Integer
Dim ans As Integer

    Delete_Proc = True
    
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text(ptxTANTO_CODE).Text)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                Delete_Proc = False
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANTO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Call Clear_Field(0)
                    Delete_Proc = False
                    
                    Exit Function
                
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                Exit Function
        End Select
    Loop
        
    Do
        sts = BTRV(BtOpDelete, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE         '����͖���
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł�<MTS.DAT>�B", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
                    Call Clear_Field(0)
                    Delete_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "�S���҃}�X�^")
                Exit Function
        End Select
    Loop
    
    Call List_Update_Proc(1)                '���X�g�{�b�N�X�X�V
    
    Call Clear_Field(0)                     '��ʃN���A�[

    Delete_Proc = False
End Function


Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case Index
        Case 0
                                            '�G���[�`�F�b�N
            sts = Err_Chk()
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            Beep
            yn = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
            Text(ptxTANTO_CODE) = ""
        Case 3
            If Trim(Text(ptxTANTO_CODE).Text) = "" Then
                Beep
                MsgBox "�폜����R�[�h���w�肵�ĉ������B", vbExclamation
            Else
                Beep
                yn = MsgBox("�폜���܂����H", vbYesNo + vbQuestion, "�m�F����")
                If yn = vbYes Then
                    If Delete_Proc() Then
                        Unload Me
                    End If
                End If
            End If
        Case 8                  '�f�[�^�o��
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
    
    Text(ptxTANTO_CODE).SetFocus

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
Dim i As Integer
Dim c As String * 128
Dim sts As Integer

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
                                '�b�r�u�t�@�C������荞��
    If GetIni("FILE", "TANTO_CSV", "SYS", c) Then
        Beep
        MsgBox "�S���҃}�X�^�f�[�^�o�͗p�t�@�C��[TANTO_CSV]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    
    TANTO_CSV = Trim(c)
    Me.Caption = Me.Caption & " " & LAST_UPDATE_DAY '2019.06.25 �^�C�g���o�[�p��F101080��me.�ɕύX
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Show
    
    Call List_Proc
    Text(ptxTANTO_CODE).SetFocus
    
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
    sts = BTRV(BtOpReset, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "�S���҃}�X�^")
    End If
    Set F1010801 = Nothing
    End
End Sub

Private Sub List1_DblClick()
Dim sts     As Integer

    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Left(List1.List(List1.ListIndex), 5))
    
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            MsgBox "�}�X�^���e���ύX����Ă��܂��B�ŐV�����ĕ\�����܂��B"
            If List_Proc() Then
                Unload Me
            End If
            Exit Sub
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
            Unload Me
    End Select
    
    Call Item_Dsp
    Text(ptxTANTO_CODE).SetFocus
        
End Sub


Private Sub List1_GotFocus()
    
    If List1.ListCount > 0 Then
        If List1.ListIndex <= 0 Then
            List1.ListIndex = 0
        End If
    End If

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
            
        Call List1_DblClick
    
    End If

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
Dim sts As Integer
Dim i   As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub

    Select Case Index
        
        Case ptxTANTO_CODE
            
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text(ptxTANTO_CODE).Text)
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call Clear_Field(1)
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�S���҃}�X�^")
                    Unload Me
            End Select
    
    End Select
    For i = Index + 1 To Text_Max
        If Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i
End Sub

Private Sub List_Update_Proc(Mode As Integer)
'----------------------------------------------------------------------------
'                   ���X�g�{�b�N�X�X�V
'----------------------------------------------------------------------------
Dim i       As Integer
Dim Edit    As String


    For i = 0 To List1.ListCount - 1
        
        
        If Trim(Text(ptxTANTO_CODE).Text) = Trim(Left(List1.List(i), 5)) Then
            List1.RemoveItem i
        End If
    
    Next i

    If Mode = 0 Then
        Edit = StrConv(TANTOREC.TANTO_CODE, vbUnicode) & "    "
        Edit = Edit & StrConv(TANTOREC.TANTO_NAME, vbUnicode) & "   "
        
        '2011.09.06
'        Edit = Edit & StrConv(TANTOREC.POST_CODE, vbUnicode)
        Edit = Edit & StrConv(TANTOREC.POST_CODE, vbUnicode) & "     "
        Edit = Edit & StrConv(TANTOREC.KUBUN, vbUnicode)
        '2011.09.06
        
        List1.AddItem Edit
    End If
End Sub
Private Function Data_Proc() As Integer

Dim FileNo          As Integer
Dim FileName        As String
Dim ret             As Integer

Dim com             As Integer
Dim sts             As Integer

    Call Input_Lock

    FileNo = FreeFile
    FileName = TANTO_CSV
    
    ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), ret) & Right(Trim(FileName), Len(Trim(FileName)) - ret)

    On Error GoTo Error_Proc

    Open (FileName) For Output As FileNo
    
    '2011.09.06
'    Write #FileNo, "�S���Һ���", "�S���Җ���", "����"
    Write #FileNo, "�S���Һ���", "�S���Җ���", "����", "�敪"
    '2011.09.06

    com = BtOpGetFirst
    Do
        DoEvents
        sts = BTRV(com, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�S���҃}�X�^")
                Exit Function
        End Select
    
        Write #FileNo, StrConv(TANTOREC.TANTO_CODE, vbUnicode),
        Write #FileNo, StrConv(TANTOREC.TANTO_NAME, vbUnicode),
        
        '2011.09.06
'        Write #FileNo, StrConv(TANTOREC.POST_CODE, vbUnicode)
        Write #FileNo, StrConv(TANTOREC.POST_CODE, vbUnicode),
        Write #FileNo, StrConv(TANTOREC.KUBUN, vbUnicode)
        '2011.09.06
    
    
        com = BtOpGetNext
    Loop

    Close #FileNo
    
    Call Input_UnLock
    
    Beep
    MsgBox "�u" & FileName & "�v�͐���ɏo�͂���܂����B"


    Exit Function
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox FileName & "���g�p���ł��B"
        Data_Proc = False
    Else
        MsgBox "Err.Number= " & Err.Number
        Data_Proc = True
    End If

    Call Input_UnLock



End Function


