VERSION 5.00
Begin VB.Form PM000601 
   Caption         =   "���i���V�X�e���@�N���X�}�X�^�����e�i���X"
   ClientHeight    =   6840
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
   ScaleHeight     =   6840
   ScaleWidth      =   11295
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.ComboBox Combo1 
      Appearance      =   0  '�ׯ�
      Height          =   360
      Index           =   0
      Left            =   1680
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   18
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txSEL_KEY 
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   4380
      Index           =   0
      ItemData        =   "PM000601.frx":0000
      Left            =   1680
      List            =   "PM000601.frx":0002
      TabIndex        =   1
      Top             =   1080
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I ��"
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Index           =   4
      Left            =   3960
      TabIndex        =   6
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�V �K"
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "�d������"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   17
      Top             =   240
      Width           =   1095
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
      TabIndex        =   16
      Top             =   6360
      Width           =   180
   End
   Begin VB.Label Label 
      Caption         =   "�N���X"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   14
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "PM000601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'�e�L�X�g�p�Y��
Private Const ptxCLASS_CODE% = 0            '�N���X�R�[�h

'���X�g�p�Y��
Private Const plstP_CLASS% = 0

'�R���{�p�Y����
Private Const pcmbSHIMUKE% = 0              '�����O

Private W_Index As Integer


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PM000601.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000601)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000601)


    PM000601.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim Item_Key    As String * 20
    
    
    Error_Check_Proc = True
    
    Select Case Mode
        Case ptxCLASS_CODE
            
        Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))    '�d������
        Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxCLASS_CODE).Text)                         '�i�ԁi�O���j
        
    
    
        sts = BTRV(BtOpGetEqual, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
            Select Case sts
                Case BtNoErr
                
            
                    Item_Key = Text1(Mode).Text
                    
                    
                    
                    
                    
                    txSEL_KEY.Text = StrConv(P_CLASSREC.SHIMUKE_CODE, vbUnicode) & Item_Key
                
                    
                    G_SCREEN_FLG = G_SCREEN_UPD
                    If Item_Input_Proc() Then           '���ד���
                        Unload Me
                    End If
            
                
                
                Case BtErrKeyNotFound
                    If List_Disp_Proc() Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function
            End Select
            
            
            
            
            
            
            
    End Select
    
        
    
        
        
    Error_Check_Proc = False
    

End Function
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���X�g�{�b�N�X�\��
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer



    List_Disp_Proc = True
    
    List1(plstP_CLASS).Clear
    
    
    '�N���X�}�X�^�ǂݍ���
    Call UniCode_Conv(K0_P_CLASS.SHIMUKE_CODE, Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2))    '�d������
    Call UniCode_Conv(K0_P_CLASS.CLASS_CODE, Text1(ptxCLASS_CODE).Text)                     '�i�ԁi�O���j
    
    com = BtOpGetGreaterEqual
    
    Do
    
        sts = BTRV(com, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_CLASSREC.SHIMUKE_CODE, vbUnicode) <> Mid(Right(Combo1(pcmbSHIMUKE).Text, 4), 1, 2) Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�N���X�}�X�^")
                Exit Function
        
        End Select



        
        List1(plstP_CLASS).AddItem StrConv(P_CLASSREC.CLASS_CODE, vbUnicode) & " " & _
                                    StrConv(P_CLASSREC.CLASS_NAME, vbUnicode)
        
        If List1(plstP_CLASS).ListCount > 99999 Then
            Exit Do
        End If
        
        
        com = BtOpGetNext
    
    Loop

    DoEvents


    If List1(plstP_CLASS).ListCount = 0 Then
        
        W_Index = -1
        Text1(pcmbSHIMUKE).SetFocus
    
    Else
        List1(plstP_CLASS).SetFocus
        List1(plstP_CLASS).ListIndex = 0
            
    End If

    List_Disp_Proc = False

End Function
Private Function Item_Input_Proc() As Integer
'----------------------------------------------------------------------------
'                   ��ƊǗ����ד��͉�ʁ@�\��
'----------------------------------------------------------------------------
    Item_Input_Proc = True

    If G_SCREEN_FLG <> G_SCREEN_INS Then
        
'        If List1(plstP_CLASS).ListCount = 0 Then
'            Exit Function                           '�ް�������������
'        End If
    
    End If
    
    PM000602.Show vbModal                           '���ד��̓t�H�[���\��
    If G_SCREEN_FLG = SYS_ERR Then
        Exit Function
    End If

    If List_Disp_Proc() Then                        'ؽ��ޯ���ĕ\��
        Exit Function
    End If
    
    If List1(plstP_CLASS).ListCount = 0 Then
        Text1(pcmbSHIMUKE).SetFocus
    Else
        List1(plstP_CLASS).SetFocus
        If (List1(plstP_CLASS).ListCount - 1) < W_Index Then
            List1(plstP_CLASS).ListIndex = List1(plstP_CLASS).ListCount - 1
        Else
            List1(plstP_CLASS).ListIndex = W_Index
        End If
    End If

    Item_Input_Proc = False

End Function


Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Sub Command1_Click(Index As Integer)

Dim ans As Integer

    Select Case Index
        Case P_CMD_Upd                      '�X�V
        Case P_CMD_DEL                      '�폜
        Case P_CMD_Ins                      '�V�K
        
            G_SCREEN_FLG = G_SCREEN_INS
            If Item_Input_Proc() Then
                Unload Me
            End If
        
        Case P_CMD_DSP                      '����/�\��
        
            If List_Disp_Proc() Then
                Unload Me
            End If
        
                    
        
        
        Case P_CMD_OUT                      '�ް��o��
        Case P_CMD_PRT                      '���
        Case P_CMD_End                      '�I��
            Unload Me
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
            Command1(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c       As String * 128
Dim i       As Integer

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
                                
                                '�N���X�}�X�^�n�o�d�m
    If P_Class_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
    Call P_CODE_TBL_Proc
                                
    Load PM000602
                                
    W_Index = -1
    
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD) Then
        Unload Me
    End If
    
    Show
    
    If Combo1(pcmbSHIMUKE).ListCount = 0 Then
        MsgBox "�d������̓o�^������܂���B"
        Unload Me
    End If
    
    Combo1(pcmbSHIMUKE).ListIndex = 0
       
    Combo1(pcmbSHIMUKE).SetFocus
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '�N���X�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�N���X�}�X�^")
        End If
    End If
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PM000601 = Nothing
    Set PM000602 = Nothing

    End
End Sub

Private Sub List1_DblClick(Index As Integer)
    Select Case Index
        Case plstP_CLASS
            
                W_Index = List1(plstP_CLASS).ListIndex
                txSEL_KEY.Text = Left(Right(Combo1(pcmbSHIMUKE).Text, 4), 2) & Left(List1(plstP_CLASS).List(List1(plstP_CLASS).ListIndex), 20)
            
                
                G_SCREEN_FLG = G_SCREEN_UPD
                If Item_Input_Proc() Then           '���ד���
                    Unload Me
                End If
    End Select
End Sub

Private Sub List1_GotFocus(Index As Integer)
    
    If List1(Index).ListCount > 0 And _
       List1(Index).ListIndex < 0 Then
        List1(Index).ListIndex = 0
    End If

End Sub

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    If Shift = vbShiftMask Then
        Call Tab_Ctrl(Shift)        '�ړ�
    Else
        Select Case Index
            Case plstP_CLASS
            
                W_Index = List1(plstP_CLASS).ListIndex
                txSEL_KEY.Text = Left(Right(Combo1(pcmbSHIMUKE).Text, 4), 2) & Left(List1(plstP_CLASS).List(List1(plstP_CLASS).ListIndex), 20)
            
                
                G_SCREEN_FLG = G_SCREEN_UPD
                If Item_Input_Proc() Then           '���ד���
                    Unload Me
                End If
        End Select
    End If

End Sub


Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then     '�G���[�`�F�b�N
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub


Private Function Code_Set_Proc(Index As Integer, KBN As String) As Integer
'----------------------------------------------------------------------------
'                   �R�[�h�}�X�^���R���{�ɃZ�b�g����B
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim OPTION1     As Integer
Dim OPTION2     As Integer

Dim wkOption    As String



Dim i           As Integer
    
    Code_Set_Proc = True
    
    Combo1(Index).Clear
    
    For i = 0 To UBound(P_KBN_TBL)
    
        If KBN = P_KBN_TBL(i).KBN_CD Then
            Key_Len = P_KBN_TBL(i).KBN_Len
            Exit For
        End If
    
    Next i
    
    If i > UBound(P_KBN_TBL) Then
        Exit Function
    End If
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, KBN)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> KBN Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�R�[�h�}�X�^")
                Exit Function
        
        End Select

        wkOption = ""
        If P_KBN_TBL(i).KBN_OP1 Then
            wkOption = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
        End If
        If P_KBN_TBL(i).KBN_OP2 Then
            wkOption = wkOption & Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        End If
        
        
        
        Combo1(Index).AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    




End Function
