VERSION 5.00
Begin VB.Form PM000301 
   Caption         =   "���ރ}�X�^�����e�i���X"
   ClientHeight    =   12975
   ClientLeft      =   1920
   ClientTop       =   2730
   ClientWidth     =   11790
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
   ScaleHeight     =   12975
   ScaleWidth      =   11790
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox txSEL_KEY 
      Height          =   375
      Left            =   7560
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      Height          =   360
      Index           =   0
      Left            =   9720
      Locked          =   -1  'True
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.ListBox List1 
      Height          =   10620
      Index           =   0
      ItemData        =   "PM000301.frx":0000
      Left            =   840
      List            =   "PM000301.frx":0002
      TabIndex        =   2
      Top             =   1080
      Width           =   10275
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   1560
      MaxLength       =   20
      TabIndex        =   1
      Top             =   360
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
      TabIndex        =   14
      Top             =   12120
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   7
      Top             =   12120
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   5
      Top             =   12120
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   12120
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   12120
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "�i�@��"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   20
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "���ޕi��"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   19
      Top             =   840
      Width           =   975
   End
   Begin VB.Label LabJIGYO 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   480
      TabIndex        =   18
      Top             =   12600
      Width           =   120
   End
   Begin VB.Label Label 
      Caption         =   "�����O"
      Height          =   255
      Index           =   0
      Left            =   8880
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "���ޕi��"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   480
      Width           =   975
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "PM000301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'�e�L�X�g�p�Y��
Private Const ptxHIN_GAI% = 0               '���ޕi��

'���X�g�p�Y��
Private Const plstITEM% = 0

'�R���{�p�Y����
Private Const pcmbNAIGAI% = 0               '�����O

Private W_Index As Integer


'Private Const LAST_UPDATE_DAY$ = "[PM00030]2016.04.22 09:45"
'Private Const LAST_UPDATE_DAY$ = "[PM00030]2016.05.19 16:45"


Private List_Max As Long                    '�ő�\������ 2009.05.29


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PM000301.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000301)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000301)


    PM000301.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
Dim Item_Key    As String * 20
Dim sts         As Integer
    
    
    Error_Check_Proc = True
    
    
    Select Case Mode
        Case ptxHIN_GAI
            
            Text1(Mode).Text = StrConv(Trim(Text1(Mode).Text), vbUpperCase)
            
            
            
            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI), 1))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
            
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                
            
                    Item_Key = Text1(Mode).Text
                    
                    
                    
                    
                    
                    txSEL_KEY.Text = Last_JGYOBU & Right(Combo1(pcmbNAIGAI).Text, 1) & Item_Key
                
                    
                    G_SCREEN_FLG = G_SCREEN_UPD
                    If Item_Input_Proc() Then           '���ד���
                        Unload Me
                    End If
            
                
                
                Case BtErrKeyNotFound
                    If List_Disp_Proc() Then
                        Exit Function
                    End If
                
                    Text1(ptxHIN_GAI).SetFocus
                 
                
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


Dim List_Cnt    As Long


    List_Disp_Proc = True
    PM000301.MousePointer = vbHourglass
    
    List1(plstITEM).Clear
    
    '�i��Ͻ��ǂݍ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)                          '���ƕ�������
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))    '�����O������
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
    com = BtOpGetGreaterEqual
    
    List_Cnt = 0
    Do
        '2009.05.29
        If List_Max = 0 Then
        Else
            If List_Cnt >= List_Max Then
                Exit Do
            End If
        End If
    
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            
        Select Case sts
            Case BtNoErr
            
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo1(pcmbNAIGAI).Text, 1) Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                Exit Function
        
        End Select

        
        List1(plstITEM).AddItem StrConv(ITEMREC.HIN_GAI, vbUnicode) & " " & _
                                    StrConv(ITEMREC.HIN_NAME, vbUnicode)
        
        
        List_Cnt = List_Cnt + 1
        
        com = BtOpGetNext
    
    Loop

    DoEvents


    If List1(plstITEM).ListCount = 0 Then
        
        W_Index = -1
        Text1(ptxHIN_GAI).SetFocus
    
    
    Else
        List1(plstITEM).SetFocus
        List1(plstITEM).ListIndex = 0
    End If
    PM000301.MousePointer = vbDefault

    List_Disp_Proc = False

End Function
Private Function Item_Input_Proc() As Integer
'----------------------------------------------------------------------------
'                   ��ƊǗ����ד��͉�ʁ@�\��
'----------------------------------------------------------------------------
Dim i       As Integer
    
    
    Item_Input_Proc = True

    If G_SCREEN_FLG <> G_SCREEN_INS Then
        
'        If List1(plstITEM).ListCount = 0 Then
'            Exit Function                           '�ް�������������
'        End If
    
    End If
    
    For i = 0 To UBound(JGYOBU_T)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            Exit For
        End If
    Next i
    
    
    
    
    
    
'    PM000302.Caption = "���i���V�X�e���@�i�ڃ}�X�^�����e�i���X�i�Ɩ��Ǘ����ځj�i" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY   '2016.04.22
    PM000302.Caption = "���ރ}�X�^�����e�i���X�i" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY                              '2016.04.22
    PM000302.Show vbModal                       '���ד��̓t�H�[���\��
    
    
    
    
    
    
    
    If G_SCREEN_FLG = SYS_ERR Then
        Exit Function
    End If

    If List_Disp_Proc() Then                        'ؽ��ޯ���ĕ\��
        Exit Function
    End If
    
    If List1(plstITEM).ListCount = 0 Then
        Text1(ptxHIN_GAI).SetFocus
    Else
        List1(plstITEM).SetFocus
        If (List1(plstITEM).ListCount - 1) < W_Index Then
            List1(plstITEM).ListIndex = List1(plstITEM).ListCount - 1
        Else
            List1(plstITEM).ListIndex = W_Index
        End If
    End If

    Item_Input_Proc = False

End Function


Private Sub Command1_Click(Index As Integer)

Dim yn As Integer

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
'    PrintForm
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

'    If App.PrevInstance Then
'        Beep
'        MsgBox "����v���O�������s���ł��B"
'        End
'    End If

                                '���O�t�@�C������荞��
'    If GetIni("FILE", "LOGF", "SYS", c) Then
    If GetIni(App.EXEName, "LOGF", App.EXEName, c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                
                                
                                
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
        
    Last_JGYOBU = SHIZAI
        
        
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
'            PM000301.Caption = "���i���V�X�e���@�i�ڃ}�X�^�����e�i���X�i�Ɩ��Ǘ����ځj�i" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY   '2016.04.22
            PM000301.Caption = "���ރ}�X�^�����e�i���X�i" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY                              '2016.04.22
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                
    Unload SubMenu(i)
                                
                                
                                
                                
                                
                                
                                
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�󕥐�}�X�^�i�d����j�n�o�d�m
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�I�}�X�^�n�o�d�m
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
    
    
                                '�ő�\��������荞�� 2009.05.29
    If GetIni(App.EXEName, "MAX_LINE", App.EXEName, c) Then
        List_Max = 0
    Else
        If IsNumeric(Trim(c)) Then
            List_Max = Val(Trim(c))
        Else
            List_Max = 0
        End If
    End If
    
    
    Call P_CODE_TBL_Proc
                                
    Load PM000302
                                
    W_Index = -1
    
    
    Combo1(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo1(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    
    Last_JGYOBU = SHIZAI
    
    Show
    
    Combo1(pcmbNAIGAI).ListIndex = 0
       
    Text1(ptxHIN_GAI).SetFocus
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
        End If
    End If
    
    
                                            '�󕥃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�󕥃}�X�^")
        End If
    End If
    
    
    
                                            '�I�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�}�X�^")
        End If
    End If
    
    
    
    
    
    
    sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PM000301 = Nothing
    Set PM000302 = Nothing

    End
End Sub

Private Sub List1_DblClick(Index As Integer)
    Select Case Index
        Case plstITEM
        
            W_Index = List1(plstITEM).ListIndex
            txSEL_KEY.Text = Last_JGYOBU & Right(Combo1(pcmbNAIGAI).Text, 1) & Left(List1(plstITEM).List(List1(plstITEM).ListIndex), 20)
        
            
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
            Case plstITEM
            
                W_Index = List1(plstITEM).ListIndex
                txSEL_KEY.Text = Last_JGYOBU & Right(Combo1(pcmbNAIGAI).Text, 1) & Left(List1(plstITEM).List(List1(plstITEM).ListIndex), 20)
            
                
                G_SCREEN_FLG = G_SCREEN_UPD
                If Item_Input_Proc() Then           '���ד���
                    Unload Me
                End If
        End Select
    End If

End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    
                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If
    
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
'    PM000301.Caption = "���i���V�X�e���@�i�ڃ}�X�^�����e�i���X�i�Ɩ��Ǘ����ځj�i" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY   '2016.04.22
    PM000301.Caption = "���ރ}�X�^�����e�i���X�i" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY                              '2016.04.22
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

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

