VERSION 5.00
Begin VB.Form PM000401 
   Caption         =   "���i���V�X�e���@�i�ڃ}�X�^�����e�i���X�i���i���x�����ځj"
   ClientHeight    =   6840
   ClientLeft      =   1920
   ClientTop       =   2730
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
   Begin VB.TextBox txSEL_KEY 
      Height          =   375
      Left            =   8640
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   1920
      Locked          =   -1  'True
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   885
   End
   Begin VB.ListBox List1 
      Height          =   4380
      Index           =   0
      ItemData        =   "PM000401.frx":0000
      Left            =   2040
      List            =   "PM000401.frx":0002
      TabIndex        =   2
      Top             =   1080
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   3480
      MaxLength       =   20
      TabIndex        =   1
      Top             =   480
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
      TabIndex        =   13
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
      Index           =   8
      Left            =   7800
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
      Index           =   7
      Left            =   6480
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
      Index           =   6
      Left            =   5640
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
      Index           =   5
      Left            =   4800
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
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
      TabIndex        =   5
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
      TabIndex        =   4
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
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
      TabIndex        =   18
      Top             =   6360
      Width           =   180
   End
   Begin VB.Label Label 
      Caption         =   "�����O"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "�i��"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   15
      Top             =   600
      Width           =   495
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "PM000401"
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


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PM000401.MousePointer = vbHourglass

    Call Ctrl_Lock(PM000401)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PM000401)


    PM000401.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
Dim Item_Key    As String * 20
Dim sts         As Integer
Dim i           As Integer
Dim j           As Integer


    Error_Check_Proc = True


    Select Case Mode
        Case ptxHIN_GAI


            '========================================================= 2007/03/19 =====
''            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
''            Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI), 1))
''            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
''
''            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
''            Select Case sts
''                Case BtNoErr
''                    Item_Key = Text1(Mode).Text
''
''                    txSEL_KEY.Text = Last_JGYOBU & Right(Combo1(pcmbNAIGAI).Text, 1) & Item_Key
''
''                    G_SCREEN_FLG = G_SCREEN_UPD
''                    If Item_Input_Proc() Then           '���ד���
''                        Unload Me
''                    End If
''
''                Case BtErrKeyNotFound
''                    If List_Disp_Proc() Then
''                        Exit Function
''                    End If
''                Case Else
''                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
''                    Exit Function
''            End Select


            If Trim(Text1(Mode).Text) = "" Then
                If List_Disp_Proc() Then
                    Exit Function
                End If
            Else
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
                    
                    
                    
                        For i = 0 To UBound(JGYOBU_T)
                            SubMenu(i).Checked = False
                        Next i
        
        
                        For i = 0 To UBound(JGYOBU_T)
                            For j = 0 To Combo1(pcmbNAIGAI).ListCount - 1
                                Call UniCode_Conv(K0_ITEM.JGYOBU, JGYOBU_T(i).CODE)
                                Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).List(j), 1))
                                Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
        
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                Select Case sts
                                    Case BtNoErr
                                        PM000401.Caption = "���i���V�X�e���@�i�ڃ}�X�^�����e�i���X�i���i���x�����ځj�i" + RTrim(JGYOBU_T(i).NAME) + ")"
                                        SubMenu(i).Checked = True
                                        LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
                                        LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
                                        Last_JGYOBU = JGYOBU_T(i).CODE
        
                                        Combo1(pcmbNAIGAI).ListIndex = j
        
                                        Item_Key = Text1(Mode).Text
                                        txSEL_KEY.Text = Last_JGYOBU & Right(Combo1(pcmbNAIGAI).Text, 1) & Item_Key
        
                                        G_SCREEN_FLG = G_SCREEN_UPD
                                        If Item_Input_Proc() Then           '���ד���
                                            Unload Me
                                        End If
        
                                        Exit For
        
                                    Case BtErrKeyNotFound
                                        Exit For
        
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                        Exit Function
                                End Select
        
                            Next j
        
                            If sts = BtNoErr Then
                                Exit For
                            End If
        
                        Next i
        
                        If sts <> BtNoErr Then
                            If List_Disp_Proc() Then
                                Exit Function
                            End If
                        End If
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Exit Function
                End Select







            End If
            '==========================================================================


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
    
    PM000401.MousePointer = vbHourglass
    
    List1(plstITEM).Clear
    
    '�i��Ͻ��ǂݍ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)                          '���ƕ�
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))    '�����O
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)              '�i�ԁi�O���j
    
    com = BtOpGetGreaterEqual
    
    
    Do
    
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
        
        
        If List1(plstITEM).ListCount > 100 Then
            Exit Do
        End If
        
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
    PM000401.MousePointer = vbDefault

    List_Disp_Proc = False

End Function
Private Function Item_Input_Proc() As Integer
'----------------------------------------------------------------------------
'                   ��ƊǗ����ד��͉�ʁ@�\��
'----------------------------------------------------------------------------
    Item_Input_Proc = True

    If G_SCREEN_FLG <> G_SCREEN_INS Then
        
'        If List1(plstITEM).ListCount = 0 Then
'            Exit Function                           '�ް�������������
'        End If
    
    End If
    
    PM000402.Show vbModal                           '���ד��̓t�H�[���\��
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
                                
                                
                                '���x������p�R���g���[���e�l��2008.05.30
    If GetIni("FILE", "labelprint", "SYS", c) Then
        Beep
        MsgBox "���x������p�R���g���[���e�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LabelPrint_F = RTrim(c)
                                
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
        
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            PM000401.Caption = "���i���V�X�e���@�i�ڃ}�X�^�����e�i���X�i���i���x�����ځj�i" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                
    Unload SubMenu(i)
                                
                                
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If L_ITEM_Open(BtOpenNomal) Then
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
                                
    Call P_CODE_TBL_Proc
                                
    Load PM000402
                                
    W_Index = -1
    
    
    Combo1(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo1(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    
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
                                            
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
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
    
                                            '�󕥐�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�󕥐�}�X�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PM000401 = Nothing
    Set PM000402 = Nothing

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
    
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    PM000401.Caption = "���i���V�X�e���@�i�ڃ}�X�^�����e�i���X�i���i���x�����ځj�j�i" + RTrim(JGYOBU_T(Index).NAME) + "�j"
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

