VERSION 5.00
Begin VB.Form F1060291 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���i���v��x���A���[�����X�g���(����PC����)"
   ClientHeight    =   6945
   ClientLeft      =   2325
   ClientTop       =   2715
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
   ScaleHeight     =   6945
   ScaleWidth      =   11295
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   3780
      TabIndex        =   33
      Top             =   2760
      Width           =   3585
      Begin VB.OptionButton Option1 
         Caption         =   "�S�߰��"
         Height          =   255
         Index           =   1
         Left            =   2205
         TabIndex        =   10
         Top             =   360
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�擪�߰�ނ̂�"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   9
         Top             =   360
         Width           =   2010
      End
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   8085
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   7560
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   7035
      MaxLength       =   2
      TabIndex        =   6
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   6510
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   5670
      MaxLength       =   2
      TabIndex        =   4
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   3
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   4410
      MaxLength       =   2
      TabIndex        =   2
      Top             =   2040
      Width           =   375
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   3735
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   3780
      MaxLength       =   2
      TabIndex        =   1
      Top             =   2040
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
      TabIndex        =   22
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
      Index           =   9
      Left            =   8640
      TabIndex        =   20
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
      Index           =   7
      Left            =   6480
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
      Index           =   6
      Left            =   5640
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
      Index           =   5
      Left            =   4800
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
      Index           =   4
      Left            =   3960
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
      Index           =   3
      Left            =   2640
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
      Index           =   2
      Left            =   1800
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
      Index           =   1
      Left            =   960
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
      Index           =   0
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "-"
      Height          =   255
      Index           =   7
      Left            =   7980
      TabIndex        =   32
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "-"
      Height          =   255
      Index           =   6
      Left            =   7455
      TabIndex        =   31
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "-"
      Height          =   255
      Index           =   5
      Left            =   6930
      TabIndex        =   30
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "�`"
      Height          =   255
      Index           =   4
      Left            =   6195
      TabIndex        =   29
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "-"
      Height          =   255
      Index           =   3
      Left            =   5460
      TabIndex        =   28
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "-"
      Height          =   255
      Index           =   2
      Left            =   4830
      TabIndex        =   27
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "-"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   26
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   255
      Index           =   33
      Left            =   2895
      TabIndex        =   25
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "�W���I��"
      Height          =   255
      Index           =   0
      Left            =   2625
      TabIndex        =   24
      Top             =   2160
      Width           =   1035
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
      TabIndex        =   23
      Top             =   6480
      Width           =   180
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1060291"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxS_SOKO% = 0                '�J�n�@�W���I�ԁ@�q��
Private Const ptxS_RETU% = 1                '�J�n�@�W���I�ԁ@��
Private Const ptxS_REN% = 2                 '�J�n�@�W���I�ԁ@�A
Private Const ptxS_DAN% = 3                 '�J�n�@�W���I�ԁ@�i

Private Const ptxE_SOKO% = 4                '�J�n�@�W���I�ԁ@�q��
Private Const ptxE_RETU% = 5                '�J�n�@�W���I�ԁ@��
Private Const ptxE_REN% = 6                 '�J�n�@�W���I�ԁ@�A
Private Const ptxE_DAN% = 7                 '�J�n�@�W���I�ԁ@�i


Private Const Text_Max% = 7                 '��ʍ��ڕʍő���ޯ��

Private Const pcmbNaigai% = 0               '�����O


Private Const LMAX% = 42                    '�œ��ő�s��
Private Const LCTL% = 99                    '
Private Const MGN_L% = 3                   '���]���i�����F�P����j
Private Const MGN_U% = 1                    '��]���i�s���F�P����j

Private Pdate As String                     '����J�n���t�iͯ�ް�p�j
Private Ptime As String                     '����J�n�����iͯ�ް�p�j


Private NormalFont  As New StdFont          '����t�H���g
Private MidFont     As New StdFont          '����t�H���g

Private OutSide     As Long                 '����ΊO�o�א�

Private GOODS_DATA  As String               '�o�̓f�[�^�t�@�C����


Private Type EE_ZAIKO_TBL_tag
    EE_LOC          As String * 8
    EE_QTY          As Long
End Type

Private EE_ZAIKO_TBL(0 To 2) As EE_ZAIKO_TBL_tag


Private Function Err_Chk(INDEX As Integer) As Integer
'----------------------------------------------------------------------------
'                   �G���[�`�F�b�N����
'----------------------------------------------------------------------------
                                            
Dim i   As Integer
Dim sts As Integer

                                            
    Err_Chk = True
            
    For i = ptxS_SOKO To ptxE_DAN
    
    
        If i = ptxS_SOKO Or ptxE_SOKO Then
        Else
        
            If IsNumeric(Text(i).Text) Then
            
                Text(i).Text = Format(CInt(Text(i).Text), "00")
            
            Else
                
                MsgBox "���͂������ڂ́A�G���[�ł�"
                Text(i).SetFocus
                Exit Function
            End If
        
        
        End If
    
    
    Next i
            
    Select Case INDEX
        Case ptxS_SOKO
        Case ptxS_RETU
        Case ptxS_REN
        Case ptxS_DAN
        Case ptxE_SOKO
        Case ptxE_RETU
        Case ptxE_REN
        Case ptxE_DAN
            If (Text(ptxS_SOKO).Text & Text(ptxS_RETU).Text & Text(ptxS_REN).Text & Text(ptxS_DAN).Text) > _
                (Text(ptxE_SOKO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                
                MsgBox "���͂������ڂ́A�G���[�ł�"
                Text(i).SetFocus
                
                
                Exit Function
            End If
    
    End Select
            
    
    
    
    
    Err_Chk = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1060291.MousePointer = vbHourglass

    Call Ctrl_Lock(F1060291)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1060291)


    F1060291.MousePointer = vbDefault

End Sub


Private Sub Command_Click(INDEX As Integer)

Dim yn      As Integer
Dim i       As Integer
    
Dim mesg    As String
    
    Select Case INDEX
        
        
        
        Case 8                              '���
            
            
            For i = ptxS_SOKO To ptxE_DAN
                If Err_Chk(i) Then
                    Exit Sub
                End If
            Next i
            
            
            Beep
            yn = MsgBox("������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                
                If Print_Proc() Then
                    Unload Me
                End If
            End If
            Text(ptxS_SOKO).SetFocus
                    
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

'
Private Sub Form_Load()

Dim c   As String * 128
Dim i   As Integer
     
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
            F1060291.Caption = "���i���v��x���A���[�����X�g���(����PC����)�i" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                
    Unload SubMenu(i)
                                
                                
                                
                                '���i���x���t�@�C������荞��
    If GetIni("FILE", "GOODS_DATA", "SYS", c) Then
        Beep
        MsgBox "'���i���x���t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    GOODS_DATA = Trim(c)
                                '�ΏۊO�o�א���荞��
    If GetIni(App.EXEName, "OUTSIDE", "SYS", c) Then
        OutSide = 0
    Else
        If IsNumeric(Trim(c)) Then
            OutSide = CLng(Trim(c))
        Else
            OutSide = 0
        End If
    End If
                                
                                
                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�t�@�C���n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�����Ϗo�א��n�o�d�m
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���i���W�v�t�@�C���n�o�d�m
    If GOODS_ONO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '����t�H���g�ݒ�(�ʏ�)
    With NormalFont
        .NAME = F1060291.FontName
        .Size = 12
    End With

                                '����t�H���g�ݒ�i���j
    With MidFont
        .NAME = F1060291.FontName
        .Size = 8
    End With


    Combo(pcmbNaigai).Clear
    Combo(pcmbNaigai).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo(pcmbNaigai).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo(pcmbNaigai).ListIndex = 0

    Option1(0).Value = True
    Option1(1).Value = False


    Show
    
    Text(ptxS_SOKO).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�����Ϗo�א��b�k�n�r�d
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�����Ϗo�א�")
        End If
    End If
                                            '���i���W�v�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, GOODS_ONO_POS, GOODS_ONOREC, Len(GOODS_ONOREC), K0_GOODS_ONO, Len(K0_GOODS_ONO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���W�v�t�@�C��")
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1060291 = Nothing

    End
End Sub


Private Sub SubMenu_Click(INDEX As Integer)
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
    F1060291.Caption = "���i���v��x���A���[�����X�g���(����PC����)�i" + RTrim(JGYOBU_T(INDEX).NAME) + "�j"
    Last_JGYOBU = JGYOBU_T(INDEX).CODE
    SubMenu(INDEX).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(INDEX).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(INDEX).COLOR)

End Sub

Private Sub Text_GotFocus(INDEX As Integer)
    
    If Text(INDEX).TabStop = True Then
        Text(INDEX) = Trim(Text(INDEX).Text)
        Text(INDEX).SelStart = 0
        Text(INDEX).SelLength = Len(Text(INDEX).Text)
    End If

End Sub

Private Sub Text_KeyDown(INDEX As Integer, KeyCode As Integer, Shift As Integer)
Dim i   As Integer
Dim sts As Integer

    

    If KeyCode <> vbKeyReturn Then Exit Sub
        
    If Err_Chk(INDEX) Then
        Exit Sub
    End If
    
    For i = INDEX + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i
End Sub


Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���i���x���A���[�����X�g�������
'----------------------------------------------------------------------------
Dim Lcnt        As Integer

Dim sts         As Integer
Dim com         As Integer

''Dim Save_Soko   As String * 2

Dim Edit        As String

Dim X_Tab       As Integer

Dim Mode        As Integer

    Print_Proc = True

    Call Input_Lock



    If Data_Make_Proc() Then        '���i���x���W�v�f�[�^�쐬
        Exit Function
    End If



    Lcnt = LCTL

    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time
    
    Call UniCode_Conv(K1_GOODS_ONO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_GOODS_ONO.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
    
    Call UniCode_Conv(K1_GOODS_ONO.AVE_SYUKA, "99999999")
    Call UniCode_Conv(K1_GOODS_ONO.Sumi_QTY, "")
    Call UniCode_Conv(K1_GOODS_ONO.Mi_QTY, "99999999")
    
    
    Call UniCode_Conv(K1_GOODS_ONO.ST_SOKO, "")
    Call UniCode_Conv(K1_GOODS_ONO.ST_RETU, "")
    Call UniCode_Conv(K1_GOODS_ONO.ST_REN, "")
    Call UniCode_Conv(K1_GOODS_ONO.ST_DAN, "")
    Call UniCode_Conv(K1_GOODS_ONO.HIN_GAI, "")
    
    
    
    
    com = BtOpGetGreater
    
    Do
        sts = BTRV(com, GOODS_ONO_POS, GOODS_ONOREC, Len(GOODS_ONOREC), K1_GOODS_ONO, Len(K1_GOODS_ONO), 1)
        Select Case sts
            Case BtNoErr
                If StrConv(GOODS_ONOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(GOODS_ONOREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                    Exit Do
                End If
                                        
                
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���i���W�v�t�@�C��")
                Exit Function
        End Select

'-------------------------------------------------  ���׈��
''        If com = BtOpGetGreater Then
''            Save_Soko = StrConv(GOODS_ONOREC.ST_SOKO, vbUnicode)
''
''            Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
''            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
''            Select Case sts
''                Case BtNoErr
''                    If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
''                        Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
''                    End If
''                Case BtErrKeyNotFound
''                    '�l�����Ȃ��������͌p��
''                    Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
''                    Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
''                Case Else
''                    Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
''                    Exit Function
''            End Select
''
''        End If
''
''        If Save_Soko <> StrConv(GOODS_ONOREC.ST_SOKO, vbUnicode) Then
''
''            Lcnt = LMAX + 1
''            Save_Soko = StrConv(GOODS_ONOREC.ST_SOKO, vbUnicode)
''
''            Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
''            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
''            Select Case sts
''                Case BtNoErr
''                    If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
''                        Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
''                    End If
''
''                Case BtErrKeyNotFound
''                        '�l�����Ȃ��������͌p��
''                    Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
''                    Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
''                Case Else
''                    Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
''                    Exit Function
''            End Select
''
''        End If
        
        
        Call UniCode_Conv(K0_SOKO.Soko_No, StrConv(GOODS_ONOREC.ST_SOKO, vbUnicode))
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
             Case BtNoErr
                  If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                      Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                  End If

              Case BtErrKeyNotFound
                      '�l�����Ȃ��������͌p��
                  Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
                  Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
              Case Else
                  Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                  Exit Function
        End Select
        
        
        If CLng(StrConv(GOODS_ONOREC.SUMI_PERCENT, vbUnicode)) > CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                        '�ݒ蔭���_���傫��
        
        Else
            '�����i�݌Ɂ��O �́A����ΏۊO 2004.08.27
            If OutSide >= CLng(StrConv(GOODS_ONOREC.AVE_SYUKA, vbUnicode)) Or _
                CLng(StrConv(GOODS_ONOREC.Mi_QTY, vbUnicode)) <= 0 Then
            Else
                
                If Head_Print_Proc(Lcnt, Mode) Then
                    Exit Function
                End If
            
                If Mode Then
                    Exit Do
                End If
                
                X_Tab = MGN_L
            
                Printer.Print Tab(X_Tab);
                                                        '�W���I��
                Edit = StrConv(GOODS_ONOREC.ST_SOKO, vbUnicode) & "-"
                Edit = Edit & StrConv(GOODS_ONOREC.ST_RETU, vbUnicode) & "-"
                Edit = Edit & StrConv(GOODS_ONOREC.ST_REN, vbUnicode) & "-"
                Edit = Edit & StrConv(GOODS_ONOREC.ST_DAN, vbUnicode)
                Printer.Print Edit;
                X_Tab = X_Tab + Len(Edit) + 2
                                                        '�i�ԁi�O���j
                Printer.Print Tab(X_Tab);

                Printer.Print StrConv(GOODS_ONOREC.HIN_GAI, vbUnicode);
                X_Tab = X_Tab + Len(StrConv(GOODS_ONOREC.HIN_GAI, vbUnicode)) + 4
                                                        '����
                Printer.Print Tab(X_Tab);
                Printer.Print StrConv(GOODS_ONOREC.PACKING_NO, vbUnicode);
                X_Tab = X_Tab + Len(StrConv(GOODS_ONOREC.PACKING_NO, vbUnicode)) + 4
                                                        '���i���ςݍ݌ɐ�
                Printer.Print Tab(X_Tab);
                Edit = Format(CLng(StrConv(GOODS_ONOREC.Sumi_QTY, vbUnicode)), "#,##0")
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If
                Printer.Print Edit;
                X_Tab = X_Tab + Len(Edit) + 2
                                                        '�����i�݌ɐ�
                Printer.Print Tab(X_Tab);
                Edit = Format(CLng(StrConv(GOODS_ONOREC.Mi_QTY, vbUnicode)), "#,##0")
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If
                Printer.Print Edit;
                X_Tab = X_Tab + Len(Edit) + 2
                                                        '�����Ϗo�א�
                Printer.Print Tab(X_Tab);
                Edit = Format(CLng(StrConv(GOODS_ONOREC.AVE_SYUKA, vbUnicode)), "#,##0")
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If
                Printer.Print Edit;
                X_Tab = X_Tab + Len(Edit) + 2
                                                        '���O���i���K�v��
                Printer.Print Tab(X_Tab);
                Edit = Format(CLng(StrConv(GOODS_ONOREC.AVE_SYUKA, vbUnicode)) - CLng(StrConv(GOODS_ONOREC.Sumi_QTY, vbUnicode)), "#,##0")
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If
                Printer.Print Edit;
                X_Tab = X_Tab + Len(Edit) + 2
                                                        '���O���i����
                Printer.Print Tab(X_Tab);
                Edit = Format(CInt(StrConv(GOODS_ONOREC.SUMI_PERCENT, vbUnicode)), "#0") & "%"
                If Len(Edit) < 10 Then
                    Edit = Space(10 - Len(Edit)) & Edit
                End If

                Printer.Print Edit;
                X_Tab = X_Tab + Len(Edit) + 5
                                                        '�ʒu�݌�
                Printer.Print Tab(X_Tab);

                If MI_ZAIKO_KENSAKU(StrConv(GOODS_ONOREC.HIN_GAI, vbUnicode)) Then
                    Exit Function
                End If

                If Len(Trim(EE_ZAIKO_TBL(0).EE_LOC)) <> 0 Then
                    Edit = Format(EE_ZAIKO_TBL(0).EE_QTY, "#0")
                    If Len(Edit) < 9 Then
                        Edit = Space(9 - Len(Edit)) & Edit
                    End If
                    Edit = Left(EE_ZAIKO_TBL(0).EE_LOC, 2) & "-" & _
                           Mid(EE_ZAIKO_TBL(0).EE_LOC, 3, 2) & "-" & _
                           Mid(EE_ZAIKO_TBL(0).EE_LOC, 5, 2) & "-" & _
                           Right(EE_ZAIKO_TBL(0).EE_LOC, 2) & Edit
                Else
                    Edit = ""
                End If

                Printer.Print Edit

                Printer.Print
            
                Lcnt = Lcnt + 2
        
            End If
            com = BtOpGetNext
        End If
    Loop

    Printer.EndDoc


    Call Input_UnLock
    
    Print_Proc = False

End Function

Private Function Head_Print_Proc(Lcnt As Integer, Mode As Integer) As Integer

Dim sts     As Integer
Dim i       As Integer

    Head_Print_Proc = True
    
    If Lcnt < LMAX Then
        
        Head_Print_Proc = False
        Exit Function
    
    End If

    Mode = False

    If Option1(0).Value Then
        If Printer.Page > 1 Then
            Mode = True
            Head_Print_Proc = False
            Exit Function
        End If
    End If
        

    If Lcnt = LCTL Then
    Else
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
    Printer.Print Tab(MGN_L);
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).CODE Then
            Printer.Print "�i" & RTrim(JGYOBU_T(i).NAME) & "�j";
            Exit For
        End If
    Next i

    Printer.Print Tab(MGN_L + 55);
    
    Printer.Print "���i���x���A���[�����X�g";
    
    
    Printer.Print Tab(MGN_L + 90);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print

    
'    Printer.Print Tab(MGN_L);
'    Printer.Print "�q�ɁF";
'    Printer.Print StrConv(SOKOREC.Soko_No, vbUnicode) & "  ";
'    Printer.Print StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "  ";
'    Printer.Print "�i�ݒ蔭���_ " & StrConv(Format(CLng(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0"), vbWide) & "���j"
'    Printer.Print

'    Printer.Print Tab(MGN_L);
'    Printer.Print "�W���I��";
'    Printer.Print Tab(MGN_L + 13);
'    Printer.Print "�i�ԁi�O���j";
'    Printer.Print Tab(MGN_L + 26);
'    Printer.Print "����(����)";
'    Printer.Print Tab(MGN_L + 38);
'    Printer.Print "���i���ύ݌�";
'    Printer.Print Tab(MGN_L + 58);
'    Printer.Print "�����i�݌�";
'    Printer.Print Tab(MGN_L + 74);
'    Printer.Print "�����Ϗo�א�";
'    Printer.Print Tab(MGN_L + 88);
'    Printer.Print "���O���i���K�v��";
'    Printer.Print Tab(MGN_L + 108);
'    Printer.Print "���O���i����"
'
'    Set Printer.Font = MidFont
'    Printer.Print Tab(MGN_L + 112);
'    Printer.Print "(�ߋ�3����ԕ���)";
'    Printer.Print Tab(MGN_L + 130);
'    Printer.Print "(�����Ϗo�א�-���i���ύ݌�)";
'    Printer.Print Tab(MGN_L + 158);
'    Printer.Print "(���i���ύ݌�/�����Ϗo�א�)"
'
'
'    Set Printer.Font = NormalFont

    Printer.Print Tab(MGN_L);
    Printer.Print "�W���I��";
    Printer.Print Tab(MGN_L + 13);
    Printer.Print "�i�ԁi�O���j";
    Printer.Print Tab(MGN_L + 37);
    Printer.Print "����";
    Printer.Print Tab(MGN_L + 49);
    Printer.Print "���ϐ�";
    Printer.Print Tab(MGN_L + 61);
    Printer.Print "�����i";
    Printer.Print Tab(MGN_L + 73);
    Printer.Print "������";
    Printer.Print Tab(MGN_L + 85);
    Printer.Print "�K�v��";
    Printer.Print Tab(MGN_L + 97);
    Printer.Print "�@��";
    Printer.Print Tab(MGN_L + 120);
    Printer.Print "�ʒu�݌�"

    Printer.Print

    Lcnt = 0
    
    Head_Print_Proc = False

End Function

Private Function Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                   �x���p�W�v�f�[�^�쐬����
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim ans         As Integer

Dim Sumi_QTY    As Long
Dim Mi_QTY      As Long
Dim AVE_QTY     As Long

Dim SKIP_FLG    As Boolean

    Data_Make_Proc = True

'---------------------------------------------------------- '�S���R�[�h�폜
    com = BtOpGetFirst
    Do
        
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, GOODS_ONO_POS, GOODS_ONOREC, Len(GOODS_ONOREC), K0_GOODS_ONO, Len(K0_GOODS_ONO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<GOODS_ONO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "���i���x���W�v�f�[�^")
                    Exit Function
            End Select
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        Do
            
            sts = BTRV(BtOpDelete, GOODS_ONO_POS, GOODS_ONOREC, Len(GOODS_ONOREC), K0_GOODS_ONO, Len(K0_GOODS_ONO), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<GOODS_ONO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "���i���x���W�v�f�[�^")
                    Exit Function
            End Select
        
        Loop
        
        com = BtOpGetNext
    
    Loop
'---------------------------------------------------------- '�i�ڃ}�X�^�x�[�X�Ńf�[�^�쐬

    Call UniCode_Conv(K6_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K6_ITEM.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
    Call UniCode_Conv(K6_ITEM.ST_SOKO, Text(ptxS_SOKO).Text)
    Call UniCode_Conv(K6_ITEM.ST_RETU, Text(ptxS_RETU).Text)
    Call UniCode_Conv(K6_ITEM.ST_REN, Text(ptxS_REN).Text)
    Call UniCode_Conv(K6_ITEM.ST_DAN, Text(ptxS_DAN).Text)
    Call UniCode_Conv(K6_ITEM.HIN_GAI, "")
    
    com = BtOpGetGreater
    
    Do
        
        DoEvents
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K6_ITEM, Len(K6_ITEM), 6)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                    '���ƕ��^�����O�u���[�N
                    Exit Do
                End If
            
                If (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) > _
                    (Text(ptxE_SOKO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                    '�q�ɔԍ��u���[�N
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                Exit Function
        End Select
        '-----------------------------------------  '���i���W�v�t�@�C���쐬
        
        SKIP_FLG = False
        
        
        If StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> GOODS_ON Then
            SKIP_FLG = True
        End If
            
        If Left(StrConv(ITEMREC.HIN_GAI, vbUnicode), 1) = "K" Then
            SKIP_FLG = True
        End If
            
        If IsNumeric(StrConv(ITEMREC.L_URIKIN1, vbUnicode)) Then
            If Val(StrConv(ITEMREC.L_URIKIN1, vbUnicode)) = 0 Then
                SKIP_FLG = True
            End If
        Else
            SKIP_FLG = True
        End If
            
        If IsNumeric(StrConv(ITEMREC.L_URIKIN2, vbUnicode)) Then
            If Val(StrConv(ITEMREC.L_URIKIN2, vbUnicode)) = 0 Then
                SKIP_FLG = True
            End If
        Else
            SKIP_FLG = True
        End If
            
        If IsNumeric(StrConv(ITEMREC.L_URIKIN3, vbUnicode)) Then
            If Val(StrConv(ITEMREC.L_URIKIN3, vbUnicode)) = 0 Then
                SKIP_FLG = True
            End If
        Else
            SKIP_FLG = True
        End If
            
        '2007.06.05 ���`�ԁ��hZ"�́A�ΏۊO
        If Trim(StrConv(ITEMREC.K_KEITAI, vbUnicode)) = "Z" Then
            SKIP_FLG = True
        End If
            
            
            
            
            
        If Not SKIP_FLG Then
                                                    '���ƕ�
            Call UniCode_Conv(GOODS_ONOREC.JGYOBU, Last_JGYOBU)
                                                    '�����O
            Call UniCode_Conv(GOODS_ONOREC.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
                                                    '�i�ԁi�O���j
            Call UniCode_Conv(GOODS_ONOREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                    '�W���I��
            Call UniCode_Conv(GOODS_ONOREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
            Call UniCode_Conv(GOODS_ONOREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
            Call UniCode_Conv(GOODS_ONOREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
            Call UniCode_Conv(GOODS_ONOREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                                    '����
            Call UniCode_Conv(GOODS_ONOREC.PACKING_NO, StrConv(ITEMREC.PACKING_NO, vbUnicode))
            
                                                    '�݌ɏW�v����
            If Zaiko_Syukei_Proc(Sumi_QTY, _
                                    Mi_QTY, _
                                    Last_JGYOBU, _
                                    Right(Combo(pcmbNaigai).Text, 1), _
                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) = SYS_ERR Then
                Exit Function
            End If
                                                    
'            If Mi_QTY = 0 Then                     '�����i�݌�=0 --> <=10 �ɕύX   2007.06.05
            If Mi_QTY <= 10 Then
            Else
                                                        '���i���ςݍ݌ɐ�
                Call UniCode_Conv(GOODS_ONOREC.Sumi_QTY, Format(Sumi_QTY, "00000000"))
                                                        '�����i�݌ɐ�
                Call UniCode_Conv(GOODS_ONOREC.Mi_QTY, Format(Mi_QTY, "00000000"))
                                                        '�����Ϗo�א�
                Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
                Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                
                AVE_QTY = 0
                sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
                Select Case sts
                    Case BtNoErr
                        Call UniCode_Conv(GOODS_ONOREC.AVE_SYUKA, StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
                        AVE_QTY = CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(GOODS_ONOREC.AVE_SYUKA, "00000000")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�����Ϗo�א�")
                        Exit Function
                End Select
                                                        '���O���i����
                If AVE_QTY = 0 Then
                    Call UniCode_Conv(GOODS_ONOREC.SUMI_PERCENT, "00000000")
                Else
                    Call UniCode_Conv(GOODS_ONOREC.SUMI_PERCENT, Format(CLng(Sumi_QTY / AVE_QTY * 100), "00000000"))
                End If
                
                
                Do
                    
                    sts = BTRV(BtOpInsert, GOODS_ONO_POS, GOODS_ONOREC, Len(GOODS_ONOREC), K0_GOODS_ONO, Len(K0_GOODS_ONO), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<GOODS_ONO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpInsert, "���i���x���W�v�f�[�^")
                            Exit Function
                    End Select
                
                Loop
            End If
        End If
        
        com = BtOpGetNext
    Loop

    Data_Make_Proc = False


End Function


Private Function MI_ZAIKO_KENSAKU(Hinban As String) As Integer
'----------------------------------------------------------------------------
'                   �����i�̏���
'----------------------------------------------------------------------------
Dim i           As Integer
Dim Sumi_QTY    As Long
Dim Mi_QTY      As Long

Dim com         As Integer
Dim sts         As Integer

    MI_ZAIKO_KENSAKU = True
    
    For i = 0 To UBound(EE_ZAIKO_TBL)
        EE_ZAIKO_TBL(i).EE_LOC = ""
        EE_ZAIKO_TBL(i).EE_QTY = 0
    Next i
    
    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_ZAIKO.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Hinban)
    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, GOODS_OFF)
    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K1_ZAIKO.Soko_No, "")
    Call UniCode_Conv(K1_ZAIKO.Retu, "")
    Call UniCode_Conv(K1_ZAIKO.Ren, "")
    Call UniCode_Conv(K1_ZAIKO.Dan, "")
    
    com = BtOpGetGreater
    Do
        DoEvents
        
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
        Select Case sts
            Case BtNoErr
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                    Exit Do
                End If
                
                If StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> Hinban Then
                    Exit Do
                End If
                
                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) <> GOODS_OFF Then
                    Exit Do
                End If
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�݌Ƀf�[�^")
                Exit Function
        End Select
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, Hinban)
        
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                
            Case BtErrKeyNotFound
            
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�i��Ͻ�")
                Exit Function
        End Select
        
        
        
        If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)) = _
            (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
                
        Else
            For i = 0 To UBound(EE_ZAIKO_TBL)
                            
                If Trim(EE_ZAIKO_TBL(i).EE_LOC) = Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                    Exit For
                Else
                    If Len(Trim(EE_ZAIKO_TBL(i).EE_LOC)) = 0 Then
                        EE_ZAIKO_TBL(i).EE_LOC = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
                        Exit For
                    End If
                End If
            Next i
        
            If i > UBound(EE_ZAIKO_TBL) Then
                Exit Do
            End If
                
        
            EE_ZAIKO_TBL(i).EE_QTY = EE_ZAIKO_TBL(i).EE_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
        End If
    
        com = BtOpGetNext
    
    Loop
    
    MI_ZAIKO_KENSAKU = False

End Function
