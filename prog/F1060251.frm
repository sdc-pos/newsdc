VERSION 5.00
Begin VB.Form F1060251 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���i���v��x���A���[�����X�g���(���i���݌ɕ�)"
   ClientHeight    =   6948
   ClientLeft      =   2328
   ClientTop       =   2712
   ClientWidth     =   11292
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
   ScaleHeight     =   6948
   ScaleWidth      =   11292
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   4725
      MaxLength       =   3
      TabIndex        =   17
      Top             =   2400
      Width           =   480
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   4725
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   16
      Top             =   1800
      Width           =   855
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   4725
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   13
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�f�[�^"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "���O���i����"
      Height          =   255
      Index           =   1
      Left            =   2835
      TabIndex        =   19
      Top             =   2520
      Width           =   1770
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "%�܂�"
      Height          =   255
      Index           =   2
      Left            =   5355
      TabIndex        =   18
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�Ώۑq��"
      Height          =   255
      Index           =   0
      Left            =   3570
      TabIndex        =   15
      Top             =   1920
      Width           =   1050
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   255
      Index           =   33
      Left            =   3570
      TabIndex        =   14
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   12
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
Attribute VB_Name = "F1060251"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxSOKO% = 0                  '�J�n�@�W���I�ԁ@�q��
Private Const ptxSUMI_PERCENT% = 0          '���O���i����

Private Const Text_Max% = 0                 '��ʍ��ڕʍő���ޯ��

Private Const pcmbNaigai% = 0               '�����O
Private Const pcmbSoko% = 1                 '�����O


Private Const LMAX% = 40                    '�œ��ő�s��
Private Const LCTL% = 99                    '
Private Const MGN_L% = 10                   '���]���i�����F�P����j
Private Const MGN_U% = 1                    '��]���i�s���F�P����j

Private Pdate As String                     '����J�n���t�iͯ�ް�p�j
Private Ptime As String                     '����J�n�����iͯ�ް�p�j


Private NormalFont  As New StdFont          '����t�H���g
Private MidFont     As New StdFont          '����t�H���g

Private OutSide     As Long                 '����ΊO�o�א�

Private GOODS_DATA  As String               '�o�̓f�[�^�t�@�C����



Private SHO_SOKO    As Variant              '���i���p�q��

'�|�W�V���j���O
Private wZAIKO_POS  As POSBLK
'�f�[�^�E�o�b�t�@
Private wZAIKOREC   As ZAIKOREC_Tag
'�L�[�E�f�[�^
Private K0_wZAIKO   As KEY0_ZAIKO

Private SHIMUKE_CODE    As String * 2       '�d������R�[�h 2008.03.03

Private Const LAST_UPDATE_DAY$ = "[F106025] 2011.07.14 12:00"


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1060251.MousePointer = vbHourglass

    Call Ctrl_Lock(F1060251)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1060251)


    F1060251.MousePointer = vbDefault

End Sub


Private Sub Command_Click(Index As Integer)

Dim yn      As Integer
Dim i       As Integer
    
Dim mesg    As String
    
Dim c       As String * 128
    
    
    Select Case Index
        
        Case 7                              '�f�[�^�o��
            
            If Err_Chk() Then
                Exit Sub
            End If
                                        
                                        
                                        '�d�������荞��   2008.03.03
            If GetIni(App.EXEName, Last_JGYOBU, "SYS", c) Then
                MsgBox "�d������̐ݒ���s���Ă��������B"
                Exit Sub
            Else
                SHIMUKE_CODE = Trim(c)
            End If
            
            
            
            Beep
            yn = MsgBox("�f�[�^�o�͂��܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                
                If Data_Proc() Then
                    Unload Me
                End If
            End If
            
        
        
        Case 8                              '���
            
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            
                                        '�d�������荞��   2008.03.03
            If GetIni(App.EXEName, Last_JGYOBU, "SYS", c) Then
                MsgBox "�d������̐ݒ���s���Ă��������B"
                Exit Sub
            Else
                SHIMUKE_CODE = Trim(c)
            End If
            
            Beep
            yn = MsgBox("������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                
                If Print_Proc() Then
                    Unload Me
                End If
            End If
                    
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
            F1060251.Caption = "���i���v��x���A���[�����X�g���(���i���݌ɕ�)�i" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY
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
    If GetIni("FILE", "GOODS_S_DATA", "SYS", c) Then
        Beep
        MsgBox "'���i���x���t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    GOODS_DATA = Trim(c)

'-----------    SYS.INI ---> (����۸���ID).INI 2011.07.14
                                '�ΏۊO�o�א���荞��
    If GetIni(App.EXEName, "OUTSIDE", App.EXEName, c) Then
        OutSide = 0
    Else
        If IsNumeric(Trim(c)) Then
            OutSide = CLng(Trim(c))
        Else
            OutSide = 0
        End If
    End If
                                
                                
                                '���i���p�q�Ɏ�荞��
    If GetIni(App.EXEName, "SHO_SOKO", App.EXEName, c) Then
        c = " "
    End If
    SHO_SOKO = Split(Trim(c), ",", -1)
'-----------    SYS.INI ---> (����۸���ID).INI 2011.07.14
                                
                                
                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�\���}�X�^�n�o�d�m
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�t�@�C���n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�t�@�C���n�o�d�m
    If wZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�����Ϗo�א��n�o�d�m
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���i���W�v�t�@�C���n�o�d�m
    If GOODS_S_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '����t�H���g�ݒ�(�ʏ�)
    With NormalFont
        .NAME = F1060251.FontName
        .Size = 12
    End With

                                '����t�H���g�ݒ�i���j
    With MidFont
        .NAME = F1060251.FontName
        .Size = 8
    End With


    Combo(pcmbNaigai).Clear
    Combo(pcmbNaigai).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo(pcmbNaigai).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo(pcmbNaigai).ListIndex = 0

    
    Combo(pcmbSoko).Clear
    For i = 0 To UBound(SHO_SOKO)
        Combo(pcmbSoko).AddItem SHO_SOKO(i)
    Next i
    Combo(pcmbSoko).ListIndex = 0

    Show
    
    Combo(pcmbNaigai).SetFocus
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
                                            '�\���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
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
                                            
    If wZAIKO_CLOSE() Then
        Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
    End If
                                            '�����Ϗo�א��b�k�n�r�d
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�����Ϗo�א�")
        End If
    End If
                                            '���i���W�v�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, GOODS_S_POS, GOODS_SREC, Len(GOODS_SREC), K0_GOODS_S, Len(K0_GOODS_S), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���W�v�t�@�C��")
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1060251 = Nothing

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
    F1060251.Caption = "���i���v��x���A���[�����X�g���(���i���݌ɕ�)�i" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub
Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���i���x���A���[�����X�g�������
'----------------------------------------------------------------------------
Dim Lcnt        As Integer

Dim sts         As Integer
Dim com         As Integer

Dim Save_Soko   As String * 2

Dim Edit        As String

Dim X_Tab       As Integer

Dim wkSUMI_PERCENT      As Long
Dim SKIP_F              As Boolean


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
    
    If Trim(Text(ptxSUMI_PERCENT).Text) = "" Then
        wkSUMI_PERCENT = 100
    Else
        wkSUMI_PERCENT = CLng(Text(ptxSUMI_PERCENT).Text)
    End If
    
    
    
    Call UniCode_Conv(K1_GOODS_S.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_GOODS_S.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
    Call UniCode_Conv(K1_GOODS_S.Soko_No, Combo(pcmbSoko).Text)
    Call UniCode_Conv(K1_GOODS_S.SUMI_PERCENT, "")
    Call UniCode_Conv(K1_GOODS_S.HIN_GAI, "")
    
'    Call UniCode_Conv(K2_GOODS_S.JGYOBU, Last_JGYOBU)
'    Call UniCode_Conv(K2_GOODS_S.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
'    Call UniCode_Conv(K2_GOODS_S.Soko_No, Text(ptxSOKO).Text)
'    Call UniCode_Conv(K2_GOODS_S.AVE_SYUKA, "zzzzzzzz")
'    Call UniCode_Conv(K2_GOODS_S.Sumi_QTY, "")
'    Call UniCode_Conv(K2_GOODS_S.Mi_QTY, "zzzzzzzz")
'    Call UniCode_Conv(K2_GOODS_S.SUMI_PERCENT, "")
'    Call UniCode_Conv(K2_GOODS_S.HIN_GAI, "")
    
    
    
    com = BtOpGetGreater
    
    Do
        sts = BTRV(com, GOODS_S_POS, GOODS_SREC, Len(GOODS_SREC), K1_GOODS_S, Len(K1_GOODS_S), 1)
        Select Case sts
            Case BtNoErr
                
                If StrConv(GOODS_SREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(GOODS_SREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                    Exit Do
                End If
                                        
                If Trim(StrConv(GOODS_SREC.Soko_No, vbUnicode)) <> Trim(Combo(pcmbSoko).Text) Then
                    Exit Do
                End If
                
                
                SKIP_F = False
                If Not IsNumeric(StrConv(GOODS_SREC.SUMI_PERCENT, vbUnicode)) Then
                    SKIP_F = True
                Else
                    If CLng(StrConv(GOODS_SREC.SUMI_PERCENT, vbUnicode)) > wkSUMI_PERCENT Then
                        SKIP_F = True
                    End If
                End If
                
                If CLng(StrConv(GOODS_SREC.Mi_QTY, vbUnicode)) <= 0 Then
                    SKIP_F = True
                End If
                
                
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���i���W�v�t�@�C��")
                Exit Function
        End Select


'-------------------------------------------------  ���׈��
        
        If Not SKIP_F Then
        
            If com = BtOpGetGreater Then
                Save_Soko = StrConv(GOODS_SREC.Soko_No, vbUnicode)
                
                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
    '                    If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
    '                        Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
    '                    End If
                    Case BtErrKeyNotFound
                        '�l�����Ȃ��������͌p��
                        Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
    '                    Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                        Exit Function
                End Select
                
            End If
            
            If Save_Soko <> StrConv(GOODS_SREC.Soko_No, vbUnicode) Then
                                
                Lcnt = LMAX + 1
                Save_Soko = StrConv(GOODS_SREC.Soko_No, vbUnicode)
                
                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
    '                    If Not IsNumeric(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
    '                        Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
    '                    End If
                    
                    Case BtErrKeyNotFound
                            '�l�����Ȃ��������͌p��
                        Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
    '                    Call UniCode_Conv(SOKOREC.ORDER_POINT, "000")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�q�Ƀ}�X�^")
                        Exit Function
                End Select
                
            End If
            
            
    '        If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
    '                        '�ݒ蔭���_���傫��
    '            Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "99999999")
    '            Call UniCode_Conv(K0_GOODS.HIN_GAI, "zzzzzzzzzzzzz")
    '            com = BtOpGetGreater
    '        Else
                '�����i�݌Ɂ��O �́A����ΏۊO 2004.08.27
    '            If OutSide >= CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) Or _
    '                CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) <= 0 Then
                
    '            If CLng(StrConv(GOODS_SREC.Mi_QTY, vbUnicode)) <= 0 Then
    '            Else
                    If Head_Print_Proc(Lcnt) Then
                        Exit Function
                    End If
                
                    X_Tab = MGN_L
                
                    Printer.Print Tab(X_Tab);
                                                            '�W���I��
                    Edit = StrConv(GOODS_SREC.ST_SOKO, vbUnicode) & "-"
                    Edit = Edit & StrConv(GOODS_SREC.ST_RETU, vbUnicode) & "-"
                    Edit = Edit & StrConv(GOODS_SREC.ST_REN, vbUnicode) & "-"
                    Edit = Edit & StrConv(GOODS_SREC.ST_DAN, vbUnicode)
                    Printer.Print Edit;
                    X_Tab = X_Tab + Len(Edit) + 5
                                                            '�i�ԁi�O���j
                    Printer.Print Tab(X_Tab);
    
                    Printer.Print StrConv(GOODS_SREC.HIN_GAI, vbUnicode);
                    X_Tab = X_Tab + Len(StrConv(GOODS_SREC.HIN_GAI, vbUnicode)) + 4
                                                            '����
                    Printer.Print Tab(X_Tab);
''2008.11.06                    Printer.Print StrConv(GOODS_SREC.PACKING_NO, vbUnicode);
''2008.11.06                    X_Tab = X_Tab + Len(StrConv(GOODS_SREC.PACKING_NO, vbUnicode)) + 10
                                                            
                                                            
                                                            
                    '2008.11.06
                    Printer.Print Left(StrConv(GOODS_SREC.KOSOU, vbUnicode), 4);
                    X_Tab = X_Tab + Len(StrConv(GOODS_SREC.PACKING_NO, vbUnicode)) + 4
                    '2008.11.06
                                                            
                                                            
                                                            
                                                            '���i���q�ɍ݌ɐ�
                    Printer.Print Tab(X_Tab);
                    Edit = Format(CLng(StrConv(GOODS_SREC.SOKO_QTY, vbUnicode)), "#,##0")
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Printer.Print Edit;
                    X_Tab = X_Tab + Len(Edit) + 2
                                                            
                                                            
                                                            '���i���ςݍ݌ɐ�
                    Printer.Print Tab(X_Tab);
                    Edit = Format(CLng(StrConv(GOODS_SREC.Sumi_QTY, vbUnicode)), "#,##0")
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Printer.Print Edit;
                    X_Tab = X_Tab + Len(Edit) + 2
                                                            '�����i�݌ɐ�
                    Printer.Print Tab(X_Tab);
                    Edit = Format(CLng(StrConv(GOODS_SREC.Mi_QTY, vbUnicode)), "#,##0")
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Printer.Print Edit;
                    X_Tab = X_Tab + Len(Edit) + 2
                                                            '�����Ϗo�א�
                    Printer.Print Tab(X_Tab);
                    Edit = Format(CLng(StrConv(GOODS_SREC.AVE_SYUKA, vbUnicode)), "#,##0")
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Printer.Print Edit;
    '                X_Tab = X_Tab + Len(Edit) + 8
                    X_Tab = X_Tab + Len(Edit) + 2
                                                            '���O���i���K�v��
                    Printer.Print Tab(X_Tab);
                    Edit = Format(CLng(StrConv(GOODS_SREC.AVE_SYUKA, vbUnicode)) - CLng(StrConv(GOODS_SREC.Sumi_QTY, vbUnicode)), "#,##0")
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Printer.Print Edit;
    '                X_Tab = X_Tab + Len(Edit) + 8
                    X_Tab = X_Tab + Len(Edit) + 2
                                                            '���O���i����
                    Printer.Print Tab(X_Tab);
                    Edit = Format(CLng(StrConv(GOODS_SREC.SUMI_PERCENT, vbUnicode)), "#0") & "%"
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
    
                    Printer.Print Edit
    '                X_Tab = X_Tab + Len(Edit) + 5
    '                                                        '�ʒu�݌�
    '                Printer.Print Tab(X_Tab);
    '
    '                If MI_ZAIKO_KENSAKU(StrConv(GOODSREC.HIN_GAI, vbUnicode)) Then
    '                    Exit Function
    '                End If
    '
    '                If Len(Trim(EE_ZAIKO_TBL(0).EE_LOC)) <> 0 Then
    '                    Edit = Format(EE_ZAIKO_TBL(0).EE_QTY, "#0")
    '                    If Len(Edit) < 9 Then
    '                        Edit = Space(9 - Len(Edit)) & Edit
    '                    End If
    '                    Edit = Left(EE_ZAIKO_TBL(0).EE_LOC, 2) & "-" & _
    '                           Mid(EE_ZAIKO_TBL(0).EE_LOC, 3, 2) & "-" & _
    '                           Mid(EE_ZAIKO_TBL(0).EE_LOC, 5, 2) & "-" & _
    '                           Right(EE_ZAIKO_TBL(0).EE_LOC, 2) & Edit
    '                End If
    
    '                Printer.Print Edit
    
                    Printer.Print
                
                    Lcnt = Lcnt + 2
            
            '    End If
            End If
            com = BtOpGetNext
'        End If
    Loop

    Printer.EndDoc


    Call Input_UnLock
    
    Print_Proc = False

End Function

Private Function Head_Print_Proc(Lcnt As Integer) As Integer

Dim sts     As Integer
Dim i       As Integer

    Head_Print_Proc = True
    
    If Lcnt < LMAX Then
        
        Head_Print_Proc = False
        Exit Function
    
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

    Printer.Print Tab(MGN_L + 35);
    
    Printer.Print "���i���x���A���[�����X�g�i���i���p�q�ɍ݌ɕ��j";
    
    
    Printer.Print Tab(MGN_L + 90);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print

    
    Printer.Print Tab(MGN_L);
    Printer.Print "�q�ɁF";
    Printer.Print StrConv(SOKOREC.Soko_No, vbUnicode) & "  ";
    Printer.Print StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "  ";
'    Printer.Print "�i�ݒ蔭���_ " & StrConv(Format(CLng(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0"), vbWide) & "���j"
    Printer.Print


    Printer.Print Tab(MGN_L);
    Printer.Print "�W���I��";
    Printer.Print Tab(MGN_L + 16);
    Printer.Print "�i�ԁi�O���j";
    Printer.Print Tab(MGN_L + 40);
    Printer.Print "����";
    
    Printer.Print Tab(MGN_L + 52);
    Printer.Print "�Y���q�ɍ݌�";
    
    
    Printer.Print Tab(MGN_L + 70);
    Printer.Print "���ϐ�";
    Printer.Print Tab(MGN_L + 82);
    Printer.Print "�����i";
    Printer.Print Tab(MGN_L + 94);
    Printer.Print "������";
    Printer.Print Tab(MGN_L + 106);
    Printer.Print "�K�v��";
    Printer.Print Tab(MGN_L + 118);
    Printer.Print "�@��";
'    Printer.Print Tab(MGN_L + 113);
'    Printer.Print "�ʒu�݌�"

    Printer.Print
    Printer.Print

    Lcnt = 0
    
    Head_Print_Proc = False

End Function

Private Function Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                   �x���p�W�v�f�[�^�쐬����
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim Upd_com         As Integer

Dim ans             As Integer

Dim Sumi_QTY        As Long
Dim Mi_QTY          As Long
Dim AVE_QTY         As Long

Dim ALL_SUMI_QTY    As Long
Dim SHO_SUMI_QTY    As Long
Dim ALL_MI_QTY      As Long
Dim SHO_MI_QTY      As Long

Dim i               As Integer
Dim j               As Integer


Dim Skip_Flg    As Boolean

    Data_Make_Proc = True

'---------------------------------------------------------- '�S���R�[�h�폜
    com = BtOpGetFirst
    Do
        
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, GOODS_S_POS, GOODS_SREC, Len(GOODS_SREC), K0_GOODS_S, Len(K0_GOODS_S), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<GOODS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
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
            
            sts = BTRV(BtOpDelete, GOODS_S_POS, GOODS_SREC, Len(GOODS_SREC), K0_GOODS_S, Len(K0_GOODS_S), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<GOODS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
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
'---------------------------------------------------------- '�݌Ƀf�[�^�x�[�X�Ńf�[�^�쐬
    
    Call UniCode_Conv(K0_wZAIKO.Soko_No, Combo(pcmbSoko).Text)
    Call UniCode_Conv(K0_wZAIKO.Retu, "01")
    Call UniCode_Conv(K0_wZAIKO.Ren, "01")
    Call UniCode_Conv(K0_wZAIKO.Dan, "01")
    Call UniCode_Conv(K0_wZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_wZAIKO.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
    Call UniCode_Conv(K0_wZAIKO.HIN_GAI, "")
    Call UniCode_Conv(K0_wZAIKO.NYUKA_DT, "")

    com = BtOpGetGreater

    Do
        DoEvents
        sts = BTRV(com, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
        Select Case sts
            Case BtNoErr
                
                If StrConv(wZAIKOREC.Soko_No, vbUnicode) <> Combo(pcmbSoko).Text Or _
                    StrConv(wZAIKOREC.Retu, vbUnicode) <> "01" Or _
                    StrConv(wZAIKOREC.Ren, vbUnicode) <> "01" Or _
                    StrConv(wZAIKOREC.Dan, vbUnicode) <> "01" Then
                    '�q�Ƀu���[�N
                    Exit Do
                End If
                                    
                If StrConv(wZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(wZAIKOREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                    '���ƕ��^�����O��ڰ�
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�݌��ް�")
                Exit Function
        End Select


        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(wZAIKOREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(wZAIKOREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(wZAIKOREC.HIN_GAI, vbUnicode))

        Skip_Flg = False

        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Skip_Flg = True
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                Exit Function
        End Select


        If Not Skip_Flg Then
            If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
                                                        '���ƕ�
                Call UniCode_Conv(K0_GOODS_S.JGYOBU, Last_JGYOBU)
                                                        '�����O
                Call UniCode_Conv(K0_GOODS_S.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
                                                        '�i�ԁi�O���j
                Call UniCode_Conv(K0_GOODS_S.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                                            
                                                        
                Do
                    sts = BTRV(BtOpGetEqual + BtSNoWait, GOODS_S_POS, GOODS_SREC, Len(GOODS_SREC), K0_GOODS_S, Len(K0_GOODS_S), 0)
                    Select Case sts
                        Case BtNoErr
                            
                            Upd_com = BtOpUpdate
                            Exit Do
                        Case BtErrKeyNotFound
                            Upd_com = BtOpInsert
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<GOODS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual + BtSNoWait, "���i���x���W�v�f�[�^")
                            Exit Function
                    End Select
                                                        
                Loop
                
                If Upd_com = BtOpInsert Then
                                                        
                    Call UniCode_Conv(GOODS_SREC.JGYOBU, StrConv(wZAIKOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(GOODS_SREC.NAIGAI, StrConv(wZAIKOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(GOODS_SREC.HIN_GAI, StrConv(wZAIKOREC.HIN_GAI, vbUnicode))
                                                                            
                    Call UniCode_Conv(GOODS_SREC.Soko_No, Combo(pcmbSoko).Text)
                                                        '�W���I��
                    Call UniCode_Conv(GOODS_SREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
                    Call UniCode_Conv(GOODS_SREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
                    Call UniCode_Conv(GOODS_SREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
                    Call UniCode_Conv(GOODS_SREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                                        
                                                        '����
                    Call UniCode_Conv(GOODS_SREC.PACKING_NO, StrConv(ITEMREC.PACKING_NO, vbUnicode))
                End If
                                                    '�Ώۑq�ɂ̍݌�
                                                                    
                If Zaiko_Syukei_Proc(SHO_SUMI_QTY, _
                                        SHO_MI_QTY, _
                                        Last_JGYOBU, _
                                        Right(Combo(pcmbNaigai).Text, 1), _
                                        StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                        Combo(pcmbSoko).Text & "01" & "01" & "01") = SYS_ERR Then
                    Exit Function
                End If
                Call UniCode_Conv(GOODS_SREC.SOKO_QTY, Format(SHO_SUMI_QTY, "00000000"))
                                                    '�݌ɏW�v����
                If Zaiko_Syukei_Proc(ALL_SUMI_QTY, _
                                        ALL_MI_QTY, _
                                        Last_JGYOBU, _
                                        Right(Combo(pcmbNaigai).Text, 1), _
                                        StrConv(ITEMREC.HIN_GAI, vbUnicode)) = SYS_ERR Then
                    Exit Function
                End If
                                                    
                                                    '���i���p�q�ɂ̍݌�
                
                For j = 0 To UBound(SHO_SOKO)
                    If Zaiko_Syukei_Proc(SHO_SUMI_QTY, _
                                            SHO_MI_QTY, _
                                            Last_JGYOBU, _
                                            Right(Combo(pcmbNaigai).Text, 1), _
                                            StrConv(ITEMREC.HIN_GAI, vbUnicode), _
                                            SHO_SOKO(j) & "01" & "01" & "01") = SYS_ERR Then
                        Exit Function
                    End If
                    ALL_SUMI_QTY = ALL_SUMI_QTY - SHO_SUMI_QTY
                Next j
                                                    '���i���ςݍ݌ɐ�(���i�����݌Ɉ���)
                Call UniCode_Conv(GOODS_SREC.Sumi_QTY, Format(ALL_SUMI_QTY, "00000000"))
                                                    '�����i�݌ɐ�
                Call UniCode_Conv(GOODS_SREC.Mi_QTY, Format(ALL_MI_QTY, "00000000"))
                                                    '�����Ϗo�א�
                Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
                Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                
                AVE_QTY = 0
                sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
                Select Case sts
                    Case BtNoErr
                        Call UniCode_Conv(GOODS_SREC.AVE_SYUKA, StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
                        AVE_QTY = CLng(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(GOODS_SREC.AVE_SYUKA, "00000000")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�����Ϗo�א�")
                        Exit Function
                End Select
                                                    '���O���i����
                If AVE_QTY = 0 Then
                    Call UniCode_Conv(GOODS_SREC.SUMI_PERCENT, "00000000")
                Else
                    Call UniCode_Conv(GOODS_SREC.SUMI_PERCENT, Format(CLng(ALL_SUMI_QTY / AVE_QTY * 100), "00000000"))
                End If
                
                
                
                
                '�����ݒ�
                Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, SHIMUKE_CODE)
                Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_KOSOU)
                Call UniCode_Conv(K0_P_COMPO.SEQNO, "010")
                sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                Select Case sts
                    Case BtNoErr
                        Call UniCode_Conv(GOODS_SREC.KOSOU, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(GOODS_SREC.KOSOU, "")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                        Exit Function
                End Select
                
                
                '�O�����ݒ�
                Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, SHIMUKE_CODE)
                Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_GAISOU)
                Call UniCode_Conv(K0_P_COMPO.SEQNO, "010")
                sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                Select Case sts
                    Case BtNoErr
                        Call UniCode_Conv(GOODS_SREC.GAISOU, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(GOODS_SREC.GAISOU, "")
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                        Exit Function
                End Select
                
                
                
                
                
                
                
                
                
                
                
                
                Do
                    
                    sts = BTRV(Upd_com, GOODS_S_POS, GOODS_SREC, Len(GOODS_SREC), K0_GOODS_S, Len(K0_GOODS_S), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<GOODS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, Upd_com, "���i���x���W�v�f�[�^")
                            Exit Function
                    End Select
                
                Loop
            End If
        
        End If

        com = BtOpGetNext


    Loop
    
        

    Data_Make_Proc = False


End Function

Private Function Data_Proc() As Integer
'----------------------------------------------------------------------------
'                   �b�r�u�f�[�^�쐬����
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim Save_Soko       As String * 2

Dim Edit            As String

Dim FileNo          As Integer
Dim fileName        As String
    
Dim wkSUMI_PERCENT      As Long
Dim SKIP_F              As Boolean
Dim FSW                 As Boolean
    
    
    Data_Proc = True

    Call Input_Lock

    fileName = GOODS_DATA
    sts = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), sts) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - sts)
    
    On Error GoTo Error_Proc
    
    FileNo = FreeFile
    Open (fileName) For Output As FileNo


    If Data_Make_Proc() Then        '���i���x���W�v�f�[�^�쐬
        Exit Function
    End If
    
    If Trim(Text(ptxSUMI_PERCENT).Text) = "" Then
        wkSUMI_PERCENT = 100
    Else
        wkSUMI_PERCENT = CLng(Text(ptxSUMI_PERCENT).Text)
    End If
    FSW = True
    
    Call UniCode_Conv(K1_GOODS_S.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K1_GOODS_S.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
    Call UniCode_Conv(K1_GOODS_S.Soko_No, Combo(pcmbSoko).Text)
    Call UniCode_Conv(K1_GOODS_S.SUMI_PERCENT, "")
    Call UniCode_Conv(K1_GOODS_S.HIN_GAI, "")
     
'    Call UniCode_Conv(K2_GOODS_S.JGYOBU, Last_JGYOBU)
'    Call UniCode_Conv(K2_GOODS_S.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
'    Call UniCode_Conv(K2_GOODS_S.Soko_No, Text(ptxSOKO).Text)
'    Call UniCode_Conv(K2_GOODS_S.AVE_SYUKA, "zzzzzzzz")
'    Call UniCode_Conv(K2_GOODS_S.Sumi_QTY, "")
'    Call UniCode_Conv(K2_GOODS_S.Mi_QTY, "zzzzzzzz")
'    Call UniCode_Conv(K2_GOODS_S.SUMI_PERCENT, "")
'    Call UniCode_Conv(K2_GOODS_S.HIN_GAI, "")
   
    com = BtOpGetGreater
    
    Do
        DoEvents
        sts = BTRV(com, GOODS_S_POS, GOODS_SREC, Len(GOODS_SREC), K1_GOODS_S, Len(K1_GOODS_S), 1)
        Select Case sts
            Case BtNoErr
                If StrConv(GOODS_SREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(GOODS_SREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                    Exit Do
                End If
                                        
                If Trim(StrConv(GOODS_SREC.Soko_No, vbUnicode)) <> Trim(Combo(pcmbSoko).Text) Then
                    Exit Do
                End If
                
                SKIP_F = False
                If Not IsNumeric(StrConv(GOODS_SREC.SUMI_PERCENT, vbUnicode)) Then
                    SKIP_F = True
                Else
                    If CLng(StrConv(GOODS_SREC.SUMI_PERCENT, vbUnicode)) > wkSUMI_PERCENT Then
                        SKIP_F = True
                    End If
                End If
                
                If CLng(StrConv(GOODS_SREC.Mi_QTY, vbUnicode)) <= 0 Then
                    SKIP_F = True
                End If
                
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���i���W�v�t�@�C��")
                Exit Function
        End Select
'-------------------------------------------------  ���׈��
        If Not SKIP_F Then
            If FSW Then
                FSW = False
                Save_Soko = StrConv(GOODS_SREC.Soko_No, vbUnicode)
                
                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
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
                        '�w�b�_�[�o��
                Write #FileNo, "*** ���i���x���A���[�����X�g�@***",
                Write #FileNo, "�쐬���t:" & Format(Now, "YYYY/MM/DD") & " " & Format(Now, "HH:MM:SS")
                        
            
                Write #FileNo, "�W���I��", "�i�ԁi�O���j", "���ށi�����j", "�Y���q�ɍ݌�", "���i���ύ݌�", "�����i�݌�", "�����Ϗo�א�", "���O���i���K�v��", "���O���i����"
                
            
                Write #FileNo, "�q�ɇ��F" & StrConv(SOKOREC.Soko_No, vbUnicode) & " " & StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                
            
            
            End If
            
            If Save_Soko <> StrConv(GOODS_SREC.Soko_No, vbUnicode) Then
                                
                Save_Soko = StrConv(GOODS_SREC.Soko_No, vbUnicode)
                
                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
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
                
                Write #FileNo, "�q�ɇ��F" & StrConv(SOKOREC.Soko_No, vbUnicode) & " " & StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                
                
            End If
            
                                                    '�W���I��
            Edit = StrConv(GOODS_SREC.ST_SOKO, vbUnicode) & "-"
            Edit = Edit & StrConv(GOODS_SREC.ST_RETU, vbUnicode) & "-"
            Edit = Edit & StrConv(GOODS_SREC.ST_REN, vbUnicode) & "-"
            Edit = Edit & StrConv(GOODS_SREC.ST_DAN, vbUnicode)
            Write #FileNo, Edit,
                                                    '�i�ԁi�O���j
            Write #FileNo, StrConv(GOODS_SREC.HIN_GAI, vbUnicode),
                                                    '����
            'Write #FileNo, Trim(StrConv(GOODSREC.PACKING_NO, vbUnicode)),      '2008.03.03
            Write #FileNo, Trim(StrConv(GOODS_SREC.KOSOU, vbUnicode)),            '2008.03.03
                                                    '�Y���q�ɍ݌ɐ�
            Edit = Format(CLng(StrConv(GOODS_SREC.SOKO_QTY, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Write #FileNo, Edit,
                                                    '���i���ςݍ݌ɐ�
            Edit = Format(CLng(StrConv(GOODS_SREC.Sumi_QTY, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Write #FileNo, Edit,
                                                    '�����i�݌ɐ�
            Edit = Format(CLng(StrConv(GOODS_SREC.Mi_QTY, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Write #FileNo, Edit,
                                                    '�����Ϗo�א�
            Edit = Format(CLng(StrConv(GOODS_SREC.AVE_SYUKA, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Write #FileNo, Edit,
                                                    '���O���i���K�v��
            Edit = Format(CLng(StrConv(GOODS_SREC.AVE_SYUKA, vbUnicode)) - CLng(StrConv(GOODS_SREC.Sumi_QTY, vbUnicode)), "#,##0")
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Write #FileNo, Edit,
                                                    '���O���i����
            Edit = Format(CInt(StrConv(GOODS_SREC.SUMI_PERCENT, vbUnicode)), "#0") & "%"
            If Len(Edit) < 10 Then
                Edit = Space(10 - Len(Edit)) & Edit
            End If
            Write #FileNo, Edit
        End If
        com = BtOpGetNext
    Loop

    Close #FileNo

    Beep
    MsgBox "�u" & fileName & "�v�͐���ɏo�͂���܂����B"

    Call Input_UnLock
    
    Data_Proc = False
    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "���g�p���ł��B"
        Call Input_UnLock         '��ʍ��ڃ��b�N����
        Data_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        Data_Proc = True
    End If


End Function

'Private Function MI_ZAIKO_KENSAKU(Hinban As String) As Integer
'----------------------------------------------------------------------------
'                   �����i�̏���
'----------------------------------------------------------------------------
'Dim i           As Integer
'Dim Sumi_QTY    As Long
'Dim Mi_QTY      As Long
'
'Dim com         As Integer
'Dim sts         As Integer
'
'    MI_ZAIKO_KENSAKU = True
'
'    For i = 0 To UBound(EE_ZAIKO_TBL)
'        EE_ZAIKO_TBL(i).EE_LOC = ""
'        EE_ZAIKO_TBL(i).EE_QTY = 0
'    Next i
'
'    Call UniCode_Conv(K1_ZAIKO.JGYOBU, Last_JGYOBU)
'    Call UniCode_Conv(K1_ZAIKO.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
'    Call UniCode_Conv(K1_ZAIKO.HIN_GAI, Hinban)
'    Call UniCode_Conv(K1_ZAIKO.GOODS_ON, GOODS_OFF)
'    Call UniCode_Conv(K1_ZAIKO.NYUKA_DT, "")
'    Call UniCode_Conv(K1_ZAIKO.SOKO_NO, "")
'    Call UniCode_Conv(K1_ZAIKO.Retu, "")
'    Call UniCode_Conv(K1_ZAIKO.Ren, "")
'    Call UniCode_Conv(K1_ZAIKO.Dan, "")
'
'    com = BtOpGetGreater
'    Do
'        DoEvents
'
'        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K1_ZAIKO, Len(K1_ZAIKO), 1)
'        Select Case sts
'            Case BtNoErr
'                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
'                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
'                    Exit Do
'                End If
'
'                If StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> Hinban Then
'                    Exit Do
'                End If
'
'                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) <> GOODS_OFF Then
'                    Exit Do
'                End If
'
'            Case BtErrEOF
'                Exit Do
'            Case Else
'                Call File_Error(sts, com, "�݌Ƀf�[�^")
'                Exit Function
'        End Select
'        For i = 0 To UBound(EE_ZAIKO_TBL)
'
'            If Trim(EE_ZAIKO_TBL(i).EE_LOC) = Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
'                Exit For
'            Else
'                If Len(Trim(EE_ZAIKO_TBL(i).EE_LOC)) = 0 Then
'                    EE_ZAIKO_TBL(i).EE_LOC = StrConv(ZAIKOREC.SOKO_NO, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)
'                    Exit For
'                End If
'            End If
'        Next i
'
'        If i > UBound(EE_ZAIKO_TBL) Then
'            Exit Do
'        End If
'
'
'        EE_ZAIKO_TBL(i).EE_QTY = EE_ZAIKO_TBL(i).EE_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
'
'
'        com = BtOpGetNext
'
'    Loop
'
'    MI_ZAIKO_KENSAKU = False
'
'End Function
Public Function wZAIKO_Open(Mode As Integer) As Integer
'****************************************************
'*      �u�ړ������v    �݌ɂn�o�d�m����
'*
'*  �݌Ƀt�@�C����ʃ|�C���^�łn�o�d�m����
'*  (�Ăь��ŋN�����ɂP�x�����Ăяo��)

'*  �߂�l: false       :����
'*          true        :�ُ�
'*          SYS_CANCEL  :�X�V��ݾ�
'****************************************************
Dim sts As Integer
Dim c As String * 128
Dim FullPath As String

Dim ans         As Integer
    
    
    wZAIKO_Open = True
                                '�݌Ƀf�[�^�@�t���p�X�捞��
    sts = GetIni("FILE", ZAIKO_ID, "SYS", c)
    
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI �ǂݍ��݃G���[")
        Exit Function
    End If
    
    FullPath = RTrim(c)
    
    Do
        sts = BTRV(BtOpOpen, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), ByVal FullPath, Len(FullPath), Mode)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrFILE_INUSE, BtErrINCOMPATIBLE_MODE_ERROR
'-------------- �n�o�d�m�����ł̎g�p���́A�����グ���ɂP�񂾂��̂͂��Ȃ̂ŁA��ɉ�ʓ��͂Ƃ��A
'               ��ݾق́A�����̋N����ݾقƂ���B
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    wZAIKO_Open = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpOpen, "�݌Ƀf�[�^")
                Exit Function
        End Select
    Loop

    wZAIKO_Open = False

End Function

Public Function wZAIKO_CLOSE() As Integer

'****************************************************
'*      �u�ړ������v    �݌ɂb�k�n�r�d����
'*
'*  �݌Ƀt�@�C����ʃ|�C���^�łb�k�n�r�d����
'*  (�Ăь��ŏI�����ɂP�x�����Ăяo��)
'*  �߂�l: false       :����
'*          true        :�ُ�
'****************************************************
Dim sts As Integer
    
    wZAIKO_CLOSE = True
    
    sts = BTRV(BtOpClose, wZAIKO_POS, wZAIKOREC, Len(wZAIKOREC), K0_wZAIKO, Len(K0_wZAIKO), 0)
    
    Select Case sts
        Case BtNoErr, BtErrNoOpen
        Case Else
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
            Exit Function
    End Select

    wZAIKO_CLOSE = False

End Function

Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   �G���[�`�F�b�N����
'----------------------------------------------------------------------------
                                            
Dim i   As Integer
Dim sts As Integer

                                            
    Err_Chk = True
            
    
    If Trim(Text(ptxSUMI_PERCENT).Text) = "" Then
    Else
        If Not IsNumeric(Text(ptxSUMI_PERCENT).Text) Then
            MsgBox "���͂������ڂ̓G���[�ł��B�i���O���i���󋵁i���j�j"
            Text(ptxSUMI_PERCENT).SetFocus
            Err_Chk = False
            Exit Function
        End If
    End If
    
    
    
    Err_Chk = False

End Function

Private Sub Text_GotFocus(Index As Integer)
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If

End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i   As Integer
Dim sts As Integer

    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
    
        Case ptxSUMI_PERCENT    '2008.03.03
    
            If Trim(Text(ptxSUMI_PERCENT).Text) = "" Then
            Else
                If Not IsNumeric(Text(ptxSUMI_PERCENT).Text) Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B�i���O���i���󋵁i���j�j"
                    Text(ptxSUMI_PERCENT).SetFocus
                    Exit Sub
                End If
            End If
    End Select
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i

End Sub
