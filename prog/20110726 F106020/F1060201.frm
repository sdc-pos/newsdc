VERSION 5.00
Begin VB.Form F1060201 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���i���v��x���A���[�����X�g���"
   ClientHeight    =   7128
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
   ScaleHeight     =   7128
   ScaleWidth      =   11292
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command1 
      Caption         =   "INI�\��"
      Height          =   372
      Left            =   9600
      TabIndex        =   26
      Top             =   120
      Width           =   1332
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   4680
      MaxLength       =   3
      TabIndex        =   22
      Top             =   2400
      Width           =   480
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   3060
      MaxLength       =   3
      TabIndex        =   20
      Top             =   2400
      Width           =   480
   End
   Begin VB.ComboBox Combo 
      Height          =   336
      Index           =   0
      Left            =   3120
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   0
      Top             =   1320
      Width           =   375
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
      TabIndex        =   12
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      Index           =   5
      Left            =   4800
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
      Index           =   2
      Left            =   1800
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
      Index           =   1
      Left            =   960
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "���o���O�����@OEM�i��(�o�׋敪ZZ) ���i���v�揜�O�׸ށF�P"
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   25
      Top             =   4320
      Width           =   6720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "���o�Ώۏ����@���������敪 1:�Ώہ@2:�Ő؈ē��� ��"
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   24
      Top             =   3840
      Width           =   6240
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "%"
      Height          =   252
      Index           =   3
      Left            =   5160
      TabIndex        =   23
      Top             =   2520
      Width           =   372
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "%�@�`"
      Height          =   240
      Index           =   2
      Left            =   3600
      TabIndex        =   21
      Top             =   2520
      Width           =   960
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "���O���i����"
      Height          =   252
      Index           =   1
      Left            =   1176
      TabIndex        =   19
      Top             =   2520
      Width           =   1776
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   "�i���󔒁F�S�q�Ɏw��j"
      Height          =   252
      Left            =   3000
      TabIndex        =   18
      Top             =   1920
      Width           =   2652
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Height          =   252
      Left            =   3600
      TabIndex        =   17
      Top             =   1440
      Width           =   2412
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   252
      Index           =   33
      Left            =   2280
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "�W���I��(�q�ɔԍ�)"
      Height          =   252
      Index           =   0
      Left            =   840
      TabIndex        =   14
      Top             =   1440
      Width           =   2292
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
      TabIndex        =   13
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
Attribute VB_Name = "F1060201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxSOKO% = 0                  '�J�n�@�W���I�ԁ@�q��
Private Const ptxFROM_SUMI_PERCENT% = 1     '���O���i���󋵂���     2011.07.04
Private Const ptxTO_SUMI_PERCENT% = 2       '���O���i���󋵂܂�     2011.07.04



Private Const Text_Max% = 2                 '��ʍ��ڕʍő���ޯ��

Private Const pcmbNaigai% = 0               '�����O


Private Const LMAX% = 36                    '�œ��ő�s��
Private Const LCTL% = 99                    '
Private Const MGN_L% = 3                    '���]���i�����F�P����j
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


Private SHIMUKE_CODE    As String * 2       '�d������R�[�h 2008.03.03


Private SORT_SEQ        As Integer          'SORT�� 2008.11.06


'''''''''''''''''''''''''''''''''''''''''''''   2011.03.31
Private Type KOUSEI_TBL
    KO_JGYOBU   As String * 1           '���ƕ�
    KO_NAIGAI   As String * 1           '�����O
    KO_SYUBETSU As String * 2           '���
    KO_HIN_GAI  As String * 20          '�i��
    KO_QTY      As Double               '����
    G_ST_SHITAN As Double               '�d����
    G_ST_URITAN As Double               '���し
    G_ST_SHIKIN As Double               '�d�����z
    G_ST_URIKIN As Double               '������z
    S_KOUSU     As Double               '��Ǝ���
    SEI_SYU_KON As Double               '�W������
    G_ST_URIKIN_KUSATU As _
                    Double              '���Ð�p
End Type




Dim SHIZAI_T                As Variant      '���ޑΏ�
Dim DOUKON_T        As Variant              '�����Ώ�
Dim KAKOU_T         As Variant              '���H�Ώ�

Dim KUSATU_F                As Boolean      '�ΏۃZ���^�[�@���� OR ���ÈȊO


Dim KOSOU_KBN       As String * 2       '���敪
Dim GAISO_KBN       As String * 2       '�O���敪

'''''''''''''''''''''''''''''''''''''''''''''   2011.03.31



'''''''''''''''''''''''''''''''''''''''''''''   2011.07.04
Dim SAMPLE_QTY      As Integer          '���{���O��
Dim NOT_Hin_Name    As Variant          '���O�i��
Dim NOT_Hin_Name_F  As Boolean          '���O�i���L��
Dim wkNOT_Hin_Name  As String

Dim TUKI1_TITLE     As String           '�����Ϗo�א�����
Dim S_TUKI1_TITLE   As String           '���Y�v��p�����Ϗo�א�����(1)
Dim S_TUKI2_TITLE   As String           '���Y�v��p�����Ϗo�א�����(1)
Dim TUKI1           As Integer
Dim TUKI2           As Integer
Dim TUKI3           As Integer
'''''''''''''''''''''''''''''''''''''''''''''   2011.07.04


Private Const LAST_UPDATE_DAY$ = "[F106020] 2011.07.12 09:00"



Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   �G���[�`�F�b�N����
'----------------------------------------------------------------------------
                                            
Dim i   As Integer
Dim sts As Integer

                                            
    Err_Chk = True
            
    If Len(Trim(Text(ptxSOKO).Text)) = 0 Then
        Label2.Caption = "�S�q��"
    Else
        Call UniCode_Conv(K0_SOKO.Soko_No, Text(ptxSOKO).Text)
            
        sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
                Label2.Caption = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                        
                If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_KASO Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B�i���z�q�Ɂj"
                    Text(ptxSOKO).SetFocus
                    Err_Chk = False
                    Exit Function
                End If
                    
            Case BtErrKeyNotFound
                Beep
                MsgBox "���͂������ڂ̓G���[�ł��B�i�q�ɖ��o�^�j"
                Text(ptxSOKO).SetFocus
                Err_Chk = False
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetGreaterEqual, "�q�Ƀ}�X�^")
                Exit Function
        End Select
            
            
    End If
    
    If Trim(Text(ptxFROM_SUMI_PERCENT).Text) = "" Then
        Text(ptxFROM_SUMI_PERCENT).Text = "000"
    Else
        If Not IsNumeric(Text(ptxFROM_SUMI_PERCENT).Text) Then
            MsgBox "���͂������ڂ̓G���[�ł��B�i���O���i���󋵁i���j�j"
            Text(ptxFROM_SUMI_PERCENT).SetFocus
            Err_Chk = False
            Exit Function
        End If
    End If
    
    '2011.07.04
    If Trim(Text(ptxTO_SUMI_PERCENT).Text) = "" Then
        Text(ptxTO_SUMI_PERCENT).Text = "999"
    Else
        If Not IsNumeric(Text(ptxTO_SUMI_PERCENT).Text) Then
            MsgBox "���͂������ڂ̓G���[�ł��B�i���O���i���󋵁i���j�j"
            Text(ptxTO_SUMI_PERCENT).SetFocus
            Err_Chk = False
            Exit Function
        End If
    End If
    
    If Val(Text(ptxFROM_SUMI_PERCENT).Text) > Val(Text(ptxTO_SUMI_PERCENT).Text) Then
        MsgBox "���͂������ڂ̓G���[�ł��B�i���O���i���󋵁i���j�j"
        Text(ptxFROM_SUMI_PERCENT).SetFocus
        Err_Chk = False
        Exit Function
    End If
    
    '2011.07.04
    
    
    
    Err_Chk = False

End Function
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1060201.MousePointer = vbHourglass

    Call Ctrl_Lock(F1060201)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1060201)


    F1060201.MousePointer = vbDefault

End Sub


Private Sub Command_Click(Index As Integer)

Dim yn      As Integer
Dim i       As Integer
    
Dim mesg    As String
    
Dim c       As String * 128     '2008.03.03
    
    Select Case Index
        
        Case 7                              '�f�[�^�o��
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            
                                        '�d�������荞��   2008.03.03
            If GetIni(App.EXEName, Last_JGYOBU, App.EXEName, c) Then
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
            
            Text(ptxSOKO).SetFocus
        
        
        Case 8                              '���
            
            If Err_Chk() Then
                Exit Sub
            End If
            
                                        '�d�������荞��   2008.03.03
            If GetIni(App.EXEName, Last_JGYOBU, App.EXEName, c) Then
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
            Text(ptxSOKO).SetFocus
                    
        Case 11                             '�I��
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub

Private Sub Command1_Click()

    MsgBox "SORT=" & SORT_SEQ & Chr(13) & Chr(10) & _
            "OUTSIDE=" & OutSide & Chr(13) & Chr(10) & _
            "Sample_QTY=" & SAMPLE_QTY & Chr(13) & Chr(10) & _
            "NOT_Hin_Name=" & wkNOT_Hin_Name & Chr(13) & Chr(10)


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
Dim sts As Integer              '2011.03.31
Dim com As Integer              '2011.03.31
     
     
     If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
    
    
    
    
    '�X�e�[�^�X�E�B���h�E���쐬����     2011.01.12
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "���i���v��x���A���[�����X�g", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
    
    
    
    
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

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1060201.Caption = "���i���v��x���A���[�����X�g����i" + RTrim(JGYOBU_T(i).NAME) + "�j" & LAST_UPDATE_DAY

            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)           '2011.01.12
                                
                                
                                '���i���x���t�@�C������荞��
    If GetIni("FILE", "GOODS_DATA", "SYS", c) Then
        Beep
        MsgBox "'���i���x���t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    GOODS_DATA = Trim(c)
                                
                                
'------------------------------ SYS.INI--> F106020.INI 2011.07.04
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
                                'SORT����荞�� 2008.11.06
    If GetIni(App.EXEName, "SORT", App.EXEName, c) Then
        SORT_SEQ = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            SORT_SEQ = 0
        Else
            SORT_SEQ = CInt(Trim(c))
        End If
    End If
                                
                                '���{�� 2011.07.04
    If GetIni(App.EXEName, "Sample_QTY", App.EXEName, c) Then
        SAMPLE_QTY = 0
    Else
        If IsNumeric(Trim(c)) Then
            SAMPLE_QTY = CLng(Trim(c))
        Else
            SAMPLE_QTY = 0
        End If
    End If
                                
                                '�i���ɂ�鏜�O 2011.07.04
    NOT_Hin_Name_F = False
    If GetIni(App.EXEName, "NOT_HIN_NAME", App.EXEName, c) Then
    Else
        wkNOT_Hin_Name = Trim(c)
        NOT_Hin_Name = Split(Trim(c), ",", -1)
        NOT_Hin_Name_F = True
    End If
                                
                                
                                
'------------------------------ SYS.INI--> F106020.INI 2011.07.04
                                
                                
'------------------------------------   2011.07.04  ���ϊ��Ԃ̊l��
    If GetIni(App.EXEName, "TUKI1", "F120050", c) Then
        TUKI1 = 3
    Else
        If Not IsNumeric(RTrim(c)) Then
            TUKI1 = 3
        Else
            TUKI1 = Val(RTrim(c))
        End If
    End If
    TUKI1_TITLE = "�����Ϗo�א�(" & Format(TUKI1, "#0") & "����)"


    If GetIni(App.EXEName, "TUKI2", "F120050", c) Then
        TUKI2 = 3
    Else
        If Not IsNumeric(RTrim(c)) Then
            TUKI2 = 3
        Else
            TUKI2 = Val(RTrim(c))
        End If
    End If
    S_TUKI1_TITLE = "���Y�v��p�����Ϗo�א�(" & Format(TUKI2, "#0") & "����)"


    If GetIni(App.EXEName, "TUKI3", "F120050", c) Then
        TUKI3 = 12
    Else
        If Not IsNumeric(RTrim(c)) Then
            TUKI3 = 12
        Else
            TUKI3 = Val(RTrim(c))
        End If
    End If
    S_TUKI2_TITLE = "���Y�v��p�����Ϗo�א�(" & Format(TUKI3, "#0") & "����)"







'------------------------------------   2011.07.01
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
'-----------------------------------------------    2011.03.31
                                '���ޑΏێ��   2011.03.31
    If GetIni("SEI0010", "SHIZAI", "SEI0010", c) Then
        
        c = "**"
        SHIZAI_T = Split(Trim(c), ",", -1)
        
    Else
        SHIZAI_T = Split(Trim(c), ",", -1)
    End If
                                
                                '�����Ώێ��
    If GetIni("SEI0010", "DOUKON", "SEI0010", c) Then
        c = "**"
        DOUKON_T = Split(Trim(c), ",", -1)
    Else
        DOUKON_T = Split(Trim(c), ",", -1)
    End If
                                '���H�Ώێ��
   If GetIni("SEI0010", "KAKOU", "SEI0010", c) Then
        c = "**"
        KAKOU_T = Split(Trim(c), ",", -1)
    Else
        KAKOU_T = Split(Trim(c), ",", -1)
    End If
                                
                                
                                
                                
                                '�Z���^�[�̎��� 2011.03.31
    If GetIni("SEI0010", "KUSATU", "SEI0010", c) Then
        KUSATU_F = False
    Else
        If Trim(c) = "1" Then
            KUSATU_F = True
        Else
            KUSATU_F = False
        End If
    End If
                                
                                
                                
                                '�����ދ敪�̊l��
    If GetIni("SEI0010", "KOSOU", "SEI0010", c) Then
        KOSOU_KBN = ""
    Else
        KOSOU_KBN = Trim(c)
    End If
                                '�O�����ދ敪�̊l��
    If GetIni("SEI0010", "GAISO", "SEI0010", c) Then
        GAISO_KBN = ""
    Else
        GAISO_KBN = Trim(c)
    End If
                                
'-----------------------------------------------    2011.03.31
                                
                                
                                
                                
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
                                '�����Ϗo�א��n�o�d�m
    If AVE_SYUKA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                
                                
                                
                                '���i���w�}�[�f�[�^�n�o�d�m
    If P_SSHIJI_O_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                '�󕥐�}�X�^�n�o�d�m   2011.07.04
    If P_UKEHARAI_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                
                                '���i���W�v�t�@�C���n�o�d�m
    If GOODS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                '�Ǘ��}�X�^�n�o�d�m
    If P_KANRI_Open(BtOpenNomal) Then
        Unload Me
    End If

    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC, Len(P_KANRIREC), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�Ǘ��}�X�^(KEY=01)")
        Unload Me
    End Select

    Call UniCode_Conv(K0_P_KANRI.REC_NO, P_ST_KANRI_DEF_No)
    sts = BTRV(BtOpGetEqual, P_KANRI_POS, P_KANRIREC02, Len(P_KANRIREC02), K0_P_KANRI, Len(K0_P_KANRI), 0)
    Select Case sts
        Case BtNoErr
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�Ǘ��}�X�^(KEY=02)")
        Unload Me
    End Select
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                '����t�H���g�ݒ�(�ʏ�)
    With NormalFont
        .NAME = F1060201.FontName
        .Size = 12
    End With

                                '����t�H���g�ݒ�i���j
    With MidFont
        .NAME = F1060201.FontName
        .Size = 8
    End With


    Combo(pcmbNaigai).Clear
    Combo(pcmbNaigai).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo(pcmbNaigai).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo(pcmbNaigai).ListIndex = 0

    Show
    
    Text(ptxSOKO).SetFocus
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
                                            '�����Ϗo�א��b�k�n�r�d
    sts = BTRV(BtOpClose, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�����Ϗo�א�")
        End If
    End If
                                            '���i���W�v�t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���i���W�v�t�@�C��")
        End If
    End If
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1060201 = Nothing

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
    F1060201.Caption = "���i���v��x���A���[�����X�g����i" + RTrim(JGYOBU_T(Index).NAME) + "�j" & LAST_UPDATE_DAY
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
Dim i   As Integer
Dim sts As Integer

    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case ptxSOKO
            
            Text(Index).Text = StrConv(Trim(Text(Index).Text), vbUpperCase)
            
            If Len(Trim(Text(ptxSOKO).Text)) = 0 Then
                Label2.Caption = "�S�q��"
            Else
                Call UniCode_Conv(K0_SOKO.Soko_No, Text(ptxSOKO).Text)
            
                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                Select Case sts
                    Case BtNoErr
                        Label2.Caption = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
                        
                        If StrConv(SOKOREC.SOKO_BUN, vbUnicode) = BUN_KASO Then
                            Beep
                            MsgBox "���͂������ڂ̓G���[�ł��B�i���z�q�Ɂj"
                            Text(ptxSOKO).SetFocus
                            Exit Sub
                        End If
                    
                    Case BtErrKeyNotFound
                        Beep
                        MsgBox "���͂������ڂ̓G���[�ł��B�i�q�ɖ��o�^�j"
                        Text(ptxSOKO).SetFocus
                        Exit Sub
                    Case Else
                        Call File_Error(sts, BtOpGetGreaterEqual, "�q�Ƀ}�X�^")
                        Exit Sub
                End Select
            
            
            End If
    
''''''''''''''''''''''''''''''''''''''''    2011.07.04
'        Case ptxSUMI_PERCENT    '2008.03.03
'
'            If Trim(Text(ptxSUMI_PERCENT).Text) = "" Then
'            Else
'                If Not IsNumeric(Text(ptxSUMI_PERCENT).Text) Then
'                    Beep
'                    MsgBox "���͂������ڂ̓G���[�ł��B�i���O���i���󋵁i���j�j"
'                    Text(ptxSUMI_PERCENT).SetFocus
'                    Exit Sub
'                End If
'            End If
    
    
        Case ptxFROM_SUMI_PERCENT

            If Trim(Text(ptxFROM_SUMI_PERCENT).Text) = "" Then
                Text(ptxFROM_SUMI_PERCENT).Text = "000"
            Else
                If Not IsNumeric(Text(ptxFROM_SUMI_PERCENT).Text) Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B�i���O���i���󋵁i���j�j"
                    Text(ptxFROM_SUMI_PERCENT).SetFocus
                    Exit Sub
                End If
            End If
    
        Case ptxTO_SUMI_PERCENT

            If Trim(Text(ptxTO_SUMI_PERCENT).Text) = "" Then
                Text(ptxTO_SUMI_PERCENT).Text = "000"
            Else
                If Not IsNumeric(Text(ptxTO_SUMI_PERCENT).Text) Then
                    Beep
                    MsgBox "���͂������ڂ̓G���[�ł��B�i���O���i���󋵁i���j�j"
                    Text(ptxTO_SUMI_PERCENT).SetFocus
                    Exit Sub
                End If
            End If
    
                
            If Val(Text(ptxFROM_SUMI_PERCENT).Text) > Val(Text(ptxTO_SUMI_PERCENT).Text) Then
                Beep
                MsgBox "���͂������ڂ̓G���[�ł��B�i���O���i���󋵁i���j�j"
                Text(ptxFROM_SUMI_PERCENT).SetFocus
                Exit Sub
            End If
    
''''''''''''''''''''''''''''''''''''''''    2011.07.04
    
    End Select
        
    For i = Index + 1 To Text_Max
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
Dim Lcnt                As Integer

Dim sts                 As Integer
Dim com                 As Integer

Dim Save_Soko           As String * 2

Dim Edit                As String

Dim X_Tab               As Integer

'Dim wkSUMI_PERCENT      As Long
Dim wkFROM_SUMI_PERCENT As Long
Dim wkTO_SUMI_PERCENT   As Long


Dim SKIP_F              As Boolean
    
    Print_Proc = True

    Call Input_Lock



    If Data_Make_Proc() Then        '���i���x���W�v�f�[�^�쐬
        Exit Function
    End If


''2011.07.04    If Trim(Text(ptxSUMI_PERCENT).Text) = "" Then
''2011.07.04        wkSUMI_PERCENT = 100
''2011.07.04    Else
''2011.07.04        wkSUMI_PERCENT = CLng(Text(ptxSUMI_PERCENT).Text)
''2011.07.04    End If

    If Trim(Text(ptxFROM_SUMI_PERCENT).Text) = "" Then
        wkFROM_SUMI_PERCENT = 0
    Else
        wkFROM_SUMI_PERCENT = CLng(Text(ptxFROM_SUMI_PERCENT).Text)
    End If


    If Trim(Text(ptxTO_SUMI_PERCENT).Text) = "" Then
        wkTO_SUMI_PERCENT = 999
    Else
        wkTO_SUMI_PERCENT = CLng(Text(ptxTO_SUMI_PERCENT).Text)
    End If

    
        
    
    Lcnt = LCTL

    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time
    
    
    If SORT_SEQ = 0 Then    '2008.11.06

        Call UniCode_Conv(K0_GOODS.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_GOODS.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        Call UniCode_Conv(K0_GOODS.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "")
        Call UniCode_Conv(K0_GOODS.HIN_GAI, "")
    Else
    
    
    
        Call UniCode_Conv(K3_GOODS.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K3_GOODS.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        Call UniCode_Conv(K3_GOODS.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K3_GOODS.AVE_SYUKA, "zzzzzzzz")
        Call UniCode_Conv(K3_GOODS.Sumi_QTY, "")
        Call UniCode_Conv(K3_GOODS.Mi_QTY, "zzzzzzzz")
        Call UniCode_Conv(K3_GOODS.SUMI_PERCENT, "")
        Call UniCode_Conv(K3_GOODS.HIN_GAI, "")
    End If
    
    
    com = BtOpGetGreater
    
    Do
        
        If SORT_SEQ = 0 Then    '2008.11.06
        
            sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
        Else
            sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K3_GOODS, Len(K3_GOODS), 3)
        End If
        
        Select Case sts
            Case BtNoErr
                
                
                
                
                
                If StrConv(GOODSREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(GOODSREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                    Exit Do
                End If
                                        
                
                If Len(Trim(Text(ptxSOKO).Text)) = 0 Then
                Else
                    If StrConv(GOODSREC.ST_SOKO, vbUnicode) <> Text(ptxSOKO).Text Then
                        Exit Do
                    End If
                End If
            
                SKIP_F = False
                If Not IsNumeric(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) Then
                    SKIP_F = True
                Else
''2011.07.04                    If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > wkSUMI_PERCENT Then
''2011.07.04                        SKIP_F = True
''2011.07.04                    End If
                    If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) < wkFROM_SUMI_PERCENT Or _
                        CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > wkTO_SUMI_PERCENT Then
                        SKIP_F = True
                    End If
                
                
                End If
                

                If CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) <= 0 Then
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
                Save_Soko = StrConv(GOODSREC.ST_SOKO, vbUnicode)
                
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
                
            End If
            
            If Save_Soko <> StrConv(GOODSREC.ST_SOKO, vbUnicode) Then
                                
                Lcnt = LMAX + 1
                Save_Soko = StrConv(GOODSREC.ST_SOKO, vbUnicode)
                
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
                
            End If
            
            
            If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                            '�ݒ蔭���_���傫��
                Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "99999999")
                Call UniCode_Conv(K0_GOODS.HIN_GAI, "zzzzzzzzzzzzzzzzzzzz")
                com = BtOpGetGreater
            Else
                '�����i�݌Ɂ��O �́A����ΏۊO 2004.08.27
                
                
                
                
                If OutSide >= CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) Or _
                    CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) <= 0 Then
                Else
                    If Head_Print_Proc(Lcnt) Then
                        Exit Function
                    End If
                
                    X_Tab = MGN_L
                
                    Printer.Print Tab(X_Tab);
                                                            '�W���I��
                    Edit = StrConv(GOODSREC.ST_RETU, vbUnicode) & "-"
                    Edit = Edit & StrConv(GOODSREC.ST_REN, vbUnicode) & "-"
                    Edit = Edit & StrConv(GOODSREC.ST_DAN, vbUnicode)
                    Printer.Print Edit;
                    X_Tab = X_Tab + Len(Edit) + 5
    '                X_Tab = X_Tab + Len(Edit) + 3
                                                            '�i�ԁi�O���j
                    Printer.Print Tab(X_Tab);
    
                    Printer.Print StrConv(GOODSREC.HIN_GAI, vbUnicode);
    '                X_Tab = X_Tab + Len(StrConv(GOODSREC.HIN_GAI, vbUnicode)) + 5
                    X_Tab = X_Tab + Len(StrConv(GOODSREC.HIN_GAI, vbUnicode)) + 4
                                                            '����
                    Printer.Print Tab(X_Tab);
'2008.11.06                    Printer.Print StrConv(GOODSREC.PACKING_NO, vbUnicode);
    '                X_Tab = X_Tab + Len(StrConv(GOODSREC.PACKING_NO, vbUnicode)) + 5
'2008.11.06                    X_Tab = X_Tab + Len(StrConv(GOODSREC.PACKING_NO, vbUnicode)) + 4
                                                            
                                                            
                                                            
                                                            
                    '2008.11.06
                    Printer.Print Left(StrConv(GOODSREC.KOSOU, vbUnicode), 4);
                    X_Tab = X_Tab + Len(StrConv(GOODSREC.PACKING_NO, vbUnicode)) + 4
                    '2008.11.06
                                                            
                                                            
                                                            
                                                            
                                                            
                                                            
                                                            
                                                            
                                                            '���i���ςݍ݌ɐ�
                    Printer.Print Tab(X_Tab);
                    Edit = Format(CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Printer.Print Edit;
    '                X_Tab = X_Tab + Len(Edit) + 8
                    X_Tab = X_Tab + Len(Edit) + 2
                                                            '�����i�݌ɐ�
                    Printer.Print Tab(X_Tab);
                    Edit = Format(CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)), "#,##0")
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Printer.Print Edit;
    '                X_Tab = X_Tab + Len(Edit) + 8
                    X_Tab = X_Tab + Len(Edit) + 2
                                                            '�����Ϗo�א�
                    Printer.Print Tab(X_Tab);
                    Edit = Format(CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)), "#,##0")
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Printer.Print Edit;
    '                X_Tab = X_Tab + Len(Edit) + 8
                    X_Tab = X_Tab + Len(Edit) + 2
                                                            '���O���i���K�v��
                    Printer.Print Tab(X_Tab);
                    Edit = Format(CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) - CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Printer.Print Edit;
    '                X_Tab = X_Tab + Len(Edit) + 8
                    X_Tab = X_Tab + Len(Edit) + 2
                                                            '���O���i����
                    Printer.Print Tab(X_Tab);
                    Edit = Format(CInt(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)), "#0") & "%"
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
    
                    Printer.Print Edit;
                    X_Tab = X_Tab + Len(Edit) + 5
                                                            '�ʒu�݌�
                    Printer.Print Tab(X_Tab);
    
                    If MI_ZAIKO_KENSAKU(StrConv(GOODSREC.HIN_GAI, vbUnicode)) Then
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
                    End If
    
                    Printer.Print Edit
    
                    Printer.Print
                
                    Lcnt = Lcnt + 2
            
                End If
            End If
            
            com = BtOpGetNext
        
        End If
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

    Printer.Print Tab(MGN_L + 55);
    
    Printer.Print "���i���x���A���[�����X�g";
    
    
    Printer.Print Tab(MGN_L + 90);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print

    
    Printer.Print Tab(MGN_L);
    Printer.Print "�q�ɁF";
    Printer.Print StrConv(SOKOREC.Soko_No, vbUnicode) & "  ";
    Printer.Print StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "  ";
    Printer.Print "�i�ݒ蔭���_ " & StrConv(Format(CLng(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0"), vbWide) & "���j"
    Printer.Print

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
Dim AVE_QTY     As Double


'2011.01.13
Dim In_Cnt      As Long
Dim Out_Cnt     As Long
'2011.01.13



'2011.07.04
Dim SKIP_FLG    As Integer
Dim i           As Integer
'2011.07.04

    Data_Make_Proc = True

In_Cnt = 0
Out_Cnt = 0

'---------------------------------------------------------- '�S���R�[�h�폜
    com = BtOpGetFirst
    Do
        
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
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
            
            sts = BTRV(BtOpDelete, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
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
'---------------------------------------------------------- '�i�ڃ}�X�^�x�[�X�Ńf�[�^�쐬

    Call UniCode_Conv(K6_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K6_ITEM.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
    Call UniCode_Conv(K6_ITEM.ST_SOKO, Trim(Text(ptxSOKO).Text))
    Call UniCode_Conv(K6_ITEM.ST_RETU, "")
    Call UniCode_Conv(K6_ITEM.ST_REN, "")
    Call UniCode_Conv(K6_ITEM.ST_DAN, "")
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
            
                If Len(Trim(Text(ptxSOKO).Text)) = 0 Then
                Else
                    If StrConv(ITEMREC.ST_SOKO, vbUnicode) <> Text(ptxSOKO).Text Then
                        '�q�ɔԍ��u���[�N
                        Exit Do
                    End If
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                Exit Function
        End Select
        '-----------------------------------------  '���i���W�v�t�@�C���쐬
        
In_Cnt = In_Cnt + 1
        
        '���o�����ύX       2011.07.04
        SKIP_FLG = False
        If StrConv(ITEMREC.GOODS_KBN, vbUnicode) <> GOODS_ON Then
            SKIP_FLG = True
        End If
                
        If StrConv(ITEMREC.GOODS_OUT_F, vbUnicode) = "1" Then
            SKIP_FLG = True
        End If
                
'2011.07.25
'        If StrConv(ITEMREC.NAI_BUHIN, vbUnicode) <> "1" And StrConv(ITEMREC.NAI_BUHIN, vbUnicode) <> "2" Then
        If StrConv(ITEMREC.NAI_BUHIN, vbUnicode) <> "1" And StrConv(ITEMREC.NAI_BUHIN, vbUnicode) <> "2" And Trim(StrConv(ITEMREC.NAI_BUHIN, vbUnicode)) <> "" Then
'2011.07.25
            
            
            SKIP_FLG = True
        End If
        
        If NOT_Hin_Name_F Then
            For i = 0 To UBound(NOT_Hin_Name)
                If InStr(1, RTrim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), NOT_Hin_Name(i)) <> 0 Then
                    SKIP_FLG = True
                    Exit For
                End If
            Next i
        End If
        '���o�����ύX       2011.07.04
                
        
        
'2011.07.04        If StrConv(ITEMREC.GOODS_KBN, vbUnicode) = GOODS_ON Then
        If Not SKIP_FLG Then
                    
            
Out_Cnt = Out_Cnt + 1
            
            
                                                    '���ƕ�
            Call UniCode_Conv(GOODSREC.JGYOBU, Last_JGYOBU)
                                                    '�����O
            Call UniCode_Conv(GOODSREC.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
                                                    '�i�ԁi�O���j
            Call UniCode_Conv(GOODSREC.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                                                    '�W���I��
            Call UniCode_Conv(GOODSREC.ST_SOKO, StrConv(ITEMREC.ST_SOKO, vbUnicode))
            Call UniCode_Conv(GOODSREC.ST_RETU, StrConv(ITEMREC.ST_RETU, vbUnicode))
            Call UniCode_Conv(GOODSREC.ST_REN, StrConv(ITEMREC.ST_REN, vbUnicode))
            Call UniCode_Conv(GOODSREC.ST_DAN, StrConv(ITEMREC.ST_DAN, vbUnicode))
                                                    '����
            Call UniCode_Conv(GOODSREC.PACKING_NO, StrConv(ITEMREC.PACKING_NO, vbUnicode))
            
                                                    '�݌ɏW�v����
            If Zaiko_Syukei_Proc(Sumi_QTY, _
                                    Mi_QTY, _
                                    Last_JGYOBU, _
                                    Right(Combo(pcmbNaigai).Text, 1), _
                                    StrConv(ITEMREC.HIN_GAI, vbUnicode)) = SYS_ERR Then
                Exit Function
            End If
                                                    '���i���ςݍ݌ɐ�
            
            
            '2011.07.04
            Sumi_QTY = Sumi_QTY - SAMPLE_QTY
            If Sumi_QTY < 0 Then
                Sumi_QTY = 0
            End If
            '2011.07.04
            
            Call UniCode_Conv(GOODSREC.Sumi_QTY, Format(Sumi_QTY, "00000000"))
                                                    '�����i�݌ɐ�
            Call UniCode_Conv(GOODSREC.Mi_QTY, Format(Mi_QTY, "00000000"))
                                                    '�����Ϗo�א�
            Call UniCode_Conv(K0_AVE_SYUKA.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K0_AVE_SYUKA.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
            Call UniCode_Conv(K0_AVE_SYUKA.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
            
            AVE_QTY = 0
            sts = BTRV(BtOpGetEqual, AVE_SYUKA_POS, AVE_SYUKAREC, Len(AVE_SYUKAREC), K0_AVE_SYUKA, Len(K0_AVE_SYUKA), 0)
            Select Case sts
                Case BtNoErr
''2011.07.04                    Call UniCode_Conv(GOODSREC.AVE_SYUKA, StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
''2011.07.04                    AVE_QTY = CDbl(StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
                    Call UniCode_Conv(GOODSREC.AVE_SYUKA, StrConv(AVE_SYUKAREC.AVE_SYUKA, vbUnicode))
                    Call UniCode_Conv(GOODSREC.S_AVE_SYUKA_QTY1, StrConv(AVE_SYUKAREC.S_AVE_SYUKA_QTY1, vbUnicode))
                    Call UniCode_Conv(GOODSREC.S_AVE_SYUKA_QTY2, StrConv(AVE_SYUKAREC.S_AVE_SYUKA_QTY2, vbUnicode))
                    AVE_QTY = CDbl(StrConv(AVE_SYUKAREC.S_AVE_SYUKA_QTY1, vbUnicode))
                Case BtErrKeyNotFound
                    Call UniCode_Conv(GOODSREC.AVE_SYUKA, "00000000")
                    Call UniCode_Conv(GOODSREC.S_AVE_SYUKA_QTY1, "00000000")
                    Call UniCode_Conv(GOODSREC.S_AVE_SYUKA_QTY2, "00000000")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�����Ϗo�א�")
                    Exit Function
            End Select
                                                    '���O���i����
            If AVE_QTY = 0 Then
                Call UniCode_Conv(GOODSREC.SUMI_PERCENT, "00000000")
            Else
                Call UniCode_Conv(GOODSREC.SUMI_PERCENT, Format(CLng(Sumi_QTY / AVE_QTY * 100), "00000000"))
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
                    Call UniCode_Conv(GOODSREC.KOSOU, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                Case BtErrKeyNotFound
                    Call UniCode_Conv(GOODSREC.KOSOU, "")
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
                    Call UniCode_Conv(GOODSREC.GAISOU, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    Call UniCode_Conv(GOODSREC.KO_QTY, StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                Case BtErrKeyNotFound
                    Call UniCode_Conv(GOODSREC.GAISOU, "")
                    Call UniCode_Conv(GOODSREC.KO_QTY, "000.00")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                    Exit Function
            End Select
            
            
            '2011.07.04
            
            Call UniCode_Conv(GOODSREC.NAI_BUHIN, StrConv(ITEMREC.NAI_BUHIN, vbUnicode))
            Call UniCode_Conv(GOODSREC.GAI_BUHIN, StrConv(ITEMREC.GAI_BUHIN, vbUnicode))
            '2011.07.04
            
            
            
            Do
                
                sts = BTRV(BtOpInsert, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
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
                        Call File_Error(sts, BtOpInsert, "���i���x���W�v�f�[�^")
                        Exit Function
                End Select
            
            Loop
        End If
        
If Right(Format(In_Cnt, "000"), 2) = "00" Then
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "�W�v�f�[�^�o�͒��I�I[" & Out_Cnt & "/" & In_Cnt & "]", Me.hwnd, 0)
End If
        
        
        
        com = BtOpGetNext
    Loop

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "�W�v�f�[�^�o�͒��I�I[" & Out_Cnt & "/" & In_Cnt & "]", Me.hwnd, 0)

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
    
''2011.07.04Dim wkSUMI_PERCENT      As Long



Dim SKIP_F              As Boolean
Dim FSW             As Boolean
    
    
'2011.01.13
Dim In_Cnt          As Long
Dim Out_Cnt         As Long

Dim JISEKI_TOTAL    As Double

Dim i               As Integer


Dim Shiji_No        As String * 8
'2011.01.13
    
    
'2011.03.31
Dim MAIN_KOUTEI(0 To 9) _
                    As Long
Dim wkTANI          As Double
Dim wkQTY           As Double

Dim KOUSEI()        As KOUSEI_TBL
Dim j               As Integer
Dim KOUSEI_FLG      As Boolean

Dim wkInt           As Integer
'2011.03.31
    
    Data_Proc = True

In_Cnt = 0
Out_Cnt = 0


'2011.07.04
Dim Line_Cnt            As Long
Dim wkFROM_SUMI_PERCENT As Long
Dim wkTO_SUMI_PERCENT   As Long

Dim wkUKEHARAI_CODE     As String * 5

Dim wkHIN_NAME          As String * 40
'2011.07.04

    Call Input_Lock





    fileName = GOODS_DATA
    sts = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), sts) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - sts)
    
    On Error GoTo Error_Proc
    
    FileNo = FreeFile
    Open (fileName) For Output As FileNo
    On Error GoTo 0


    If Data_Make_Proc() Then        '���i���x���W�v�f�[�^�쐬
        Exit Function
    End If
    
    
''2011.07.04    If Trim(Text(ptxSUMI_PERCENT).Text) = "" Then
''2011.07.04        wkSUMI_PERCENT = 100
''2011.07.04    Else
''2011.07.04        wkSUMI_PERCENT = CLng(Text(ptxSUMI_PERCENT).Text)
''2011.07.04    End If
    
    
    '2011.07.04
    If Trim(Text(ptxFROM_SUMI_PERCENT).Text) = "" Then
        wkFROM_SUMI_PERCENT = 0
    Else
        wkFROM_SUMI_PERCENT = CDbl(Text(ptxFROM_SUMI_PERCENT).Text)
    End If
    
    If Trim(Text(ptxTO_SUMI_PERCENT).Text) = "" Then
        wkTO_SUMI_PERCENT = 999
    Else
        wkTO_SUMI_PERCENT = CDbl(Text(ptxTO_SUMI_PERCENT).Text)
    End If
    '2011.07.04
    
    
    
    
    
    
    FSW = True
    
    

    If SORT_SEQ = 0 Then        '2008.11.06


        Call UniCode_Conv(K0_GOODS.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_GOODS.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        Call UniCode_Conv(K0_GOODS.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "")
        Call UniCode_Conv(K0_GOODS.HIN_GAI, "")
    Else
    
        Call UniCode_Conv(K3_GOODS.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K3_GOODS.NAIGAI, Right(Combo(pcmbNaigai).Text, 1))
        Call UniCode_Conv(K3_GOODS.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K3_GOODS.AVE_SYUKA, "zzzzzzzz")
        Call UniCode_Conv(K3_GOODS.Sumi_QTY, "")
        Call UniCode_Conv(K3_GOODS.Mi_QTY, "zzzzzzzz")
        Call UniCode_Conv(K3_GOODS.SUMI_PERCENT, "")
        Call UniCode_Conv(K3_GOODS.HIN_GAI, "")
    
    End If
    
    com = BtOpGetGreater
    
    
        
    '2011.07.04
    Line_Cnt = 0
    '2011.07.12
    In_Cnt = 0
    Out_Cnt = 0
    '2011.07.12
    
    Do
        If SORT_SEQ = 0 Then        '2008.11.06
            sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K0_GOODS, Len(K0_GOODS), 0)
        Else
            sts = BTRV(com, GOODS_POS, GOODSREC, Len(GOODSREC), K3_GOODS, Len(K3_GOODS), 3)
        End If
        Select Case sts
            Case BtNoErr
                If StrConv(GOODSREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(GOODSREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNaigai).Text, 1) Then
                    Exit Do
                End If
                                        
                
                If Len(Trim(Text(ptxSOKO).Text)) = 0 Then
                Else
                    If StrConv(GOODSREC.ST_SOKO, vbUnicode) <> Text(ptxSOKO).Text Then
                        Exit Do
                    End If
                End If
                
                SKIP_F = False
                If Not IsNumeric(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) Then
                    SKIP_F = True
                Else
''2011.07.04                    If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > wkSUMI_PERCENT Then
''2011.07.04                        SKIP_F = True
''2011.07.04                    End If
                
''2011.07.04
                    If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) < wkFROM_SUMI_PERCENT Then
                        SKIP_F = True
                    End If
                    
                    If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > wkTO_SUMI_PERCENT Then
                        SKIP_F = True
                    End If
''2011.07.04
                End If

                If CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)) <= 0 Then
                    SKIP_F = True
                End If

            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���i���W�v�t�@�C��")
                Exit Function
        End Select

        
'-------------------------------------------------  ���׈��
        
        In_Cnt = In_Cnt + 1
        
        
        If Not SKIP_F Then
        
        


        
        
        
            If FSW Then
                
                FSW = False
                        
                Save_Soko = StrConv(GOODSREC.ST_SOKO, vbUnicode)
                
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
''2011.07.12                Line_Cnt = Line_Cnt + 1
''2011.07.12                Write #FileNo, "*** ���i���x���A���[�����X�g�@***"
''2011.07.04                Write #FileNo, "�쐬���t:" & Format(Now, "YYYY/MM/DD") & " " & Format(Now, "HH:MM:SS")
                        
            
'                Write #FileNo, "�W���I��", "�i�ԁi�O���j", "���ށi�����j", "���i���ύ݌�", "�����i�݌�", "�����i�@�ʒu��1", "�����i�@�ʒu��2", "�����i�@�ʒu��3", "�����Ϗo�א�", "���O���i���K�v��", "���O���i����"
''2011.01.13                Write #FileNo, "�W���I��", "�i�ԁi�O���j", "���ށi�����j", "���i���ύ݌�", "�����i�݌�", "�����i�@�ʒu��1", "�����i�@�ʒu��2", "�����i�@�ʒu��3", "�����Ϗo�א�", "���O���i���K�v��", "���O���i����", "�H���@�i���^�j"
                
''2011.03.31                Write #FileNo, "�W���I��", "�i�ԁi�O���j", "���ށi�����j", "���i���ύ݌�", "�����i�݌�", "�����i�@�ʒu��1", "�����i�@�ʒu��2", "�����i�@�ʒu��3", "�����Ϗo�א�", "���O���i���K�v��", "���O���i����", "���ύH���@�i���^�j", "���эH���@�i���^�j"
                
                
''2011.07.04                Write #FileNo, "�W���I��", "�i�ԁi�O���j", "���ށi�����j", "���i���ύ݌�", "�����i�݌�", "�����i�@�ʒu��1", "�����i�@�ʒu��2", "�����i�@�ʒu��3", "�����Ϗo�א�", "���O���i���K�v��", "���O���i����", "���ύH���@�i���^�j", "���эH���@�i���^�j", "��ƍH��"
                
                
                '2011.07.04
                Line_Cnt = Line_Cnt + 1
''2011.07.12                Write #FileNo, "�W���I��", "�i�ԁi�O���j", "���ށi�����j", "���i���\�萔", "�O���i��", "�O���g�p����", "���i���ύ݌�", "�����i�݌�", "�����i�@�ʒu1", "�����i�@�ʒu2", TUKI1_TITLE, S_TUKI1_TITLE, S_TUKI2_TITLE, "���O���i���K�v��", "���O���i����", "���ύH���@�i���^�j", "���эH���@�i���^�j", "��ƍH��", "���������敪", "�C�O�����敪", "���i��������z��"
                '2011.07.12
                Write #FileNo, "�W���I��", "�i�ԁi�O���j", "���ށi�����j", "���i���\�萔", "���i���\��H��", "�O���i��", "�O���g�p����", "���i���ύ݌�", "�����i�݌�", "�����i�@�ʒu1", "�����i�@�ʒu�݌�1", "�����i�@�ʒu2", "�����i�@�ʒu�݌�2", TUKI1_TITLE, S_TUKI1_TITLE, S_TUKI2_TITLE, "���O���i���K�v��", "���O���i����", "���ύH���@�i���^�j", "���эH���@�i���^�j", "��ƍH��", "���������敪", "�C�O�����敪", "���i��������z��"
                
                
            
''2011.07.04                Write #FileNo, "�q�ɇ��F" & StrConv(SOKOREC.Soko_No, vbUnicode) & " " & StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "(�����_" & Format(CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0") & "%)"
                
            
            
            End If
            
            If Save_Soko <> StrConv(GOODSREC.ST_SOKO, vbUnicode) Then
                                
                Save_Soko = StrConv(GOODSREC.ST_SOKO, vbUnicode)
                
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
                
''2011.07.04                Write #FileNo, "�q�ɇ��F" & StrConv(SOKOREC.Soko_No, vbUnicode) & " " & StrConv(SOKOREC.SOKO_NAME, vbUnicode) & "(�����_" & Format(CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)), "#0") & "%)"
                
                
            End If
            
            
            If CLng(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)) > CInt(StrConv(SOKOREC.ORDER_POINT, vbUnicode)) Then
                            '�ݒ蔭���_���傫��
                Call UniCode_Conv(K0_GOODS.SUMI_PERCENT, "99999999")
                Call UniCode_Conv(K0_GOODS.HIN_GAI, "zzzzzzzzzzzzz")
                com = BtOpGetGreaterEqual
            Else
                
                If OutSide >= CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) Then
                Else
                    Line_Cnt = Line_Cnt + 1
                
                
                                                            '�W���I��
                                    
                    Edit = StrConv(SOKOREC.Soko_No, vbUnicode) & "-"
                    Edit = Edit & StrConv(GOODSREC.ST_RETU, vbUnicode) & "-"
                    Edit = Edit & StrConv(GOODSREC.ST_REN, vbUnicode) & "-"
                    Edit = Edit & StrConv(GOODSREC.ST_DAN, vbUnicode)
                    Write #FileNo, Edit,
                                                            '�i�ԁi�O���j
    
                    Write #FileNo, Trim(StrConv(GOODSREC.HIN_GAI, vbUnicode)),
                                                            '����
                    'Write #FileNo, Trim(StrConv(GOODSREC.PACKING_NO, vbUnicode)),      '2008.03.03
                    Write #FileNo, Trim(StrConv(GOODSREC.KOSOU, vbUnicode)),            '2008.03.03
                    Write #FileNo, "",                                                  '2011.07.04
                                                                    
                    '���i���\�萔(���͍���)     2011.07.12
                                                            
                                                            
                    '���i���\��H��             2011.07.12
                    Write #FileNo, "=round(D" & Format(Line_Cnt, "#") & "*S" & Format(Line_Cnt, "#") & ",1)",
                                                            
                                                            
                                                            
                    '�O���� 2011.07.04
                    Write #FileNo, Trim(StrConv(GOODSREC.GAISOU, vbUnicode)),
                    '�O�����g�p���� 2011.07.04
                    If Val(StrConv(GOODSREC.KO_QTY, vbUnicode)) = 0 Then
                        Write #FileNo, 0,
                    Else
                        Write #FileNo, "=roundup(D" & Format(Line_Cnt, "#") & "/" & CDbl(StrConv(GOODSREC.KO_QTY, vbUnicode)) & ",0)",
                    End If
                                                            '���i���ςݍ݌ɐ�
                    Edit = Format(CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Write #FileNo, Edit,
                                                            '�����i�݌ɐ�
                    Edit = Format(CLng(StrConv(GOODSREC.Mi_QTY, vbUnicode)), "#,##0")
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Write #FileNo, Edit,
                                                            
                    If MI_ZAIKO_KENSAKU(StrConv(GOODSREC.HIN_GAI, vbUnicode)) Then
                        Exit Function
                    End If
                                                            '�����i�ʒu��
                    If Len(Trim(EE_ZAIKO_TBL(0).EE_LOC)) = 0 Then
''2011.07.12                        Write #FileNo, ,
                        '2011.07.12
                        Write #FileNo, , ,
                    Else
                        Edit = Left(EE_ZAIKO_TBL(0).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(0).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(0).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(0).EE_LOC, 2)
                        
                        '2011.07.12
                        Write #FileNo, Edit,
                        Edit = ""
                        Edit = Format(EE_ZAIKO_TBL(0).EE_QTY, "#0")
                        '2011.07.12
                        Write #FileNo, Edit,
                    End If
                                                            
                    If Len(Trim(EE_ZAIKO_TBL(1).EE_LOC)) = 0 Then
''2011.07.12                        Write #FileNo, ,
                        Write #FileNo, , ,
                    Else
                        Edit = Left(EE_ZAIKO_TBL(1).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(1).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(1).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(1).EE_LOC, 2)
                        
                        
                        '2011.07.12
                        Write #FileNo, Edit,
                        Edit = ""
                        Edit = Format(EE_ZAIKO_TBL(1).EE_QTY, "#0")
                        '2011.07.12
                        
                        
''2011.07.12                        Edit = Edit & " " & Format(EE_ZAIKO_TBL(1).EE_QTY, "#0")
                        Write #FileNo, Edit,
                    End If
                                                            
''2011.07.04                    If Len(Trim(EE_ZAIKO_TBL(2).EE_LOC)) = 0 Then
''2011.07.04                        Write #FileNo, ,
''2011.07.04                    Else
''2011.07.04                        Edit = Left(EE_ZAIKO_TBL(2).EE_LOC, 2) & "-" & Mid(EE_ZAIKO_TBL(2).EE_LOC, 3, 2) & "-" & Mid(EE_ZAIKO_TBL(2).EE_LOC, 5, 2) & "-" & Right(EE_ZAIKO_TBL(2).EE_LOC, 2)
''2011.07.04                        Edit = Edit & " " & Format(EE_ZAIKO_TBL(2).EE_QTY, "#0")
''2011.07.04                        Write #FileNo, Edit,
''2011.07.04                    End If
                                                            
                                                            '�����Ϗo�א�
                    Edit = Format(CDbl(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)), "#,##0.0")
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Write #FileNo, Edit,
                                                            
                                                            '���Y�v�挎���Ϗo�א�(1)    2011.07.04
                    Edit = Format(CDbl(StrConv(GOODSREC.S_AVE_SYUKA_QTY1, vbUnicode)), "#,##0.0")
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Write #FileNo, Edit,
                                                            '���Y�v�挎���Ϗo�א�(2)    2011.07.04
                    Edit = Format(CDbl(StrConv(GOODSREC.S_AVE_SYUKA_QTY2, vbUnicode)), "#,##0.0")
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Write #FileNo, Edit,
                                                            
                                                            
                    '2011.07.04
                                                            '���O���i���K�v��
                    'Edit = Format(CLng(StrConv(GOODSREC.AVE_SYUKA, vbUnicode)) - CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
                    Edit = Format(CLng(StrConv(GOODSREC.S_AVE_SYUKA_QTY1, vbUnicode)) - CLng(StrConv(GOODSREC.Sumi_QTY, vbUnicode)), "#,##0")
                    '2011.07.04
                    
                    
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Write #FileNo, Edit,
                                                            '���O���i����
                    Edit = Format(CInt(StrConv(GOODSREC.SUMI_PERCENT, vbUnicode)), "#0") & "%"
                    If Len(Edit) < 10 Then
                        Edit = Space(10 - Len(Edit)) & Edit
                    End If
                    Write #FileNo, Edit,
                    
                    
                    
                    '2008.09.19
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(GOODSREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(GOODSREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(GOODSREC.HIN_GAI, vbUnicode))
                    
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                            If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
                            
                            
                            
                                Edit = Format(CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode)), "#0.0")
                            
                            Else
                                Edit = ""
                            
                            
                            
                            End If
                            wkHIN_NAME = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                            Write #FileNo, Edit,
                        
                        
                                                    
                        Case BtErrKeyNotFound
                
                            Edit = ""
                            
                            wkHIN_NAME = ""
                            
                            
                                                
                        
                            Write #FileNo, Edit,
                        
                            '2011.07.04
                            Call UniCode_Conv(ITEMREC.GAI_BUHIN, "")
                            '2011.07.04
                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
                    
                    
                    
                    
                    '2011.01.13 ���эH��
                    Call UniCode_Conv(K1_P_SSHIJI_O.KAN_F, P_KAN_ON)
                    Call UniCode_Conv(K1_P_SSHIJI_O.SHIMUKE_CODE, SHIMUKE_CODE)
                    Call UniCode_Conv(K1_P_SSHIJI_O.JGYOBU, StrConv(GOODSREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K1_P_SSHIJI_O.NAIGAI, StrConv(GOODSREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K1_P_SSHIJI_O.HIN_GAI, StrConv(GOODSREC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(K1_P_SSHIJI_O.KAN_DT, "zzzzzzzz")
                    Call UniCode_Conv(K1_P_SSHIJI_O.Shiji_No, "zzzzzzzz")
                    
                    JISEKI_TOTAL = 0
                    
                    Shiji_No = ""
                                    
If Trim(StrConv(GOODSREC.HIN_GAI, vbUnicode)) = "AZB03-728-0S" Then
Debug.Print
End If
                    Do
                    
                    
                    
                        DoEvents
                        wkUKEHARAI_CODE = ""
                        sts = BTRV(BtOpGetLess, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K1_P_SSHIJI_O, Len(K1_P_SSHIJI_O), 1)
                        Select Case sts
                            Case BtNoErr
                                If StrConv(P_SSHIJI_O_REC.KAN_F, vbUnicode) <> P_KAN_ON Or _
                                    StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) <> SHIMUKE_CODE Or _
                                    StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) <> StrConv(GOODSREC.JGYOBU, vbUnicode) Or _
                                    StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) <> StrConv(GOODSREC.NAIGAI, vbUnicode) Or _
                                    StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode) <> StrConv(GOODSREC.HIN_GAI, vbUnicode) Then
                                    Exit Do
                                Else
                                    
                                    
                                    
                                    If Val(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) <= 0 Then
                                    Else
                                    
                                        
                                        Shiji_No = StrConv(P_SSHIJI_O_REC.Shiji_No, vbUnicode)
                                                        
                                        JISEKI_TOTAL = 0
                            
                            Debug.Print StrConv(P_SSHIJI_O_REC.Shiji_No, vbUnicode)
                                        
                                        For i = 0 To 9
                                        
                                            If IsNumeric(StrConv(P_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)) And IsNumeric(StrConv(P_SSHIJI_O_REC.GENKA_TBL(i).TIMES, vbUnicode)) Then
                                                JISEKI_TOTAL = JISEKI_TOTAL + Round(Val(StrConv(P_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)) * Val(StrConv(P_SSHIJI_O_REC.GENKA_TBL(i).TIMES, vbUnicode)), 2)
                                            End If
                                        
                                        Next i
                                    
                                        wkUKEHARAI_CODE = StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode)
                                    
                                        Exit Do
                                    End If
                                                                    
                                End If
                        
                            Case BtErrEOF
                                Exit Do
                            Case Else
                        
                                Call File_Error(sts, BtOpGetLess, "�w�}�[�f�[�^(�e)")
                                Exit Function
                        End Select
                    
                    Loop
                    
                    Edit = ""
                    If IsNumeric(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) Then
                        If Val(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) <> 0 Then
                            Edit = Format(Round(CDbl(JISEKI_TOTAL / Val(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode))), 2), "#0.0")
                        End If
                    End If
                    
                    Write #FileNo, Edit,
                    
                    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    2011.03.31
                    For i = 0 To UBound(MAIN_KOUTEI)
                        MAIN_KOUTEI(i) = 0
                    Next i
                                        
                                        
                    '�@
                    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode)) Then
                        
                        wkTANI = Val(StrConv(P_KANRIREC02.MAIN_KOTEI(0).KOTEI, vbUnicode))
                    Else
                        wkTANI = 0
                    End If
                    If IsNumeric(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode)) Then
                        '2009.09.18
                        If IsDate(Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 1, 4) & "/" & Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 5, 2) & "/" & Mid(StrConv(ITEMREC.S_KOUSU_SET_DATE, vbUnicode), 7, 4)) Then
                            wkQTY = Val(StrConv(ITEMREC.SEI_LABEL_QTY, vbUnicode))
                        Else
                            wkQTY = 1
                        End If
                    Else
                        wkQTY = 1
                    End If
                    MAIN_KOUTEI(0) = wkTANI * wkQTY





                    '�A
                    '-------------------�@�\�����e�[�u���W�J
                        
                    Erase KOUSEI
                    i = -1
        
                    KOUSEI_FLG = False
                                    
                            
If Trim(StrConv(GOODSREC.HIN_GAI, vbUnicode)) = "CWE20C2985" Then
    Debug.Print
End If
                    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, SHIMUKE_CODE)
                    Call UniCode_Conv(K0_P_COMPO.JGYOBU, StrConv(GOODSREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_P_COMPO.NAIGAI, StrConv(GOODSREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, StrConv(GOODSREC.HIN_GAI, vbUnicode))
                       
                    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, P_HEAD)
                    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
                                    
                    com = BtOpGetGreater
                                    
                    Do
                        DoEvents
                
                        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        Select Case sts
                            Case BtNoErr
                
                                
                                If StrConv(P_COMPO_K_REC.SHIMUKE_CODE, vbUnicode) <> SHIMUKE_CODE Or _
                                    StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> StrConv(GOODSREC.JGYOBU, vbUnicode) Or _
                                    StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) <> StrConv(GOODSREC.NAIGAI, vbUnicode) Or _
                                    Trim(StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode)) <> Trim(StrConv(GOODSREC.HIN_GAI, vbUnicode)) Then
                                
                                    Exit Do
                            
                                End If
                
                            Case BtErrEOF
                                Exit Do
                            Case Else
                                Call File_Error(sts, BtOpGetNext, "�\���}�X�^")
                                Exit Function
                        End Select
                
                        If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_KOSOU Then
                            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, KOSOU_KBN)
                        End If
                        If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = P_GAISOU Then
                            Call UniCode_Conv(P_COMPO_K_REC.KO_SYUBETSU, GAISO_KBN)
                        End If
                
                        i = i + 1
                        KOUSEI_FLG = True
                            
                        ReDim Preserve KOUSEI(0 To i)
                        '���ƕ�
                        KOUSEI(i).KO_JGYOBU = StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode)
                        '�����O
                        KOUSEI(i).KO_NAIGAI = StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode)
                        
                        '���
                        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
                        Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode))
                        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
                            
                        Select Case sts
                            Case BtNoErr
                                KOUSEI(i).KO_SYUBETSU = Left(StrConv(P_CODEREC.C_Code, vbUnicode), 2)
                            Case BtErrKeyNotFound
                                KOUSEI(i).KO_SYUBETSU = ""
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�R�[�h�}�X�^")
                                Exit Function
                        
                        End Select
                        
                        '�i��
                        KOUSEI(i).KO_HIN_GAI = StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode)
                         
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_COMPO_K_REC.KO_JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_COMPO_K_REC.KO_NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            
                            Case BtErrKeyNotFound
                                    
                                Call UniCode_Conv(ITEMREC.SEI_KBN, "")
                                Call UniCode_Conv(ITEMREC.G_ST_URITAN, "")
                                Call UniCode_Conv(ITEMREC.S_KOUSU, "")
                                Call UniCode_Conv(ITEMREC.SEI_SYU_KON, "")
                            
                            
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                Exit Function
                        End Select
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        '����
                        If IsNumeric(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) Then
                            KOUSEI(i).KO_QTY = CDbl(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode))
                        Else
                            KOUSEI(i).KO_QTY = 1#
                        End If
                        '�d���P��
                        If IsNumeric(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode)) Then
                            KOUSEI(i).G_ST_SHITAN = CDbl(StrConv(ITEMREC.G_ST_SHITAN, vbUnicode))
                        Else
                            KOUSEI(i).G_ST_SHITAN = 0#
                        End If
                    
                        '����P��
                        Select Case StrConv(ITEMREC.SEI_KBN, vbUnicode)
                        
                            Case "1"
                               KOUSEI(i).G_ST_URITAN = 0#
                            Case "2"
                               KOUSEI(i).G_ST_URITAN = 0#
                            Case Else
                                If IsNumeric(StrConv(ITEMREC.G_ST_URITAN, vbUnicode)) Then
                                    KOUSEI(i).G_ST_URITAN = CDbl(StrConv(ITEMREC.G_ST_URITAN, vbUnicode))
                                Else
                                    KOUSEI(i).G_ST_URITAN = 0#
                                End If
                        End Select
                        '�d�����z�v
                        KOUSEI(i).G_ST_SHIKIN = 0#
                        For j = 0 To UBound(SHIZAI_T)
                            If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SHIZAI_T(j) Then
                                
                                
                                If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                                    
                                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                                        
                                        If CDbl(KOUSEI(i).KO_QTY) = 0 Then '2010.02.22
                                            KOUSEI(i).G_ST_SHIKIN = 0#
                                        Else
                                            KOUSEI(i).G_ST_SHIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).G_ST_SHITAN)) / CDbl(KOUSEI(i).KO_QTY), 2)
                                        End If
                                    Else
                                        KOUSEI(i).G_ST_SHIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).KO_QTY)) * CDbl(KOUSEI(i).G_ST_SHITAN), 2)
                                    End If
                                End If
                                Exit For
                            End If
                        
                        Next j
                       '������z�v
                        KOUSEI(i).G_ST_URIKIN = 0
                        KOUSEI(i).G_ST_URIKIN_KUSATU = 0
                
                        For j = 0 To UBound(SHIZAI_T)
                       
                            If Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "1" And Trim(StrConv(ITEMREC.SEI_KBN, vbUnicode)) <> "2" Then
                       
                       
                                If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = SHIZAI_T(j) Then
                                    If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                                    
                                    
                                    
                                        If CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) = 0 Then
                                            KOUSEI(i).G_ST_URIKIN = 0#
                                        Else
                                            KOUSEI(i).G_ST_URIKIN = CDbl(ToHalfAdjust(1 / CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), 3))
                                        End If
                                        KOUSEI(i).G_ST_URIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).G_ST_URITAN)) * CDbl(KOUSEI(i).G_ST_URIKIN), 2)
                                    Else
                                        KOUSEI(i).G_ST_URIKIN = ToRoundUp(CCur(CDbl(KOUSEI(i).KO_QTY) * CDbl(KOUSEI(i).G_ST_URITAN)), 2)
                                    End If
                        
                                
                                Else
                               
                                    If KUSATU_F Then
                                
                                        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                                        
                                        
                                            If CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)) = 0 Then
                                                KOUSEI(i).G_ST_URIKIN_KUSATU = 0
                                            Else
                                                KOUSEI(i).G_ST_URIKIN_KUSATU = CDbl(ToHalfAdjust(1 / CCur(StrConv(P_COMPO_K_REC.KO_QTY, vbUnicode)), 3))
                                            End If
                                            KOUSEI(i).G_ST_URIKIN_KUSATU = ToRoundUp(CCur(CDbl(KOUSEI(i).G_ST_URITAN)) * CDbl(KOUSEI(i).G_ST_URIKIN_KUSATU), 2)
                                        
                                        Else
                                            KOUSEI(i).G_ST_URIKIN_KUSATU = ToRoundUp(CCur(CDbl(KOUSEI(i).KO_QTY)) * CDbl(KOUSEI(i).G_ST_URITAN), 2)
                                        End If
                                    
                                    
                                    End If
                                End If
                            End If
                        Next j
                        
                        If StrConv(P_COMPO_K_REC.KO_SYUBETSU, vbUnicode) = GAISO_KBN Then
                            KOUSEI(i).S_KOUSU = 0
                            KOUSEI(i).SEI_SYU_KON = 0
                        Else
                            '��Ǝ���
                            If IsNumeric(StrConv(ITEMREC.S_KOUSU, vbUnicode)) Then
                                KOUSEI(i).S_KOUSU = CDbl(StrConv(ITEMREC.S_KOUSU, vbUnicode))
                            Else
                                KOUSEI(i).S_KOUSU = 0#
                            End If
                            '�W������
                            If IsNumeric(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode)) Then
                                KOUSEI(i).SEI_SYU_KON = CDbl(StrConv(ITEMREC.SEI_SYU_KON, vbUnicode))
                            Else
                                KOUSEI(i).SEI_SYU_KON = 0#
                            End If
                        End If
                    Loop


                    wkInt = 0
                    If KOUSEI_FLG Then
                        For i = 0 To UBound(KOUSEI)
                        
                            For j = 0 To UBound(SHIZAI_T)
                                If KOUSEI(i).KO_SYUBETSU = SHIZAI_T(j) Then
                                    wkInt = wkInt + CInt(ToRoundUp(CCur(KOUSEI(i).S_KOUSU) * CDbl(KOUSEI(i).KO_QTY), 0))
                                    Exit For
                                End If
                        
                            Next j
                        
                        Next i
                    End If
                    
                    wkTANI = wkInt
                    wkQTY = 1
                    MAIN_KOUTEI(1) = wkTANI * wkQTY

                    '�B
                    wkInt = 0
                    If KOUSEI_FLG Then
                        For i = 0 To UBound(KOUSEI)
                        
                            For j = 0 To UBound(DOUKON_T)
                                If KOUSEI(i).KO_SYUBETSU = DOUKON_T(j) Then
                                    
                                    If IsNumeric(KOUSEI(i).KO_QTY) Then
                                        wkInt = wkInt + CInt(ToRoundUp(CCur(KOUSEI(i).KO_QTY), 0))
                                    End If
                                    
                                    
                                    
                                    Exit For
                                End If
                        
                            Next j
                        
                        Next i
                    End If
                    If IsNumeric(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode)) Then
                        wkTANI = CInt(StrConv(P_KANRIREC02.MAIN_KOTEI(3).KOTEI, vbUnicode))
                    Else
                        wkTANI = 0#
                    End If
                    wkQTY = wkInt
                    MAIN_KOUTEI(2) = wkTANI * wkQTY




                    '�C
                    wkInt = 0
                    If KOUSEI_FLG Then
                        For i = 0 To UBound(KOUSEI)
                        
                            For j = 0 To UBound(KAKOU_T)
                                If KOUSEI(i).KO_SYUBETSU = KAKOU_T(j) Then
                                    If IsNumeric(KOUSEI(i).S_KOUSU) Then
                                        wkInt = wkInt + CInt(KOUSEI(i).S_KOUSU)
                                    End If
                                    Exit For
                                End If
                        
                            Next j
                        
                        Next i
                    End If
                    wkTANI = wkInt
                    wkQTY = 1
                    MAIN_KOUTEI(3) = wkTANI * wkQTY
                    
                    '�D
                    wkInt = 0
                    If KOUSEI_FLG Then
                        For i = 0 To UBound(KOUSEI)
                            
                            
                            For j = 0 To UBound(SHIZAI_T)
                            
                                If KOUSEI(i).KO_SYUBETSU = SHIZAI_T(j) Then
                                    If IsNumeric(KOUSEI(i).SEI_SYU_KON) Then
                                        wkInt = wkInt + CInt(KOUSEI(i).SEI_SYU_KON)
                                    End If
                                End If
                            
                            Next j
                            
                        Next i
                    End If
                    wkTANI = wkInt
                    wkQTY = 1
                    MAIN_KOUTEI(4) = wkTANI * wkQTY


                    '�v
                    wkInt = 0
                    For i = 0 To UBound(MAIN_KOUTEI)
                    
                        wkInt = wkInt + MAIN_KOUTEI(i)
                    Next i
                    Write #FileNo, Format(ToHalfAdjust(CCur(wkInt) / 60, 1), "#0.0"),


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''    2011.03.31
                    
                    
'DEBUG                    Shiji_No = ""
'DEBUG                    Write #FileNo, Shiji_No
                    
                    
                    '���������敪   2011.07.04
                    Write #FileNo, StrConv(GOODSREC.NAI_BUHIN, vbUnicode),
                    '�C�O�����敪   2011.07.04
                    Write #FileNo, StrConv(ITEMREC.NAI_BUHIN, vbUnicode),
                    '���i��������z��   2011.07.04
                    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, wkUKEHARAI_CODE)
                    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
                    Select Case sts
                    Case BtNoErr
                        
                        Case BtErrKeyNotFound
                            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_NAME, "")
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�󕥐�}�X�^")
                            Exit Function
                    End Select
                    Write #FileNo, wkUKEHARAI_CODE & " " & StrConv(P_UKEHARAIREC.UKEHARAI_NAME, vbUnicode),
                    
''DEBUG                    Write #FileNo, wkHIN_NAME,
                    
                    Write #FileNo,
                    
                    '2011.07.12
                    Out_Cnt = Out_Cnt + 1
                    
                    
                End If
            End If
            
            
If Right(Format(In_Cnt, "000"), 2) = "00" Then
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "CSV�f�[�^�o�͒��I�I[" & Out_Cnt & "/" & In_Cnt & "]", Me.hwnd, 0)
    DoEvents
End If
            
            com = BtOpGetNext
        End If
    Loop

    Close #FileNo

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                "CSV�f�[�^�o�͊����I�I[" & Out_Cnt & "/" & In_Cnt & "]", Me.hwnd, 0)



    Beep
    DoEvents
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
    
    
        com = BtOpGetNext
    
    Loop
    
    MI_ZAIKO_KENSAKU = False

End Function
' ------------------------------------------------------------------------
'       �w�肵�����x�̐��l�ɐ؂�グ���܂��B
'
' @Param    dValue      �ۂߑΏۂ̔{���x���������_���B
' @Param    iDigits     �߂�l�̗L�������̐��x�B
' @Return               iDigits �ɓ��������x�̐��l�ɐ؂�グ��ꂽ���l�B
' ------------------------------------------------------------------------
Private Function ToRoundUp(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    
        


    dCoef = (10 ^ iDigits)



    
    
    
    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundUp = (Int(dValue * dCoef) + 1) / dCoef
        Case Is < 0
            ToRoundUp = (Fix(dValue * dCoef) - 1) / dCoef
        Case Else
            ToRoundUp = dValue
    End Select


'    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
'        Case Is > 0
'            ToRoundUp = (Int(dValue * dCoef + 0.9)) / dCoef
'        Case Is < 0
'            ToRoundUp = (Fix(dValue * dCoef - 0.9)) / dCoef
'        Case Else
'            ToRoundUp = dValue
'    End Select



End Function

' ------------------------------------------------------------------------
'       �w�肵�����x�̐��l�ɐ؂�̂Ă��܂��B
'
' @Param    dValue      �ۂߑΏۂ̔{���x���������_���B
' @Param    iDigits     �߂�l�̗L�������̐��x�B
' @Return               iDigits �ɓ��������x�̐��l�ɐ؂�̂Ă�ꂽ���l�B
' ------------------------------------------------------------------------
Public Function ToRoundDown(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    Select Case CDbl(dValue * dCoef) - Fix(dValue * dCoef)
        Case Is > 0
            ToRoundDown = Int(dValue * dCoef) / dCoef
        Case Is < 0
            ToRoundDown = Fix(dValue * dCoef) / dCoef
        Case Else
            ToRoundDown = dValue
    End Select
End Function





' ------------------------------------------------------------------------
'       �w�肵�����x�̐��l�Ɏl�̌ܓ����܂��B
'
' @Param    dValue      �ۂߑΏۂ̔{���x���������_���B
' @Param    iDigits     �߂�l�̗L�������̐��x�B
' @Return               iDigits �ɓ��������x�̐��l�Ɏl�̌ܓ����ꂽ���l�B
' ------------------------------------------------------------------------
Public Function ToHalfAdjust(ByVal dValue As Currency, ByVal iDigits As Integer) As Currency
    Dim dCoef As Double

    dCoef = (10 ^ iDigits)

    If dValue > 0 Then
        ToHalfAdjust = Int(CDbl(dValue * dCoef + 0.5)) / dCoef
    Else
        ToHalfAdjust = Fix(CDbl(dValue * dCoef - 0.5)) / dCoef
    End If
End Function

Private Sub Text_LostFocus(Index As Integer)

    Select Case Index
        Case ptxSOKO
            Text(Index).Text = StrConv(Trim(Text(Index).Text), vbUpperCase)
    End Select


End Sub
