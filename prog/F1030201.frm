VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form F1030201 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�`�[�ԍ��w��o�ɕ\����uV1.03�v"
   ClientHeight    =   6840
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   13950
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   13950
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton Command1 
      Caption         =   "�S�đI��"
      Enabled         =   0   'False
      Height          =   375
      Left            =   11340
      TabIndex        =   35
      Top             =   840
      Width           =   1380
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   6780
      MaxLength       =   6
      TabIndex        =   10
      Top             =   600
      Width           =   810
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   1530
      MaxLength       =   4
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   2250
      MaxLength       =   2
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   2730
      MaxLength       =   2
      TabIndex        =   6
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   3810
      MaxLength       =   4
      TabIndex        =   7
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   4530
      MaxLength       =   2
      TabIndex        =   8
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   5010
      MaxLength       =   2
      TabIndex        =   9
      Top             =   600
      Width           =   375
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   2
      Left            =   8100
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   3
      Top             =   120
      Width           =   3360
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   6780
      MaxLength       =   8
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   4410
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   1575
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   120
      Width           =   1095
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
      TabIndex        =   23
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
      TabIndex        =   22
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
      TabIndex        =   21
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "��@��"
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
      TabIndex        =   20
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
      Index           =   7
      Left            =   6480
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   12
      Top             =   5880
      Width           =   855
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   4215
      Left            =   315
      OleObjectBlob   =   "F1030201.frx":0000
      TabIndex        =   11
      Top             =   1320
      Width           =   13140
   End
   Begin VB.Label Label3 
      Height          =   375
      Left            =   12705
      TabIndex        =   38
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   12705
      TabIndex        =   37
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11970
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`�[�ԍ�"
      Height          =   255
      Index           =   11
      Left            =   5745
      TabIndex        =   34
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�ח\���"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   33
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   1
      Left            =   2130
      TabIndex        =   32
      Top             =   720
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   2
      Left            =   2610
      TabIndex        =   31
      Top             =   720
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   240
      Index           =   5
      Left            =   3450
      TabIndex        =   30
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   6
      Left            =   4410
      TabIndex        =   29
      Top             =   720
      Width           =   120
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   240
      Index           =   7
      Left            =   4890
      TabIndex        =   28
      Top             =   720
      Width           =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�א�"
      Height          =   255
      Index           =   10
      Left            =   5985
      TabIndex        =   27
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����敪"
      Height          =   255
      Index           =   4
      Left            =   3360
      TabIndex        =   26
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����敪"
      Height          =   255
      Index           =   3
      Left            =   210
      TabIndex        =   25
      Top             =   240
      Width           =   975
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
      TabIndex        =   24
      Top             =   6360
      Width           =   180
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Checked         =   -1  'True
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1030201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxMUKE_CODE% = 0             '������R�[�h�i����́j

Private Const ptxS_DEN_DT_YY% = 1           '�J�n�@�o�ח\����@�N
Private Const ptxS_DEN_DT_MM% = 2           '�J�n�@�o�ח\����@��
Private Const ptxS_DEN_DT_DD% = 3           '�J�n�@�o�ח\����@��
Private Const ptxE_DEN_DT_YY% = 4           '�I���@�o�ח\����@�N
Private Const ptxE_DEN_DT_MM% = 5           '�I���@�o�ח\����@��
Private Const ptxE_DEN_DT_DD% = 6           '�I���@�o�ח\����@��
Private Const ptxDEN_NO% = 7                '�`�[�ԍ�

Private Const Text_Max% = 7                 '��ʍ��ڕʍő���ޯ��

Private Const pcmbPRINT_KBN% = 0            '����敪
Private Const pcmbCyu_Kbn% = 1              '�����敪
Private Const pcmbMUKE_Code% = 2            '������


Dim SYUKA As New XArrayDB

Private Const Min_Row% = 1              '�ŏ��s��

Dim Max_Row    As Long                  '�O���b�h�ő�\������   2013.12.15 integer-->long

Dim SYUKA_DATA  As String               '�o�׃f�[�^�t���p�X


Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 14             '�ő��



Private Const ColDummy% = 0             '�_�~�[

Private Const ColSEL% = 1               '�I
Private Const ColCyu_Kbn% = 2           '�����敪����
Private Const ColCyu_Kbn_Name% = 3      '�����敪����
Private Const ColMUKE_Code% = 4        '�o�א溰�ށi��\���j
Private Const ColMUKE_Name% = 5         '�o�א於
Private Const ColDEN_DT% = 6            '�`�[���t
Private Const ColID_NO% = 7             '�`�[�h�c
Private Const ColDEN_NO% = 8            '�`�[��
Private Const ColHIN_GAI% = 9           '�i�ځi�O���j
Private Const ColSURYO% = 10            '�o�א��i�\��j
Private Const ColJITU_SURYO% = 11       '�o�א��i���сj
Private Const ColPrint% = 12            '�o�ɕ\����}�[�N
Private Const ColHIN_NAI% = 13          '�i�ځi�����j
Private Const ColHIN_Name% = 14         '�i��


Private Const Print_KBN0$ = "�V�K�@"
Private Const Print_KBN1$ = "�Ĉ��"
Private Const Print_KBN_SIN$ = "0"
Private Const Print_KBN_SAI$ = "1"

Private KASO_NYUKA_SOKO As String * 2       '���z�@���בq�ɔԍ�
Private KASO_SYOHN_SOKO As String * 2       '���z�@���i���q�ɔԍ�
Private KASO_NAI_SOKO As String * 2         '���z�@���E�q�ɔԍ�


Private Const LMAX% = 46                    '�œ��ő�s��
Private Const MGN_L% = 10                   '���]���i�����F�P����j
Private Const MGN_U% = 1                    '��]���i�s���F�P����j

Dim Pdate As String                         '����J�n���t�iͯ�ް�p�j
Dim Ptime As String                         '����J�n�����iͯ�ް�p�j


Dim NormalFont As New StdFont               '����t�H���g
Dim Code39Font As New StdFont               '����t�H���g

Dim NON_MUKE_CODE() As String * 8           '���O������R�[�h
Dim NON_MUKE_FLG    As Boolean

Dim ALL_Check       As Boolean              '�S���Ώ�

Dim Print_Cnt       As Long

Private Sub Combo_Click(Index As Integer)
    Select Case Index
        Case pcmbCyu_Kbn
            
            
            Text(ptxMUKE_CODE).SetFocus
    End Select

End Sub

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
        Case pcmbCyu_Kbn
            Text(ptxMUKE_CODE).SetFocus
        Case pcmbMUKE_Code
            
            Text(ptxMUKE_CODE).Text = Trim(Right(Combo(Index).Text, 16))
            
            
            
            
            If List_Disp_Proc Then
                Unload Me
            End If
    End Select

End Sub


Private Sub Command_Click(Index As Integer)

Dim ans As Integer

    Select Case Index
        
        Case 7                              '����
            If List_Disp_Proc() Then
                Unload Me
            End If
        
                    
        
        
        
        Case 8                              '���
            
            
            
            ans = MsgBox("�u�o�ɕ\�v������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                
                TDBGrid1.Update
                
                If Print_Proc() Then
                    Unload Me
                End If
            
                            
                            
                ALL_Check = False
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
            
            End If
        
        
        Case 11                             '�I��
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub

Private Sub Command1_Click()
    
    If Not ALL_Check Then
        ALL_Check = True
        Command1.Caption = "�S�ĉ���"
    
    Else
        ALL_Check = False
        Command1.Caption = "�S�đI��"
    
    End If

    If List_Disp_Proc() Then
        Unload Me
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
Dim i As Integer
Dim c As String * 128
Dim sts As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If


    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)

                    '�ő�\�������̊l��
    If GetIni(App.EXEName, "LISTMAX", "SYS", c) Then
        Beep
        MsgBox "�ő�\�������̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    Max_Row = CInt(RTrim(c))
                                
                                '���ƕ���荞��
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1030201.Caption = "�`�[�ԍ��w��o�ɕ\����uV1.02�v�i" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
                                '���׉��z�q�Ɏ�荞��
    If GetIni(App.EXEName, "KASO_NYUKA_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_NYUKA_SOKO = RTrim(c)
                                '���i�����z�q�Ɏ�荞��
    If GetIni(App.EXEName, "KASO_SYOHN_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_SYOHN_SOKO = RTrim(c)
                                '���E���z�q�Ɏ�荞��
    If GetIni(App.EXEName, "KASO_NAI_SOKO", "SYS", c) Then
        c = ""
    End If
    KASO_NAI_SOKO = RTrim(c)
                                
                                '���O������R�[�h�l��
    i = 0
    NON_MUKE_FLG = False
    Do
        If GetIni(App.EXEName, "MUKE" & Format(i + 1, "00"), "SYS", c) Then
            Exit Do
        End If
    
        If RTrim(c) = "NON" Then
            Exit Do
        End If
    
        ReDim Preserve NON_MUKE_CODE(0 To i)
    
        NON_MUKE_CODE(i) = RTrim(c)
        NON_MUKE_FLG = True
    
        i = i + 1
    Loop

                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '������Ǘ��}�X�^�n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌ɂn�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F1030201.FontName
        .Size = 10
    End With
                                '����t�H���g�ݒ�i�o�[�R�[�h�j
    With Code39Font
        .NAME = Label1.FontName
        .Size = Label1.FontSize
    End With

    ALL_Check = False

'������ݒ�
    If MTS_Set_Proc() Then
        Unload Me
    End If

                                '��ʏ����ݒ�
    Combo(pcmbPRINT_KBN).AddItem "      " & "   " & " "
    Combo(pcmbPRINT_KBN).AddItem Print_KBN0 & "   " & Print_KBN_SIN
    Combo(pcmbPRINT_KBN).AddItem Print_KBN1 & "   " & Print_KBN_SAI
    Combo(pcmbPRINT_KBN).ListIndex = 0

'���ޏ����ݒ�
    
    Combo(pcmbCyu_Kbn).AddItem "�S��" & "   " & " "
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_1 & "   " & CYU_KBN_TUK
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_2 & "   " & CYU_KBN_SPO
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_3 & "   " & CYU_KBN_HJU
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_4 & "   " & CYU_KBN_TOK
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_E & "   " & CYU_KBN_BOU
    Combo(pcmbCyu_Kbn).AddItem CYU_KBN_T & "   " & CYU_KBN_KIN
    Combo(pcmbCyu_Kbn).ListIndex = 0

    Combo(pcmbPRINT_KBN).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^")
        End If
    End If
                                            '������Ǘ��}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "������Ǘ��}�X�^")
        End If
    End If
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�o�ח\��b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��")
        End If
    End If
                                            '�݌ɂb�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��")
        End If
    End If
    
    sts = BTRV(BtOpReset, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
    If JGYOBU_T(Index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1030201.Caption = "�`�[�ԍ��w��o�ɕ\����uV1.02�v�i" + RTrim(JGYOBU_T(i).NAME) + ")"
    SubMenu(Index).Checked = True
    If Last_JGYOBU <> JGYOBU_T(Index).CODE Then
        Last_JGYOBU = JGYOBU_T(Index).CODE
        LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
        LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

    End If

End Sub

Private Function MTS_Set_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer
Dim Edit        As String

    MTS_Set_Proc = True
    
    Call Input_Lock
    
    
    Combo(pcmbMUKE_Code).Clear
    
    Edit = "�S�o�א�" & "   "
    Edit = Edit & "                "
    Combo(pcmbMUKE_Code).AddItem Edit
    
    
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K1_MTS, Len(K1_MTS), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "������}�X�^")
                Exit Function
        End Select
        
        Edit = StrConv(MTSREC.MUKE_NAME, vbUnicode) & "   "
        Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)
        
        
        Combo(pcmbMUKE_Code).AddItem Edit
    
        com = BtOpGetNext
    
    Loop

    If Combo(pcmbMUKE_Code).ListCount <= 0 Then
    Else
        Combo(pcmbMUKE_Code).ListIndex = 0
    End If

    Call Input_UnLock

    MTS_Set_Proc = False
End Function


Private Function List_Disp_Proc() As Integer

Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
Dim Row         As Long
    
Dim Skip_Flg    As Boolean
    
    
    
Dim wkDEN_No    As String
    
    
    F1030201.MousePointer = vbHourglass
    
    
    List_Disp_Proc = True
                                    
'    Call Input_Lock
                                    
    For i = ptxS_DEN_DT_YY To ptxE_DEN_DT_DD
    
        If IsNumeric(Trim(Text(i).Text)) Then
        
        
            Text(i).Text = Right(Format(CInt(Text(i).Text), "0000"), Text(i).MaxLength)
        
        End If
    
    
    Next i
                                    
    Text(ptxMUKE_CODE).Text = Trim(Right(Combo(pcmbMUKE_Code).Text, 16))
                                    
    If Trim(Text(ptxMUKE_CODE).Text) = "" Then
        Call UniCode_Conv(MTSREC.MUKE_CODE, "")
        Call UniCode_Conv(MTSREC.SS_CODE, "")
    Else
        Call UniCode_Conv(K2_MTS.MUKE_CODE, Text(ptxMUKE_CODE).Text)
        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K2_MTS, Len(K2_MTS), 2)
        Select Case sts
            Case BtNoErr
                            
            Case BtErrKeyNotFound
                            
                Call UniCode_Conv(K3_MTS.SS_CODE, Text(ptxMUKE_CODE).Text)
                                                    
                sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K3_MTS, Len(K3_MTS), 3)
                Select Case sts
                    Case BtNoErr
                                    
                    Case BtErrKeyNotFound
                        Call UniCode_Conv(MTSREC.MUKE_CODE, "")
                        Call UniCode_Conv(MTSREC.SS_CODE, "")
                                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
                        Unload Me
                End Select

            Case Else
                Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
                Unload Me
        End Select
    End If

    For i = 0 To Combo(pcmbMUKE_Code).ListCount - 1 '������

        If Right(Combo(pcmbMUKE_Code).List(i), 16) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode) Then
            Combo(pcmbMUKE_Code).ListIndex = i
            Exit For
        End If
    

    Next
                                    
    '��ǂ�
'    sts = BTRV(BtOpGetFirst, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K2_Y_SYU, Len(K2_Y_SYU), 2)
'
'    Select Case sts
'        Case BtNoErr
'            Skip_Flg = False
'        Case BtErrEOF
'            Skip_Flg = True
'        Case Else
'            Unload Me
'    End Select
                                    
                                        '�e�[�u�����Z�b�g
    Set SYUKA = Nothing
    
    
    
                                    '�o�ח\��ǂݍ��݊J�n
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU) '���ƕ�
                                                    '�����敪
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, "")
                                                    '������
    
    
    Row = Min_Row - 1
        
        
    
    com = BtOpGetGreaterEqual
    
    Do
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    
        Select Case sts
            Case BtNoErr
        
            Case BtErrEOF
                Exit Do
            Case Else
                Unload Me
        End Select
                                '���ƕ� KEY��ڰ�
        If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
            
'Call Log_Out(LOG_F, "1=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
            Exit Do
        End If
        
        Skip_Flg = False
                                
                                '�����敪 KEY��ڰ�
        If Right(Combo(pcmbCyu_Kbn).Text, 1) <> " " Then
            If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> Right(Combo(pcmbCyu_Kbn).Text, 1) Then
                Skip_Flg = True
'Call Log_Out(LOG_F, "2=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
            End If
        End If
                                
                                '������ KEY��ڰ�
        If Trim(Text(ptxMUKE_CODE).Text) <> "" Then
            If Trim(Right(Combo(pcmbMUKE_Code).Text, 16)) <> "" Then
                If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) <> Trim(Left(Right(Combo(pcmbMUKE_Code).Text, 16), 8)) Or _
                    Trim(StrConv(Y_SYUREC.SS_CODE, vbUnicode)) <> Trim(Right(Combo(pcmbMUKE_Code).Text, 8)) Then
                    Skip_Flg = True
'Call Log_Out(LOG_F, "3=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
                End If
            End If
        
        Else
            If NON_MUKE_FLG Then
                For i = 0 To UBound(NON_MUKE_CODE)
                    If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = Trim(NON_MUKE_CODE(i)) Then
                        Skip_Flg = True
'Call Log_Out(LOG_F, "4=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
                        Exit For
                    End If
                Next i
            End If
        End If
            
        
        
                                '����������
        If CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) = CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) Then
            Skip_Flg = True
'''Call Log_Out(LOG_F, "5=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
        End If
                                
                                
                                '����敪
        If Trim(Right(Combo(pcmbPRINT_KBN).Text, 1)) <> "" Then
            If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
                If IsDate(Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 1, 4) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 5, 2) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 7, 2)) Then
                    Skip_Flg = True
'Call Log_Out(LOG_F, "6=" & StrConv(Y_SYUREC.ID_NO, vbUnicode) & "-" & StrConv(Y_SYUREC.PRINT_YMD, vbUnicode))
                End If
            Else
                If Not IsDate(Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 1, 4) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 5, 2) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 7, 2)) Then
                    Skip_Flg = True
                    Skip_Flg = True
'Call Log_Out(LOG_F, "7=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
                End If
            End If
        End If
        
                                '�`�[���t�͈�(�J�n)
        If StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) < (Text(ptxS_DEN_DT_YY).Text & Text(ptxS_DEN_DT_MM).Text & Text(ptxS_DEN_DT_DD).Text) Then
            Skip_Flg = True
'Call Log_Out(LOG_F, "8=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
        End If
                                '�`�[���t�͈�(�I��)
        If Trim(Text(ptxE_DEN_DT_YY).Text & Text(ptxE_DEN_DT_MM).Text & Text(ptxE_DEN_DT_DD).Text) <> "" Then
            If StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) > (Text(ptxE_DEN_DT_YY).Text & Text(ptxE_DEN_DT_MM).Text & Text(ptxE_DEN_DT_DD).Text) Then
                Skip_Flg = True
'Call Log_Out(LOG_F, "9=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
            End If
        End If
                                '�`�[�ԍ�
        If Trim(Text(ptxDEN_NO).Text) <> "" Then
            If Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode)) <> Text(ptxDEN_NO) Then
                Skip_Flg = True
'Call Log_Out(LOG_F, "10=" & StrConv(Y_SYUREC.ID_NO, vbUnicode))
            End If
        Else
'''--->�`�[������������p�~
'''            If IsNumeric(Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode))) Then
'''                wkDEN_No = Trim(Format(CLng(StrConv(Y_SYUREC.DEN_NO, vbUnicode))))
'''            Else
'''                wkDEN_No = Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode))
'''            End If
'''            If Len(wkDEN_No) > 5 Then
'''                Skip_Flg = True
'''            End If
        
        End If
        
        If Skip_Flg Then
        Else
            Row = Row + 1
            If Row > Max_Row Then
                Beep
                MsgBox "�ő�\���s���𒴂��܂����B"
                Exit Do
            End If
                    
            
            
            If Grid_Set_Proc(Row) Then
                Unload Me
            End If
        End If
        
        com = BtOpGetNext
        
        DoEvents
    Loop
                                
                                
    If Row = (Min_Row - 1) Then
                                '�f�[�^�Ȃ�
        Command1.Enabled = False
        ALL_Check = False
    Else
                                'DB�e�[�u�������N
        SYUKA.QuickSort Min_Row, (SYUKA.UpperBound(1)), ColCyu_Kbn, XORDER_ASCEND, XTYPE_STRING, _
                                                            ColMUKE_Code, XORDER_ASCEND, XTYPE_STRING
    
        
        SYUKA.ReDim Min_Row, Row + 1, Min_Col, Max_Col
        SYUKA(Row + 1, ColDummy) = "--------------------------"
        
        Command1.Enabled = True
    
    
    End If
    
    
    
    TDBGrid1.Style.Locked = True
    
    
    
Label2.Caption = Row
    
    
    Set TDBGrid1.Array = SYUKA
    
    
    TDBGrid1.ReBind
    TDBGrid1.Update
    
'    Call Input_UnLock
    F1030201.MousePointer = vbDefault
    
'    Combo(pcmbMUKE_Code).SetFocus
    
    List_Disp_Proc = False

    
End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1030201.MousePointer = vbHourglass

    Call Ctrl_Lock(F1030201)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1030201)


    F1030201.MousePointer = vbDefault

End Sub

Private Function Grid_Set_Proc(Row As Long) As Integer

Dim sts As Integer

    
    Grid_Set_Proc = True

    

    SYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                                                                
    SYUKA(Row, ColSEL) = ALL_Check                              '�I��
                                                                
                                                                '�����敪
    SYUKA(Row, ColCyu_Kbn) = StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
    
    
    Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
        Case CYU_KBN_TUK    '����
            SYUKA(Row, ColCyu_Kbn_Name) = CYU_KBN_1
        Case CYU_KBN_SPO    '�X�|�b�g(�ً})
            SYUKA(Row, ColCyu_Kbn_Name) = CYU_KBN_2
        Case CYU_KBN_HJU    '��[
            SYUKA(Row, ColCyu_Kbn_Name) = CYU_KBN_3
        Case CYU_KBN_TOK    '����(��ďo��)
            SYUKA(Row, ColCyu_Kbn_Name) = CYU_KBN_4
        Case CYU_KBN_BOU    '�f��
            SYUKA(Row, ColCyu_Kbn_Name) = CYU_KBN_E
    End Select
                                                                    
                                                                    '�o�א溰��
    SYUKA(Row, ColMUKE_Code) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)
                                                                    '�o�א於��
    Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
    Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(Y_SYUREC.SS_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)

    Select Case sts
        Case BtNoErr
            SYUKA(Row, ColMUKE_Name) = StrConv(MTSREC.MUKE_NAME, vbUnicode)
        Case BtErrKeyNotFound
            SYUKA(Row, ColMUKE_Name) = StrConv(MTSREC.MUKE_CODE, vbUnicode)
        Case Else
            Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
            Exit Function
    End Select
                                                                    '�`�[���t
    SYUKA(Row, ColDEN_DT) = Left(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 4) & "/" & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" & Right(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 2)
    SYUKA(Row, ColID_NO) = StrConv(Y_SYUREC.ID_NO, vbUnicode)       '�h�c��
    SYUKA(Row, ColDEN_NO) = StrConv(Y_SYUREC.DEN_NO, vbUnicode)     '�`�[��
    SYUKA(Row, ColHIN_GAI) = StrConv(Y_SYUREC.ITEM_NO, vbUnicode)
    SYUKA(Row, ColHIN_GAI) = StrConv(Y_SYUREC.HIN_NO, vbUnicode)    '�i�ԁi�O���j
                                                                    '�o�א��i�\��j
    SYUKA(Row, ColSURYO) = Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#0")
                                                                    '�o�א��i���сj
    SYUKA(Row, ColJITU_SURYO) = Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), "#0")
                                                                    '����敪
    If IsDate(Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 1, 4) & "/" & _
        Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 5, 2) & "/" & _
        Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 7, 2)) Then
        SYUKA(Row, ColPrint) = "��"
    Else
        SYUKA(Row, ColPrint) = ""
    End If
    
    SYUKA(Row, ColHIN_NAI) = StrConv(Y_SYUREC.HIN_NAI, vbUnicode)   '�i�ԁi�����j
                                                                    '�i�ڃ}�X�^�Ǎ���
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            SYUKA(Row, ColHIN_Name) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Case BtErrKeyNotFound
            SYUKA(Row, ColHIN_Name) = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    End Select
    
    Grid_Set_Proc = False
End Function

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim sts As Integer
Dim i   As Integer

    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    Select Case Index
        Case ptxMUKE_CODE
            
            If Trim(Text(Index).Text) = "" Then
                Call UniCode_Conv(MTSREC.MUKE_CODE, "")
                Call UniCode_Conv(MTSREC.SS_CODE, "")
            Else
                Call UniCode_Conv(K2_MTS.MUKE_CODE, Text(Index).Text)
                sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K2_MTS, Len(K2_MTS), 2)
                Select Case sts
                    Case BtNoErr
                        If Len(Trim(StrConv(MTSREC.SS_CODE, vbUnicode))) <> 0 Then
                            Beep
                            MsgBox "���͂������ڂ̓G���[�ł��B(������R�[�h)"
                            Exit Sub
                        End If
                                    
                    Case BtErrKeyNotFound
                                    
                        Call UniCode_Conv(K3_MTS.SS_CODE, Text(Index).Text)
                                                            
                        sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K3_MTS, Len(K3_MTS), 3)
                        Select Case sts
                            Case BtNoErr
                                            
                            Case BtErrKeyNotFound
                                Beep
                                MsgBox "���͂������ڂ̓G���[�ł��B(������R�[�h)"
                                Exit Sub
                                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
                                Unload Me
                        End Select
    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
                        Unload Me
                End Select
            End If

            For i = 0 To Combo(pcmbMUKE_Code).ListCount - 1 '������
    
                If Right(Combo(pcmbMUKE_Code).List(i), 16) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode) Then
                    Combo(pcmbMUKE_Code).ListIndex = i
                    Exit For
                End If
            
    
            Next

    End Select

    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i

End Sub
Private Function Print_Proc() As Integer

Dim Lcnt            As Integer


Dim SAVE_Cyu_Kbn    As String * 1
Dim SAVE_MUKE_CODE  As String * 16
Dim PRI_HIN_GAI     As String * 13
Dim Betu_LOCATION   As String * 8

Dim com             As Integer
Dim sts             As Integer
Dim ans             As Integer
    

Dim SUMI_QTY        As Long
Dim MI_QTY          As Long
Dim ZAIKO_QTY       As Long
Dim TEMP_QTY        As Long
Dim RetBuf          As String
    
Dim RePrint         As Boolean
    
    Print_Proc = True

    Call Input_Lock
    
    
Print_Cnt = 0
    
    
    
    Lcnt = 99
    
    Set Printer.Font = NormalFont
    
    Printer.Orientation = vbPRORLandscape
    Pdate = Date
    Ptime = Time

    com = BtOpGetGreaterEqual
    
    Do
        DoEvents
                                            '�o�ח\��f�[�^�ǂݍ���
        sts = Y_Syu_Get(RePrint, com)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Exit Function
        End Select
                                            
        If Lcnt = 99 Then
            SAVE_Cyu_Kbn = StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
            SAVE_MUKE_CODE = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(Y_SYUREC.SS_CODE, vbUnicode)
        Else
                                            '�����敪�̃u���[�N
            If SAVE_Cyu_Kbn <> StrConv(Y_SYUREC.CYU_KBN, vbUnicode) Then
                Lcnt = LMAX + 1
                SAVE_Cyu_Kbn = StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
            End If
                                            '������̃u���[�N
            If SAVE_MUKE_CODE <> StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(Y_SYUREC.SS_CODE, vbUnicode) Then
                Lcnt = LMAX + 1
                SAVE_MUKE_CODE = StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & StrConv(Y_SYUREC.SS_CODE, vbUnicode)
            End If
        End If

        If Lcnt > LMAX Then                 '�w�b�_�[�R���g���[��
            If Head_Proc(SAVE_Cyu_Kbn, Lcnt) Then
                Exit Function
            End If
            PRI_HIN_GAI = ""
        End If
                                            
        '-----------------------------------------------------  '�P�s��
        If Left(StrConv(Y_SYUREC.HIN_NO, vbUnicode), 13) <> PRI_HIN_GAI Then
            PRI_HIN_GAI = Left(StrConv(Y_SYUREC.HIN_NO, vbUnicode), 13)
                                            '���׈��
                                            
                                            
            Printer.Print Tab(MGN_L - 5);
            If RePrint Then
                Printer.Print "��";
            End If
                                            
            Printer.Print Tab(MGN_L);
                                            '�W���I��
            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2) & "-";
            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2) & "-";
            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 5, 2) & "-";
            Printer.Print Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 7, 2);

            Printer.Print Tab(MGN_L + 13);
                                            '�i��(�O)
            Printer.Print Left(StrConv(Y_SYUREC.HIN_NO, vbUnicode), 13);

            Printer.Print Tab(MGN_L + 27);
                                            '�W���I�@�݌ɐ�
            If Len(Trim(StrConv(Y_SYUREC.HTANABAN, vbUnicode))) = 0 Then
                SUMI_QTY = 0
                MI_QTY = 0
            Else
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                        MI_QTY, _
                                        Last_JGYOBU, _
                                        StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                        StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                        StrConv(Y_SYUREC.HTANABAN, vbUnicode)) Then
                    Exit Function
                End If
            End If
                       
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            
            If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = "S8" Then
                If Left(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 2) = "S8" Then
                                            '�ʒu�I����
                    If Tana_Kensaku(Betu_LOCATION) Then
                        Print_Proc = True
                        Exit Function
                    End If
                
                
                Else
                                            
                    If S8_LOCATION_Proc("S8", Betu_LOCATION) Then
                        Exit Function
                    Else
                        If Trim(Betu_LOCATION) = "" Then
                            If Tana_Kensaku(Betu_LOCATION) Then
                                Print_Proc = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Else
                                            '�ʒu�I����
                If Tana_Kensaku(Betu_LOCATION) Then
                    Print_Proc = True
                    Exit Function
                End If
            
            End If
            
            
            SUMI_QTY = 0
            MI_QTY = 0
            
            If Len(Trim(Betu_LOCATION)) = 0 Then
            Else
                                            '�ʒu�I�@�݌ɐ�
                Printer.Print Tab(MGN_L + 38);
                Printer.Print Left(Betu_LOCATION, 2) & "-" _
                                & Mid(Betu_LOCATION, 3, 2) & "-" _
                                & Mid(Betu_LOCATION, 5, 2) & "-" _
                                & Right(Betu_LOCATION, 2);
                
                If Zaiko_Syukei_Proc(SUMI_QTY, _
                                        MI_QTY, _
                                        Last_JGYOBU, _
                                        StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                        StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                        Betu_LOCATION) Then
                    Exit Function
                End If
            End If
            
            Printer.Print Tab(MGN_L + 49);
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            '���i�������E�݌ɐ�
            Printer.Print Tab(MGN_L + 58);
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    Last_JGYOBU, _
                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                    KASO_SYOHN_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
            TEMP_QTY = SUMI_QTY + MI_QTY
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    Last_JGYOBU, _
                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                    KASO_NAI_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
            ZAIKO_QTY = TEMP_QTY + SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
                                            
                                            '���בq�ɍ݌�
            Printer.Print Tab(MGN_L + 67);
            If Zaiko_Syukei_Proc(SUMI_QTY, _
                                    MI_QTY, _
                                    Last_JGYOBU, _
                                    StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                    StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                    KASO_NYUKA_SOKO & "01" & "01" & "01") Then
                Exit Function
            End If
                        
            ZAIKO_QTY = SUMI_QTY + MI_QTY
            RetBuf = Format(ZAIKO_QTY, "#,##0")
            If Len(RetBuf) < 9 Then
                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
            End If
            Printer.Print RetBuf;
        End If
        
        '2003.06.03�i�����敪�j
'        Printer.Print Tab(MGN_L + 76);
'        Select Case StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
'            Case CYU_KBN_SPO
'                Printer.Print " ��";
'            Case CYU_KBN_HJU
'                Printer.Print " ��";
'            Case Else
'                Printer.Print " �@";
'        End Select
        '2003.06.03
                    
                                            '�`�[��
        Printer.Print Tab(MGN_L + 80);
        Printer.Print Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 6);


        Printer.Print Tab(MGN_L + 90);
        TEMP_QTY = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode) - CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)))
        RetBuf = Format(TEMP_QTY, "#,##0")
        If Len(RetBuf) < 9 Then
            RetBuf = Space(9 - Len(RetBuf)) & RetBuf
        End If
        Printer.Print RetBuf;

        Printer.Print Tab(MGN_L + 110);
                                                '����t�H���g�ݒ�i�b�������R�X�j
        Set Printer.Font = Code39Font
                            '�o�[�R�[�h(*�`�[ID*)
'        Printer.Print "*" & StrConv(Y_SYUREC.JGYOBU, vbUnicode) & StrConv(Y_SYUREC.ID_NO, vbUnicode) & "*";
        Printer.Print "*" & StrConv(Y_SYUREC.ID_NO, vbUnicode) & "*";
                                                '����t�H���g�ݒ�i�ʏ�j
        Set Printer.Font = NormalFont
        
        '-----------------------------------------------------  '�Q�s��
        Printer.Print Tab(MGN_L + 80);
        Printer.Print StrConv(Y_SYUREC.ID_NO, vbUnicode);

        Printer.Print
        Printer.Print
        
        Lcnt = Lcnt + 3

                                                '������t�ݒ�X�V
'        If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
'            Call UniCode_Conv(Y_SYUREC.PRINT_YMD, Format(Now, "YYYYMMDD"))
'
'            Do
'
'                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
'                Select Case sts
'                    Case BtNoErr
'                        Exit Do
'                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'
'                        Beep
'                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
'                        If ans = vbCancel Then
'                            Print_Proc = SYS_CANCEL
'                            Exit Function
'                        End If
'                    Case Else
'                        Call File_Error(sts, BtOpUpdate, "�o�ח\��")
'                        Print_Proc = SYS_ERR
'                        Exit Function
'
'                End Select
'
'
'            Loop
'        End If
        
 Print_Cnt = Print_Cnt + 1
        com = BtOpGetNext
        
    Loop

    If Lcnt <> 99 Then
        Printer.EndDoc
    End If


Label3.Caption = Print_Cnt
    Call Input_UnLock

    Print_Proc = False

End Function
                                    
Private Function Head_Proc(CYU_KBN As String, Lcnt As Integer) As Integer
Dim i               As Integer
Dim sts             As Integer
Dim CYU_KBN_NAME    As String

    Head_Proc = True

    If Lcnt <> 99 Then
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
    
    
    Printer.Print Tab(MGN_L + 41);
    Select Case CYU_KBN
        Case CYU_KBN_TUK            '����
            CYU_KBN_NAME = CYU_KBN_1
        Case CYU_KBN_SPO            '��߯�
            CYU_KBN_NAME = CYU_KBN_2
        Case CYU_KBN_HJU            '��[
            CYU_KBN_NAME = CYU_KBN_3
        Case CYU_KBN_TOK            '������
            CYU_KBN_NAME = CYU_KBN_4
        Case CYU_KBN_BOU            '�f��
            CYU_KBN_NAME = CYU_KBN_E
    End Select
    
    
    
    Printer.Print "�w" & CYU_KBN_NAME & "�x�o�ɕ\";
    
    
    Printer.Print Tab(MGN_L + 91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print                                      '97.10.14

    Printer.Print Tab(MGN_L);
    Printer.Print "������F";
    Printer.Print "[" & StrConv(Y_SYUREC.MUKE_CODE, vbUnicode) & "]" & "[" & StrConv(Y_SYUREC.SS_CODE, vbUnicode) & "]";
    Printer.Print Tab(MGN_L + 30);
    Call UniCode_Conv(K0_MTS.MUKE_CODE, StrConv(Y_SYUREC.MUKE_CODE, vbUnicode))
    Call UniCode_Conv(K0_MTS.SS_CODE, StrConv(Y_SYUREC.SS_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    Select Case sts
        Case BtNoErr
            Printer.Print "[" & StrConv(MTSREC.MUKE_NAME, vbUnicode) & "]";
            Printer.Print "[" & StrConv(MTSREC.SS_NAME, vbUnicode) & "]";
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "������Ǘ��}�X�^")
            Exit Function
    End Select
    
    Set Printer.Font = Code39Font
    
    If Len(Trim(StrConv(Y_SYUREC.SS_CODE, vbUnicode))) <> 0 Then
        Printer.Print "*" & Trim(StrConv(Y_SYUREC.SS_CODE, vbUnicode)) & "*";
    Else
        Printer.Print "*" & Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) & "*";
    End If
    Set Printer.Font = NormalFont
    
    
    Printer.Print
    'Printer.Print                              '97.10.14
'    Printer.Print Tab(MGN_L + 90); "����OK  ";
                                        '����t�H���g�ݒ�
'    Set Printer.Font = Code39Font
'    Printer.Print "*OK*"
'    Set Printer.Font = NormalFont
                                                '97.10.14 �����܂�
    Printer.Print

    Printer.Print Tab(MGN_L);
    Printer.Print "�W���I��";
    Printer.Print Tab(MGN_L + 13);
    Printer.Print "�i�ԁi�O���j";
    Printer.Print Tab(MGN_L + 26);
    Printer.Print "�W���I�݌�";
    Printer.Print Tab(MGN_L + 38);
    Printer.Print "�ʒu�I��";
    Printer.Print Tab(MGN_L + 50);
    Printer.Print "�ʒu�݌�";
    Printer.Print Tab(MGN_L + 59);
    Printer.Print "���i����";
    Printer.Print Tab(MGN_L + 68);
    Printer.Print "���בq��";
    Printer.Print Tab(MGN_L + 80);
    Printer.Print "�`�[��";
    Printer.Print Tab(MGN_L + 93);
    Printer.Print "�o�א�";
    Printer.Print

    Printer.Print

    Lcnt = 8 + MGN_U

    Head_Proc = False
End Function
Private Function Tana_Kensaku(Betu_LOCATION As String) As Integer

Dim sts As Integer

    Tana_Kensaku = True
    
    Betu_LOCATION = ""
    
    Call UniCode_Conv(K6_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K6_ZAIKO.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    Call UniCode_Conv(K6_ZAIKO.NYUKA_DT, "")
    Call UniCode_Conv(K6_ZAIKO.SOKO_NO, "")
    Call UniCode_Conv(K6_ZAIKO.Retu, "")
    Call UniCode_Conv(K6_ZAIKO.Ren, "")
    Call UniCode_Conv(K6_ZAIKO.Dan, "")
    
    Do
        sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K6_ZAIKO, Len(K6_ZAIKO), 6)
        Select Case sts
                Case BtNoErr
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(Y_SYUREC.NAIGAI, vbUnicode) Or _
                    Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(Y_SYUREC.HIN_NO, vbUnicode)) Then
                    Exit Do
                End If
                If StrConv(ZAIKOREC.SOKO_NO, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 1, 2) Or _
                   StrConv(ZAIKOREC.Retu, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 3, 2) Or _
                   StrConv(ZAIKOREC.Ren, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 5, 2) Or _
                   StrConv(ZAIKOREC.Dan, vbUnicode) <> Mid(StrConv(Y_SYUREC.HTANABAN, vbUnicode), 7, 2) Then
                                                '�V�X�e���q�ɂ̔���
                    Call UniCode_Conv(K0_SOKO.SOKO_NO, StrConv(ZAIKOREC.SOKO_NO, vbUnicode))
                    sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
                    Select Case sts
                        Case BtNoErr
                            If StrConv(SOKOREC.SOKO_BUN, vbUnicode) <> BUN_KASO Then
                                Betu_LOCATION = StrConv(ZAIKOREC.SOKO_NO, vbUnicode) & _
                                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                                StrConv(ZAIKOREC.Dan, vbUnicode)
                                Exit Do
                        
                            End If
                        Case BtErrKeyNotFound
                                                '�l�����Ȃ��̂œǂݔ�΂�
                        Case Else
                            Call File_Error(sts, BtOpGetGreater, "�q�Ƀ}�X�^")
                            Exit Function
                    End Select
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetGreater, "�݌Ƀf�[�^")
                Exit Function
        End Select
            
            
    Loop
    
    Tana_Kensaku = False

End Function


Private Function Y_Syu_Get(RePrint As Boolean, com As Integer) As Integer

Dim sts         As Integer
Dim OP          As Integer
Dim ans         As Integer

Dim i           As Integer
Dim Skip_Flg    As Boolean

    
    
    Y_Syu_Get = False
    
    
    
    If com = BtOpGetGreaterEqual Then
                                        '�ŏ��̂j�d�x�Z�b�g
        Call UniCode_Conv(K5_Y_SYU.JGYOBU, Last_JGYOBU)
        If Right(Combo(pcmbCyu_Kbn).Text, 1) <> " " Then
            Call UniCode_Conv(K5_Y_SYU.KEY_CYU_KBN, Right(Combo(pcmbCyu_Kbn).Text, 1))
        Else
            Call UniCode_Conv(K5_Y_SYU.KEY_CYU_KBN, "")
        End If
        Call UniCode_Conv(K5_Y_SYU.KEY_MUKE_CODE, "")
        Call UniCode_Conv(K5_Y_SYU.KEY_SS_CODE, "")
        Call UniCode_Conv(K5_Y_SYU.HTANABAN, "")
        Call UniCode_Conv(K5_Y_SYU.KEY_SYUKA_YMD, "")
        Call UniCode_Conv(K5_Y_SYU.KEY_HIN_NO, "")
    End If
    
    OP = com + BtSNoWait
    
    Do
        Do
            sts = BTRV(OP, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K5_Y_SYU, Len(K5_Y_SYU), 5)
            Select Case sts
                Case BtNoErr
                    '���ƕ�����ڰ�
                    If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                        
                        sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K5_Y_SYU, Len(K5_Y_SYU), 5)
                        If sts Then
                            Call File_Error(sts, BtOpUnlock, "�o�ח\��t�@�C��")
                            Y_Syu_Get = sts
                            Exit Function
                        End If
                        Y_Syu_Get = BtErrEOF
                        Exit Function
                    End If
                    '�w�肪�L��Β����敪������
                    If Right(Combo(pcmbCyu_Kbn).Text, 1) <> " " Then
                        If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> Right(Combo(pcmbCyu_Kbn).Text, 1) Then
                            
                            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K5_Y_SYU, Len(K5_Y_SYU), 5)
                            If sts Then
                                Call File_Error(sts, BtOpUnlock, "�o�ח\��t�@�C��")
                                Y_Syu_Get = sts
                                Exit Function
                            End If
                            
                            Y_Syu_Get = BtErrEOF
                            Exit Function
                        End If
                    End If
                    Exit Do
                Case BtErrEOF
                    Y_Syu_Get = sts
                    Exit Function
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Y_Syu_Get = BtErrEOF
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, OP, "�o�ח\��t�@�C��")
                    Y_Syu_Get = sts
                    Exit Function
            End Select
        
        Loop
                    
        Skip_Flg = False
                                '������ KEY��ڰ�
        If Trim(Text(ptxMUKE_CODE).Text) <> "" Then
            If Trim(Right(Combo(pcmbMUKE_Code).Text, 16)) <> "" Then
                If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) <> Trim(Left(Right(Combo(pcmbMUKE_Code).Text, 16), 8)) Or _
                    Trim(StrConv(Y_SYUREC.SS_CODE, vbUnicode)) <> Trim(Right(Combo(pcmbMUKE_Code).Text, 8)) Then
                    
                    
                    
                    sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K5_Y_SYU, Len(K5_Y_SYU), 5)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "�o�ח\��t�@�C��")
                        Y_Syu_Get = sts
                        Exit Function
                    End If
                    Skip_Flg = True
                End If
            End If
        Else
            If NON_MUKE_FLG Then
                For i = 0 To UBound(NON_MUKE_CODE)
                    If Trim(StrConv(Y_SYUREC.MUKE_CODE, vbUnicode)) = Trim(NON_MUKE_CODE(i)) Then
                        Skip_Flg = True
                        Exit For
                    End If
                Next i
            End If
        End If
        
                                '����������
        If CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) = CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) Then
            Skip_Flg = True
        End If
                                '����敪
        If Trim(Right(Combo(pcmbPRINT_KBN).Text, 1)) <> "" Then
            If Right(Combo(pcmbPRINT_KBN).Text, 1) = Print_KBN_SIN Then
                If IsDate(Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 1, 4) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 5, 2) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 7, 2)) Then
                    Skip_Flg = True
                End If
            Else
                If Not IsDate(Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 1, 4) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 5, 2) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 7, 2)) Then
                    Skip_Flg = True
                End If
            End If
        End If
        
                                '�`�[���t�͈�(�J�n)
        If StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) < (Text(ptxS_DEN_DT_YY).Text & Text(ptxS_DEN_DT_MM).Text & Text(ptxS_DEN_DT_DD).Text) Then
            Skip_Flg = True
        End If
                                '�`�[���t�͈�(�I��)
        If Trim(Text(ptxE_DEN_DT_YY).Text & Text(ptxE_DEN_DT_MM).Text & Text(ptxE_DEN_DT_DD).Text) <> "" Then
            If StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) > (Text(ptxE_DEN_DT_YY).Text & Text(ptxE_DEN_DT_MM).Text & Text(ptxE_DEN_DT_DD).Text) Then
                Skip_Flg = True
            End If
        End If
                                '�`�[�ԍ�
        If Trim(Text(ptxDEN_NO).Text) <> "" Then
            If Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode)) <> Text(ptxDEN_NO) Then
                Skip_Flg = True
            End If
        Else
'''�`�[�������w��A�p�~
'''            If Len(Trim(StrConv(Y_SYUREC.DEN_NO, vbUnicode))) > 5 Then
'''                Skip_Flg = True
'''            End If
        End If
                                
        If Not Skip_Flg Then
                    
            Skip_Flg = True
                    
            For i = Min_Row To SYUKA.UpperBound(1)
        
                If StrConv(Y_SYUREC.ID_NO, vbUnicode) = SYUKA(i, ColID_NO) Then
                    If SYUKA(i, ColSEL) Then
                        Skip_Flg = False
                        Exit For
                    End If
                End If
        
            Next i
            
            If Not Skip_Flg Then
                If Not IsDate(Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 1, 4) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 5, 2) & "/" & _
                    Mid(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode), 7, 2)) Then
                    RePrint = False
            
        
                    Call UniCode_Conv(Y_SYUREC.PRINT_YMD, Format(Now, "YYYYMMDD"))
                    
                    Do
                
                        sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K6_Y_SYU, Len(K6_Y_SYU), 6)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Y_Syu_Get = BtErrEOF
                                    Exit Function
                                End If
                            Case Else
                                Call File_Error(sts, BtOpUpdate, "�o�ח\��")
                                Y_Syu_Get = sts
                                Exit Function
                                
                        End Select
                    Loop
            
                
                Else
                    RePrint = True
                
                
                    sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K5_Y_SYU, Len(K5_Y_SYU), 5)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "�o�ח\��t�@�C��")
                        Y_Syu_Get = sts
                        Exit Function
                    End If
                
                
                End If
            
                Y_Syu_Get = BtNoErr
                Exit Function
            End If
                    
        Else
            
            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K5_Y_SYU, Len(K5_Y_SYU), 5)
            If sts Then
                Call File_Error(sts, BtOpUnlock, "�o�ח\��t�@�C��")
                Y_Syu_Get = sts
                Exit Function
            End If
                    
        End If
                    
                    
    
        OP = BtOpGetNext + BtSNoWait
    
    Loop
End Function


Private Function S8_LOCATION_Proc(SOKO_NO As String, _
                                        Betu_LOCATION As String) As Integer


Dim sts     As Integer


    S8_LOCATION_Proc = SYS_ERR


    Betu_LOCATION = ""


    Call UniCode_Conv(K4_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K4_ZAIKO.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K4_ZAIKO.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    Call UniCode_Conv(K4_ZAIKO.SOKO_NO, SOKO_NO)
    Call UniCode_Conv(K4_ZAIKO.Retu, "")
    Call UniCode_Conv(K4_ZAIKO.Ren, "")
    Call UniCode_Conv(K4_ZAIKO.Dan, "")
    
    sts = BTRV(BtOpGetGreater, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
    Select Case sts
        Case BtNoErr
            If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(Y_SYUREC.NAIGAI, vbUnicode) Or _
                Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(Y_SYUREC.HIN_NO, vbUnicode)) Or _
                StrConv(ZAIKOREC.SOKO_NO, vbUnicode) <> SOKO_NO Then
            Else
                Betu_LOCATION = StrConv(ZAIKOREC.SOKO_NO, vbUnicode) & _
                                StrConv(ZAIKOREC.Retu, vbUnicode) & _
                                StrConv(ZAIKOREC.Ren, vbUnicode) & _
                                StrConv(ZAIKOREC.Dan, vbUnicode)
            End If
        Case BtErrEOF
        Case Else
            Call File_Error(sts, BtOpGetGreater, "�݌Ƀf�[�^")
            Exit Function
    End Select


    S8_LOCATION_Proc = False

End Function


