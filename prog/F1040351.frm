VERSION 5.00
Begin VB.Form F1040351 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�I�ʍ݌Ɉꗗ�\���"
   ClientHeight    =   8250
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   11430
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
   ScaleHeight     =   8250
   ScaleWidth      =   11430
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   3
      Left            =   4410
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   5
      Top             =   1680
      Width           =   2430
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   3885
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1680
      Width           =   540
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   3885
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1080
      Width           =   1380
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   2
      Left            =   9450
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   2
      Top             =   360
      Width           =   852
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   11
      Left            =   6720
      MaxLength       =   13
      TabIndex        =   15
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   10
      Left            =   3885
      MaxLength       =   13
      TabIndex        =   14
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   9
      Left            =   6720
      MaxLength       =   2
      TabIndex        =   13
      Top             =   4080
      Width           =   435
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   8
      Left            =   3885
      MaxLength       =   2
      TabIndex        =   12
      Top             =   4080
      Width           =   435
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   6720
      MaxLength       =   2
      TabIndex        =   11
      Top             =   3480
      Width           =   435
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   3885
      MaxLength       =   2
      TabIndex        =   10
      Top             =   3480
      Width           =   435
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   6720
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2880
      Width           =   435
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   3885
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2880
      Width           =   435
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   6720
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   1
      Top             =   360
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������f"
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
      Left            =   4830
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   6720
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2280
      Width           =   435
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   3360
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   360
      Width           =   1380
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   3885
      MaxLength       =   2
      TabIndex        =   6
      Top             =   2280
      Width           =   435
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
      Left            =   10290
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   9450
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   8610
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   7770
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   6510
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   5670
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   4830
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   3990
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   2625
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   1785
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   945
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6960
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
      Left            =   105
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6960
      Width           =   855
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�݌Ɏ��x"
      Height          =   240
      Index           =   14
      Left            =   2520
      TabIndex        =   46
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�ȍ~"
      Height          =   240
      Index           =   13
      Left            =   5355
      TabIndex        =   45
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�O��ړ���"
      Height          =   255
      Index           =   12
      Left            =   2310
      TabIndex        =   44
      Top             =   1200
      Width           =   1275
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   255
      Index           =   11
      Left            =   8610
      TabIndex        =   43
      Top             =   480
      Width           =   750
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
      Height          =   375
      Left            =   105
      TabIndex        =   42
      Top             =   7440
      Width           =   225
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ԁi�O���j"
      Height          =   255
      Index           =   10
      Left            =   2310
      TabIndex        =   41
      Top             =   4800
      Width           =   1485
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   255
      Index           =   9
      Left            =   5880
      TabIndex        =   40
      Top             =   4800
      Width           =   330
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�@�@�@�i"
      Height          =   255
      Index           =   8
      Left            =   2310
      TabIndex        =   39
      Top             =   4200
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   255
      Index           =   7
      Left            =   5880
      TabIndex        =   38
      Top             =   4080
      Width           =   330
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�@�@�@�A"
      Height          =   255
      Index           =   6
      Left            =   2310
      TabIndex        =   37
      Top             =   3600
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   255
      Index           =   5
      Left            =   5880
      TabIndex        =   36
      Top             =   3600
      Width           =   330
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�I�ԁ@��"
      Height          =   255
      Index           =   4
      Left            =   2310
      TabIndex        =   35
      Top             =   3000
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   34
      Top             =   3000
      Width           =   330
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ד��ʖ���"
      Height          =   255
      Index           =   2
      Left            =   5145
      TabIndex        =   33
      Top             =   480
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������ł�"
      Height          =   255
      Left            =   4935
      TabIndex        =   32
      Top             =   5400
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   31
      Top             =   2400
      Width           =   330
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "��I���"
      Height          =   255
      Index           =   33
      Left            =   2310
      TabIndex        =   30
      Top             =   480
      Width           =   960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "�q�ɇ�"
      Height          =   255
      Index           =   0
      Left            =   2310
      TabIndex        =   29
      Top             =   2400
      Width           =   750
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
Attribute VB_Name = "F1040351"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const ptxDATE% = 0                  '�O��ړ���
Private Const ptxSYUSHI% = 1                '�݌Ɏ��x


Private Const ptxS_SOKO_NO% = 2             '�J�n�@�q�ɇ�
Private Const ptxE_SOKO_NO% = 3             '�I���@�q�ɇ�
Private Const ptxS_RETU% = 4                '�J�n�@�I�ԁ@��
Private Const ptxE_RETU% = 5                '�I���@�I�ԁ@��
Private Const ptxS_REN% = 6                 '�J�n�@�I�ԁ@�A
Private Const ptxE_REN% = 7                 '�I���@�I�ԁ@�A
Private Const ptxS_DAN% = 8                 '�J�n�@�I�ԁ@�i
Private Const ptxE_DAN% = 9                 '�I���@�I�ԁ@�i
Private Const ptxS_HIN_GAI% = 10            '�J�n�@�i�ԁi�O���j
Private Const ptxE_HIN_GAI% = 11            '�J�n�@�i�ԁi�O���j



Private Const Text_Max% = 9                 '��ʍ��ڕʍő���ޯ��

Private Const pcmbTANA_INF% = 0             '��I���
Private Const pcmbDETA% = 1                 '���ד��ʖ��׈��
Private Const pcmbNAIGAI% = 2               '�����O

Private Const pcmbSYUSHI% = 3               '���x

Private Const LMAX% = 44                    '�œ��ő�s��
Private Const MGN_L% = 5                    '���]���i�����F�P����j
Private Const MGN_U% = 1                    '��]���i�s���F�P����j

Dim Pdate As String                         '����J�n���t�iͯ�ް�p�j
Dim Ptime As String                         '����J�n�����iͯ�ް�p�j

Dim NormalFont As New StdFont               '����t�H���g

Dim PRT_CAN As Boolean                      '����r���L�����Z���v��


Private Const TANA_INF_NO$ = "1"            '��I������@�̃��X�g�{�b�N�X���e
Private Const TANA_INF_ALL$ = "2"
Private Const TANA_INF_ONLY$ = "3"
Private Const TANA_INF1$ = "��I����"
Private Const TANA_INF2$ = "��I�L��"
Private Const TANA_INF3$ = "��I�̂�"

Private Const DETA_ON$ = "0"                '���׈�����@�̃��X�g�{�b�N�X���e
Private Const DETA_OFF$ = "1"

Private Const DETA0$ = "���חL��"
Private Const DETA1$ = "���ז���"
Dim TZAIKO_DATA  As String                  '�݌Ƀf�[�^�t���p�X

Private Function Print_Proc() As Integer

Dim Soko_COM        As Integer
Dim TANA_COM        As Integer
Dim ZAIKO_COM       As Integer
Dim sts             As Integer

Dim RetBuf          As String

Dim Sum_Yuko_Z_Qty  As Long
Dim SAVE_NAIGAI     As String * 1
Dim SAVE_HIN_GAI    As String * 13

Dim PRI_TANA        As String * 8
Dim PRI_NAIGAI      As String * 1
Dim PRI_HIN_GAI     As String * 13

Dim LCNT            As Integer
    
Dim SKIP_F          As Boolean
    
    
    Print_Proc = True
'������́u������f�v�ȊO�̃C�x���g�擾�s��
    Call Input_Lock           '��ʍ��ڃ��b�N
    Label1.Visible = True
    Command1.Visible = True
    Command1.Enabled = True

    PRT_CAN = False

'����J�n
    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '�p���̒��ӂ���ɂ��Ĉ��
    Pdate = Date
    Ptime = Time

    LCNT = 99
    
    SAVE_NAIGAI = ""
    SAVE_HIN_GAI = ""
    Sum_Yuko_Z_Qty = 0

    Call UniCode_Conv(K0_SOKO.Soko_No, Text(ptxS_SOKO_NO).Text)
    
    Soko_COM = BtOpGetGreaterEqual

    Do
        DoEvents
        
        sts = BTRV(Soko_COM, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        
        Select Case sts
            Case BtNoErr
                If StrConv(SOKOREC.Soko_No, vbUnicode) > Text(ptxE_SOKO_NO).Text Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, Soko_COM, "�q�Ƀ}�X�^")
                Exit Function
        End Select
        If (StrConv(SOKOREC.JGYOBU, vbUnicode) = Last_JGYOBU Or _
            StrConv(SOKOREC.JGYOBU, vbUnicode) = JGYOBU_NON) Then
            '����Ώۂ̑q�ɁH(���ƕ����w�莖�ƕ��^���ƕ�����)
            If StrConv(SOKOREC.NAIGAI, vbUnicode) = NAIGAI_NON Or _
                Right(Combo(pcmbNAIGAI).Text, 1) = NAIGAI_NON Then
            Else
                If StrConv(SOKOREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                    Exit Do
                End If
            End If
            
            If LCNT <> 99 Then
                LCNT = LMAX + 1
            End If
            
            Call UniCode_Conv(K0_TANA.Soko_No, StrConv(SOKOREC.Soko_No, vbUnicode))
            Call UniCode_Conv(K0_TANA.Retu, Text(ptxS_RETU).Text)
            Call UniCode_Conv(K0_TANA.Ren, Text(ptxS_REN).Text)
            Call UniCode_Conv(K0_TANA.Dan, Text(ptxS_DAN).Text)
            
            TANA_COM = BtOpGetGreaterEqual

            Do
                DoEvents

                sts = BTRV(TANA_COM, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                Select Case sts
                    Case BtNoErr
                        If (StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)) _
                            > (Text(ptxE_SOKO_NO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                            Exit Do
                        End If
                    
                    
                        If StrConv(SOKOREC.Soko_No, vbUnicode) <> StrConv(TANAREC.Soko_No, vbUnicode) Then
                            Exit Do
                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, TANA_COM, "�I�}�X�^")
                        Exit Function
                End Select
                                            '�݌Ƀf�[�^�ǂݍ��݊J�n
                Call UniCode_Conv(K5_ZAIKO.Soko_No, StrConv(TANAREC.Soko_No, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.Retu, StrConv(TANAREC.Retu, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.Ren, StrConv(TANAREC.Ren, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.Dan, StrConv(TANAREC.Dan, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K5_ZAIKO.NAIGAI, NAIGAI_NON)
                Call UniCode_Conv(K5_ZAIKO.HIN_GAI, "")
                Call UniCode_Conv(K5_ZAIKO.NYUKA_DT, "")
                                
                Sum_Yuko_Z_Qty = 0
                SAVE_NAIGAI = ""
                SAVE_HIN_GAI = ""
                                
                ZAIKO_COM = BtOpGetGreater
                
                Do
                    DoEvents
                
                    If PRT_CAN Then
                        Printer.KillDoc
                        Call Input_UnLock   '��ʍ��ڃ��b�N����
                        Label1.Visible = False
                        Command1.Visible = False
                        Print_Proc = False
                        Exit Function
                    End If
                
                    sts = BTRV(ZAIKO_COM, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K5_ZAIKO, Len(K5_ZAIKO), 5)
                    Select Case sts
                        Case BtNoErr
                            If StrConv(ZAIKOREC.Soko_No, vbUnicode) <> StrConv(TANAREC.Soko_No, vbUnicode) Or _
                                StrConv(ZAIKOREC.Retu, vbUnicode) <> StrConv(TANAREC.Retu, vbUnicode) Or _
                                StrConv(ZAIKOREC.Ren, vbUnicode) <> StrConv(TANAREC.Ren, vbUnicode) Or _
                                StrConv(ZAIKOREC.Dan, vbUnicode) <> StrConv(TANAREC.Dan, vbUnicode) Or _
                                StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                                            '�I�ԁ^���ƕ��u���[�N
                                If Len(Trim(SAVE_NAIGAI)) = 0 Then
                                            '�݌ɂ���������
                                    If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_NO Then
                                        If LCNT > LMAX Then
                                            Call Print_Head(LCNT)
                                            PRI_TANA = ""
                                        End If
                                        Printer.Print Tab(MGN_L);
                                        Printer.Print StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode)
                                        Printer.Print
                                        LCNT = LCNT + 2
                                    End If
                                Else
                                    If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_ONLY And _
                                        Right(Combo(pcmbDETA).Text, 1) <> DETA_ON Then
                                        If TOTAL_PRINT(LCNT, PRI_TANA, SAVE_NAIGAI, SAVE_HIN_GAI, Sum_Yuko_Z_Qty) Then
                                            Exit Function
                                        End If
                                    End If
                                        
                                    Printer.Print       '�P�s���s
                                    LCNT = LCNT + 1
                                End If
                                
                                Exit Do
                            
                            End If
                        Case BtErrEOF
                            If Len(Trim(SAVE_NAIGAI)) = 0 Then
                                            '�݌ɂ���������
                                If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_NO Then
                                    If LCNT > LMAX Then
                                        Call Print_Head(LCNT)
                                        PRI_TANA = ""
                                    End If
                                    Printer.Print Tab(MGN_L);
                                    Printer.Print StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode)
                                    Printer.Print
                                    LCNT = LCNT + 2
                                End If
                            Else
                                If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_ONLY And _
                                    Right(Combo(pcmbDETA).Text, 1) <> DETA_ON Then
                                    If TOTAL_PRINT(LCNT, PRI_TANA, SAVE_NAIGAI, SAVE_HIN_GAI, Sum_Yuko_Z_Qty) Then
                                        Exit Function
                                    End If
                                End If
                                    
                                Printer.Print       '�P�s���s
                                LCNT = LCNT + 1
                            
                            End If
                            
                            Exit Do
                        Case Else
                            Call File_Error(sts, ZAIKO_COM, "�݌Ƀf�[�^")
                            Exit Function
                    End Select
                
                                    
                
                    SKIP_F = False
                
                
                    If Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) < Trim(Text(ptxS_HIN_GAI).Text) Or _
                        Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) > Trim(Text(ptxE_HIN_GAI).Text) Then
                        SKIP_F = True
                    End If
                
                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ZAIKOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ZAIKOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ZAIKOREC.HIN_GAI, vbUnicode))
                
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                            SKIP_F = True
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
                
                
                    If Trim(Text(ptxDATE).Text) <> "" Then
                        If StrConv(ITEMREC.LAST_NYU_DT, vbUnicode) < Format(Text(ptxDATE).Text, "YYYYMMDD") And _
                            StrConv(ITEMREC.LAST_SYU_DT, vbUnicode) < Format(Text(ptxDATE).Text, "YYYYMMDD") Then
                            SKIP_F = True
                        End If
                    End If
                
                
                    If Trim(Text(ptxSYUSHI).Text) <> "" Then
                        If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) <> Trim(Text(ptxSYUSHI).Text) Then
                            SKIP_F = True
                        End If
                    End If
                
                
                
                
                
                
                    If (Right(Combo(pcmbNAIGAI).Text, 1) <> NAIGAI_NON And _
                        Right(Combo(pcmbNAIGAI).Text, 1) <> StrConv(ZAIKOREC.NAIGAI, vbUnicode)) Or _
                        SKIP_F Then
                                                '���O�ΏۊO
                    
                                            
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    Else
                            
                        If Right(Combo(pcmbTANA_INF).Text, 1) = TANA_INF_ONLY Then
                                                '��I�̂�
                            Exit Do
                            
                        End If
                            
                        If Len(Trim(SAVE_NAIGAI)) = 0 Then
                            SAVE_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                            SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                        End If
                
                        If SAVE_NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                            SAVE_HIN_GAI <> Left(StrConv(ZAIKOREC.HIN_GAI, vbUnicode), 13) Then
                            If Right(Combo(pcmbDETA).Text, 1) <> DETA_ON Then
                                If TOTAL_PRINT(LCNT, PRI_TANA, SAVE_NAIGAI, SAVE_HIN_GAI, Sum_Yuko_Z_Qty) Then
                                    Exit Function
                                End If
                            End If
                            
                            Printer.Print           '1�s���s
                            LCNT = LCNT + 1
                            
                            SAVE_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                            SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                                                    
                            Sum_Yuko_Z_Qty = 0
                                                     
                            PRI_NAIGAI = ""
                            PRI_HIN_GAI = ""
                            
                        End If
                                                    
                                                    
                        Sum_Yuko_Z_Qty = Sum_Yuko_Z_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                                    
                        If Right(Combo(pcmbDETA).Text, 1) = DETA_ON Then
                                                    '���׈��
                            If LCNT > LMAX Then
                                Call Print_Head(LCNT)
                                PRI_TANA = ""
                            End If
                                '�I��
                            If PRI_TANA <> (StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)) Then
                                Printer.Print Tab(MGN_L);
                                Printer.Print StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode);
                                PRI_TANA = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
                            End If
                                '�����O
                            Printer.Print Tab(MGN_L + 10);
                            If SAVE_NAIGAI = NAIGAI_NAI Then
                                Printer.Print NAIGAI1;
                            Else
                                Printer.Print NAIGAI2;
                            End If
                                '�i��
                            Printer.Print Tab(MGN_L + 18);
                            Printer.Print SAVE_HIN_GAI;
                                '�i��
                            Printer.Print Tab(MGN_L + 39);
                            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                            Call UniCode_Conv(K0_ITEM.NAIGAI, SAVE_NAIGAI)
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, SAVE_HIN_GAI)
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                    Printer.Print LeftB(Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)), 44);
                                
                                
                                
                                
                                Case BtErrKeyNotFound
                                
                                
                                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                                
                                
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                    Exit Function
                            End Select
                                '���ד�
                            Printer.Print Tab(MGN_L + 66);
                            Printer.Print Left(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 4) & "/" & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & Right(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 2);
                                '�i�ԁi�����j
                            Printer.Print Tab(MGN_L + 78);
                            Printer.Print Left(StrConv(ZAIKOREC.HIN_NAI, vbUnicode), 13);
                                
                            If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                Printer.Print "(��)";
                            Else
                                Printer.Print "(��)";
                            End If
                                                        
                                '�L���݌ɐ�
                            Printer.Print Tab(MGN_L + 99);
                            RetBuf = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#,##0")
                            If Len(RetBuf) < 9 Then
                                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                            End If
                            Printer.Print RetBuf;
                                               
                                
                                '�݌v�L���݌ɐ�
                            Printer.Print Tab(MGN_L + 110);
                            RetBuf = Format(Sum_Yuko_Z_Qty, "#,##0")
                            If Len(RetBuf) < 9 Then
                                RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                            End If
                            Printer.Print RetBuf;
                                '�W���I��
                            Printer.Print Tab(MGN_L + 120);
                            Printer.Print StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)
                            
                            
                    
                            LCNT = LCNT + 1
                        End If
                    End If
                    
                    ZAIKO_COM = BtOpGetNext
                
                Loop
                
                
                TANA_COM = BtOpGetNext

            Loop

        End If
    
        Soko_COM = BtOpGetNext
    
    Loop

    If LCNT <> 99 Then
        Printer.EndDoc
    End If

    Call Input_UnLock               '��ʍ��ڃ��b�N����
    Label1.Visible = False
    Command1.Visible = False

    Print_Proc = False

End Function
Private Sub Print_Head(LCNT As Integer)
'�w�b�_���
Dim i As Integer

    If LCNT < 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
                                        '�w�b�_�[�i�P�j
    Printer.Print Tab(3);
    For i = 0 To UBound(JGYOBU_T)
        If Last_JGYOBU = JGYOBU_T(i).CODE Then
            Printer.Print "�i" & RTrim(JGYOBU_T(i).NAME) & "�j";
            Exit For
        End If
    Next i
    Printer.Print Tab(36);
    Printer.Print "������  �I�ʍ݌Ɉꗗ�\  ������";
    Printer.Print Tab(91);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")

    Printer.Print
                                        '�w�b�_�[�i�Q�j
    Printer.Print Tab(MGN_L);
    Printer.Print "�q�ɁF";
    Printer.Print StrConv(SOKOREC.Soko_No, vbUnicode);
    Printer.Print " ";
    Printer.Print StrConv(SOKOREC.SOKO_NAME, vbUnicode);
    
    Printer.Print " ";
    
    If Trim(Text(ptxSYUSHI).Text) <> "" Then
        Printer.Print Text(ptxSYUSHI).Text & ":" & Left(Combo(pcmbSYUSHI).Text, Len(Combo(pcmbSYUSHI).Text) - 3);
    End If
    
    Printer.Print
    Printer.Print

                                        '�w�b�_�[�i�R�j
    Printer.Print Tab(MGN_L);
    Printer.Print "�I��";
    Printer.Print Tab(MGN_L + 10);
    Printer.Print "�����O";
    Printer.Print Tab(MGN_L + 18);
    Printer.Print "�i�ԁi�O���j";
    Printer.Print Tab(MGN_L + 39);
    Printer.Print "�i  ��  ";
    Printer.Print Tab(MGN_L + 66);
    Printer.Print "���ד�";
    Printer.Print Tab(MGN_L + 78);
    Printer.Print "�i�ԁi�����j";
    Printer.Print Tab(MGN_L + 102);
    Printer.Print "�݌ɐ�";
    
    If Right(Combo(pcmbDETA).Text, 1) = DETA_ON Then
        Printer.Print Tab(MGN_L + 113);
        Printer.Print "�݌v��";
    End If
    
    Printer.Print Tab(MGN_L + 120);
    Printer.Print "�W���I��";
    
    
    
    Printer.Print

    Printer.Print

    LCNT = 7 + MGN_U

End Sub
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1040351.MousePointer = vbHourglass

    Call Ctrl_Lock(F1040351)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1040351)


    F1040351.MousePointer = vbDefault

End Sub


Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �R���{�{�b�N�X���́i�j�����c�������j����
'----------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case pcmbTANA_INF           '��I���
            If Right(Combo(pcmbTANA_INF).Text, 1) = TANA_INF_ONLY Then
                Combo(pcmbDETA).Enabled = False
                Combo(pcmbDETA).TabStop = False
                Combo(pcmbNAIGAI).SetFocus
            Else
                Combo(pcmbDETA).Enabled = True
                Combo(pcmbDETA).TabStop = True
                Combo(pcmbDETA).SetFocus
            End If
        Case pcmbDETA               '���ד��ʖ���
            Combo(pcmbNAIGAI).SetFocus
        Case pcmbNAIGAI             '�����O
            Text(ptxDATE).SetFocus
    
        Case pcmbSYUSHI
            Text(ptxSYUSHI).Text = Right(Combo(pcmbSYUSHI).Text, 3)
            Text(ptxS_SOKO_NO).SetFocus
    
    
    End Select


End Sub


Private Sub Combo_LostFocus(Index As Integer)
    Select Case Index
        Case pcmbTANA_INF           '��I���
        Case pcmbDETA               '���ד��ʖ���
        Case pcmbNAIGAI             '�����O
    
        Case pcmbSYUSHI
            Text(ptxSYUSHI).Text = Right(Combo(pcmbSYUSHI).Text, 3)
    
    
    End Select

End Sub

Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        
        
        Case 7                              '�f�[�^
            If Err_Chk() Then
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("�u�I�ʍ݌Ɉꗗ�\�v�f�[�^�o�͂��܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If OUTPUT_Proc() Then
                    Unload Me
                End If
'                Call Clear_Field
            End If
            Combo(pcmbTANA_INF).SetFocus
        
        
        Case 8                              '���
            
            If Err_Chk() Then
                Exit Sub
            End If
            
            Beep
            ans = MsgBox("�u�I�ʍ݌Ɉꗗ�\�v������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Print_Proc() Then
                    Unload Me
                End If
'                Call Clear_Field
            End If
            Combo(pcmbTANA_INF).SetFocus
                    
        Case 11                             '�I��
            Unload Me
        Case Else
            Beep
    End Select
End Sub
Private Sub Command1_Click()
    PRT_CAN = True
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
    LOG_F = Trim(c)
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
            F1040351.Caption = "�I�ԕʍ݌Ɉꗗ�\����i" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
    
    
    

                                '�I�ʍ݌Ƀt�@�C������荞��
    If GetIni("FILE", "TZAIKO_DATA", "SYS", c) Then
        Beep
        MsgBox "�I�ʍ݌Ƀt�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    TZAIKO_DATA = Trim(c)
                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�I�}�X�^�n�o�d�m
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
    '����Ͻ���`
    Call P_CODE_TBL_Proc
        
    '���x�Z�b�g
    If Code_Set_Proc(pcmbSYUSHI, P_KBN03_CD, 1) Then
        Unload Me
    End If
                                
                                
                                
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F1040351.FontName
        .Size = F1040351.FontSize
    End With
    Set Printer.Font = NormalFont
                                '��ʏ����ݒ�
    Combo(pcmbTANA_INF).AddItem TANA_INF1 & "   " & TANA_INF_NO
    Combo(pcmbTANA_INF).AddItem TANA_INF2 & "   " & TANA_INF_ALL
    Combo(pcmbTANA_INF).AddItem TANA_INF3 & "   " & TANA_INF_ONLY
    Combo(pcmbTANA_INF).ListIndex = 0
    
    Combo(pcmbDETA).AddItem DETA0 & "   " & DETA_ON
    Combo(pcmbDETA).AddItem DETA1 & "   " & DETA_OFF
    Combo(pcmbDETA).ListIndex = 0
    
    Combo(pcmbNAIGAI).AddItem NAIGAI0 & "   " & NAIGAI_NON
    Combo(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo(pcmbNAIGAI).ListIndex = 0
    
    Combo(pcmbTANA_INF).SetFocus
    
    

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
                                            '�I�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�}�X�^")
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
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1040351 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
    If JGYOBU_T(Index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1040351.Caption = "�I�ʍ݌Ɉꗗ�\����i" + RTrim(JGYOBU_T(Index).NAME) + ")"
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

    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i
End Sub



Private Function Err_Chk()
    
Dim i As Integer
    
    Err_Chk = True



'���t
    If Trim(Text(ptxDATE).Text) = "" Then
    Else
        If Not IsDate(Text(ptxDATE).Text) Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B"
            Text(ptxS_SOKO_NO).SetFocus
            Exit Function
        End If
    End If
'�q�ɔԍ�

    If Len(Text(ptxE_SOKO_NO).Text) = 0 Then
        Text(ptxE_SOKO_NO).Text = "zz"
    End If

    If Text(ptxS_SOKO_NO).Text > Text(ptxE_SOKO_NO).Text Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(ptxS_SOKO_NO).SetFocus
        Exit Function
    End If

'�I��
    For i = ptxS_RETU To ptxE_DAN
        Select Case i
            Case ptxS_RETU, ptxS_REN, ptxS_DAN
                If Len(Text(i).Text) = 0 Then
                    Text(i).Text = "00"
                End If
            Case ptxE_RETU, ptxE_REN, ptxE_DAN
                If Len(Text(i).Text) = 0 Then
                    Text(i).Text = "99"
                End If
        End Select
        If IsNumeric(Text(i).Text) Then
            Text(i).Text = Format(CInt(Text(i).Text), "00")
        End If
    Next i


    If Text(ptxS_RETU).Text & Text(ptxS_REN).Text & Text(ptxS_DAN).Text _
        > Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(ptxS_RETU).SetFocus
        Exit Function
    End If
'�i��(�O��)
    If Len(Text(ptxE_HIN_GAI).Text) = 0 Then
        Text(ptxE_HIN_GAI).Text = String(Len(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)), "z")
    End If

    If Text(ptxS_HIN_GAI).Text > Text(ptxE_HIN_GAI).Text Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(ptxS_HIN_GAI).SetFocus
        Exit Function
    End If
    
    Err_Chk = False

End Function

Private Function TOTAL_PRINT(LCNT As Integer, _
                                PRI_TANA As String, _
                                SAVE_NAIGAI As String, _
                                SAVE_HIN_GAI As String, _
                                Sum_Yuko_Z_Qty As Long) As Integer

Dim sts     As Integer
Dim RetBuf  As String
    
    TOTAL_PRINT = True
    
    If LCNT > LMAX Then
        Call Print_Head(LCNT)
        PRI_TANA = ""
    End If
                                '�I��
    If PRI_TANA <> (StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)) Then
        Printer.Print Tab(MGN_L);
        Printer.Print StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode);
        PRI_TANA = StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)
    End If
                                '�����O
    Printer.Print Tab(MGN_L + 10);
    If SAVE_NAIGAI = NAIGAI_NAI Then
        Printer.Print NAIGAI1;
    Else
        Printer.Print NAIGAI2;
    End If
                                '�i��
    Printer.Print Tab(MGN_L + 18);
    Printer.Print SAVE_HIN_GAI;
                                '�i��
    Printer.Print Tab(MGN_L + 39);
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, SAVE_NAIGAI)
    Call UniCode_Conv(K0_ITEM.HIN_GAI, SAVE_HIN_GAI)
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            Printer.Print Left(StrConv(ITEMREC.HIN_NAME, vbUnicode), 25);
        Case BtErrKeyNotFound
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    End Select
                                '�L���݌ɐ�
    Printer.Print Tab(MGN_L + 99);
    RetBuf = Format(Sum_Yuko_Z_Qty, "#,##0")
    If Len(RetBuf) < 9 Then
        RetBuf = Space(9 - Len(RetBuf)) & RetBuf
    End If
    Printer.Print RetBuf;
                                '�W���I��
    Printer.Print Tab(MGN_L + 120);
    Printer.Print StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)

    LCNT = LCNT + 1
                    
    TOTAL_PRINT = False
                    
                    
End Function
Private Function OUTPUT_Proc() As Integer
    
Dim sts             As Integer
Dim Soko_COM        As Integer
Dim TANA_COM        As Integer
Dim ZAIKO_COM       As Integer
Dim Ret             As Integer
    
Dim Sum_Yuko_Z_Qty  As Long
Dim SAVE_HIN_GAI    As String * 13
Dim SAVE_NAIGAI     As String * 1

Dim FileNo          As Long
Dim fileName        As String

Dim c               As String * 128
Dim Soko_No         As String * 2

Dim SKIP_F          As Boolean

    
    OUTPUT_Proc = True
'���s�����̓C�x���g�擾�s��
    Call Input_Lock         '��ʍ��ڃ��b�N

    FileNo = FreeFile
    fileName = TZAIKO_DATA
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Error_Proc
    
    Open (fileName) For Output As FileNo

    If Trim(Text(ptxSYUSHI).Text) <> "" Then
        Write #FileNo, " " & Trim(Text(ptxSYUSHI).Text) & "�F" & Trim(Left(Combo(pcmbSYUSHI).Text, Len(Combo(pcmbSYUSHI).Text) - 3))
    End If



    Write #FileNo, "�I��", "�����O", "�i�ԁi�O�j", "�i��", "���ד�", "�i�ԁi���j", "���^����", "�݌ɐ�", "�݌v��", "�W���I��"



    Sum_Yuko_Z_Qty = 0

    Call UniCode_Conv(K0_SOKO.Soko_No, Text(ptxS_SOKO_NO).Text)
    
    Soko_COM = BtOpGetGreaterEqual

    Do
        DoEvents
        
        sts = BTRV(Soko_COM, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        
        Select Case sts
            Case BtNoErr
                If StrConv(SOKOREC.Soko_No, vbUnicode) > Text(ptxE_SOKO_NO).Text Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, Soko_COM, "�q�Ƀ}�X�^")
                Exit Function
        End Select
        If (StrConv(SOKOREC.JGYOBU, vbUnicode) = Last_JGYOBU Or _
            StrConv(SOKOREC.JGYOBU, vbUnicode) = JGYOBU_NON) Then
            '����Ώۂ̑q�ɁH(���ƕ����w�莖�ƕ��^���ƕ�����)
            If StrConv(SOKOREC.NAIGAI, vbUnicode) = NAIGAI_NON Or _
                Right(Combo(pcmbNAIGAI).Text, 1) = NAIGAI_NON Then
            Else
                If StrConv(SOKOREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                    Exit Do
                End If
            End If
            
            Call UniCode_Conv(K0_TANA.Soko_No, StrConv(SOKOREC.Soko_No, vbUnicode))
            Call UniCode_Conv(K0_TANA.Retu, Text(ptxS_RETU).Text)
            Call UniCode_Conv(K0_TANA.Ren, Text(ptxS_REN).Text)
            Call UniCode_Conv(K0_TANA.Dan, Text(ptxS_DAN).Text)
            
            TANA_COM = BtOpGetGreaterEqual

            Do
                DoEvents

                sts = BTRV(TANA_COM, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                Select Case sts
                    Case BtNoErr
                        If (StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)) _
                            > (Text(ptxE_SOKO_NO).Text & Text(ptxE_RETU).Text & Text(ptxE_REN).Text & Text(ptxE_DAN).Text) Then
                            Exit Do
                        End If
                    
                    
                        If StrConv(SOKOREC.Soko_No, vbUnicode) <> StrConv(TANAREC.Soko_No, vbUnicode) Then
                            Exit Do
                        End If
                    
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        Call File_Error(sts, TANA_COM, "�I�}�X�^")
                        Exit Function
                End Select
                                            '�݌Ƀf�[�^�ǂݍ��݊J�n
                Call UniCode_Conv(K5_ZAIKO.Soko_No, StrConv(TANAREC.Soko_No, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.Retu, StrConv(TANAREC.Retu, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.Ren, StrConv(TANAREC.Ren, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.Dan, StrConv(TANAREC.Dan, vbUnicode))
                Call UniCode_Conv(K5_ZAIKO.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K5_ZAIKO.NAIGAI, NAIGAI_NON)
                Call UniCode_Conv(K5_ZAIKO.HIN_GAI, "")
                Call UniCode_Conv(K5_ZAIKO.NYUKA_DT, "")
                                
                Sum_Yuko_Z_Qty = 0
                SAVE_NAIGAI = ""
                SAVE_HIN_GAI = ""
                                
                ZAIKO_COM = BtOpGetGreater
                
                Do
                    DoEvents
                
                
                    sts = BTRV(ZAIKO_COM, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K5_ZAIKO, Len(K5_ZAIKO), 5)
                    Select Case sts
                        Case BtNoErr
                            If StrConv(ZAIKOREC.Soko_No, vbUnicode) <> StrConv(TANAREC.Soko_No, vbUnicode) Or _
                                StrConv(ZAIKOREC.Retu, vbUnicode) <> StrConv(TANAREC.Retu, vbUnicode) Or _
                                StrConv(ZAIKOREC.Ren, vbUnicode) <> StrConv(TANAREC.Ren, vbUnicode) Or _
                                StrConv(ZAIKOREC.Dan, vbUnicode) <> StrConv(TANAREC.Dan, vbUnicode) Or _
                                StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                                            '�I�ԁ^���ƕ��u���[�N
                                If Len(Trim(SAVE_NAIGAI)) = 0 Then
                                            '�݌ɂ���������
                                    If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_NO Then
                                    
                                                                        
                                        Write #FileNo, StrConv(TANAREC.Soko_No, vbUnicode) & "-" & StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode)
                                    
                                    
                                    
                                    End If
                                Else
                                    If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_ONLY And _
                                        Right(Combo(pcmbDETA).Text, 1) <> DETA_ON Then
                                                    
                                        Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                                        Call UniCode_Conv(K0_ITEM.NAIGAI, SAVE_NAIGAI)
                                        Call UniCode_Conv(K0_ITEM.HIN_GAI, SAVE_HIN_GAI)
                                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr
                                            Case BtErrKeyNotFound
                                                
                                                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                                                Call UniCode_Conv(ITEMREC.ST_REN, "")
                                                Call UniCode_Conv(ITEMREC.ST_DAN, "")
                                            
                                            
                                                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                                            
                                            Case Else
                                                Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                                Exit Function
                                        End Select
                                                    
                                        
                                        
                                        
                                        Write #FileNo, StrConv(TANAREC.Soko_No, vbUnicode) & "-" & StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode), SAVE_NAIGAI, SAVE_HIN_GAI, StrConv(ITEMREC.HIN_NAME, vbUnicode), StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)
                                    
                                    End If
                                End If
                                
                                Exit Do
                            
                            End If
                        Case BtErrEOF
                            If Len(Trim(SAVE_NAIGAI)) = 0 Then
                                            '�݌ɂ���������
                                If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_NO Then
                                    
                                                                        
                                    Write #FileNo, StrConv(TANAREC.Soko_No, vbUnicode) & "-" & StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode)
                                End If
                            Else
                                If Right(Combo(pcmbTANA_INF).Text, 1) <> TANA_INF_ONLY And _
                                    Right(Combo(pcmbDETA).Text, 1) <> DETA_ON Then
                                    
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, SAVE_NAIGAI)
                                    Call UniCode_Conv(K0_ITEM.HIN_GAI, SAVE_HIN_GAI)
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        Case BtErrKeyNotFound
                                        
                                            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                                        
                                            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                            Call UniCode_Conv(ITEMREC.ST_RETU, "")
                                            Call UniCode_Conv(ITEMREC.ST_REN, "")
                                            Call UniCode_Conv(ITEMREC.ST_DAN, "")
                                        
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                            Exit Function
                                    End Select
                                                
                                    
                                        
                                    Write #FileNo, StrConv(TANAREC.Soko_No, vbUnicode) & "-" & StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode), SAVE_NAIGAI, SAVE_HIN_GAI, StrConv(ITEMREC.HIN_NAME, vbUnicode), StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)
                                    
                                    
                                    
                                End If
                            
                            End If
                            
                            Exit Do
                        Case Else
                            Call File_Error(sts, ZAIKO_COM, "�݌Ƀf�[�^")
                            Exit Function
                    End Select
                
                
                
                    SKIP_F = False
If Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) = "B015" Then
 Debug.Print StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
End If
                
                    If Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) < Trim(Text(ptxS_HIN_GAI).Text) Or _
                        Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) > Trim(Text(ptxE_HIN_GAI).Text) Then
                        SKIP_F = True
                    End If
                
                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ZAIKOREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ZAIKOREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ZAIKOREC.HIN_GAI, vbUnicode))
                
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                    Select Case sts
                        Case BtNoErr
                        
                        Case BtErrKeyNotFound
                            SKIP_F = True
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Function
                    End Select
                
                
                    If Trim(Text(ptxDATE).Text) <> "" Then
                        If StrConv(ITEMREC.LAST_NYU_DT, vbUnicode) < Format(Text(ptxDATE).Text, "YYYYMMDD") And _
                            StrConv(ITEMREC.LAST_SYU_DT, vbUnicode) < Format(Text(ptxDATE).Text, "YYYYMMDD") Then
                            SKIP_F = True
                        End If
                    End If
                
                
                    If Trim(Text(ptxSYUSHI).Text) <> "" Then
                        If Trim(StrConv(ITEMREC.G_SYUSHI, vbUnicode)) <> Trim(Text(ptxSYUSHI).Text) Then
                            SKIP_F = True
                        End If
                    End If
                
                
                
                
                
                
                
                    If (Right(Combo(pcmbNAIGAI).Text, 1) <> NAIGAI_NON And _
                        Right(Combo(pcmbNAIGAI).Text, 1) <> StrConv(ZAIKOREC.NAIGAI, vbUnicode)) Or _
                        SKIP_F Then
                                                '���O�ΏۊO
                    Else
                            
                        If Right(Combo(pcmbTANA_INF).Text, 1) = TANA_INF_ONLY Then
                                                '��I�̂�
                            Exit Do
                            
                        End If
                            
                        If Len(Trim(SAVE_NAIGAI)) = 0 Then
                            SAVE_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                            SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                        End If
                
                        If Trim(SAVE_NAIGAI) <> Trim(StrConv(ZAIKOREC.NAIGAI, vbUnicode)) Or _
                            Trim(SAVE_HIN_GAI) <> Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) Then
                            If Right(Combo(pcmbDETA).Text, 1) <> DETA_ON Then
                                Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                                Call UniCode_Conv(K0_ITEM.NAIGAI, SAVE_NAIGAI)
                                Call UniCode_Conv(K0_ITEM.HIN_GAI, SAVE_HIN_GAI)
                                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                Select Case sts
                                    Case BtNoErr
                                    Case BtErrKeyNotFound
                                        Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                        Call UniCode_Conv(ITEMREC.ST_RETU, "")
                                        Call UniCode_Conv(ITEMREC.ST_REN, "")
                                        Call UniCode_Conv(ITEMREC.ST_DAN, "")
                                        Call UniCode_Conv(ITEMREC.HIN_NAME, "")

                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                        Exit Function
                                End Select


                                
''                                Write #FileNo, StrConv(TANAREC.Soko_No, vbUnicode) & "-" & StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode), SAVE_NAIGAI, SAVE_HIN_GAI, StrConv(ITEMREC.HIN_NAME, vbUnicode), StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)
                            
                            End If
                            
                            SAVE_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                            SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                                                
                            Sum_Yuko_Z_Qty = 0
                                                     
                            
                                                    
                                                    
                        End If
                    End If
                    
                    
                    If (Right(Combo(pcmbNAIGAI).Text, 1) <> NAIGAI_NON And _
                        Right(Combo(pcmbNAIGAI).Text, 1) <> StrConv(ZAIKOREC.NAIGAI, vbUnicode)) Or _
                        SKIP_F Then
                                                '���O�ΏۊO
                    Else
                            
                        If Right(Combo(pcmbTANA_INF).Text, 1) = TANA_INF_ONLY Then
                                                '��I�̂�
                            Exit Do
                            
                        End If
                            
                        If Len(Trim(SAVE_NAIGAI)) = 0 Then
                            SAVE_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                            SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                        End If
                
                        If SAVE_NAIGAI <> StrConv(ZAIKOREC.NAIGAI, vbUnicode) Or _
                            SAVE_HIN_GAI <> Left(StrConv(ZAIKOREC.HIN_GAI, vbUnicode), 13) Then
                            If Right(Combo(pcmbDETA).Text, 1) <> DETA_ON Then
'                                If TOTAL_OUTPUT(FileNo, SAVE_NAIGAI, SAVE_HIN_GAI, Sum_Yuko_Z_Qty) Then
'                                    Exit Function
'                                End If
                            End If
                            
                            
                            SAVE_NAIGAI = StrConv(ZAIKOREC.NAIGAI, vbUnicode)
                            SAVE_HIN_GAI = StrConv(ZAIKOREC.HIN_GAI, vbUnicode)
                                                    
                            Sum_Yuko_Z_Qty = 0
                                                     
                            
                        End If
                                                    
                                                    
                        Sum_Yuko_Z_Qty = Sum_Yuko_Z_Qty + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                                                    
                        If Right(Combo(pcmbDETA).Text, 1) = DETA_ON Then
                                '�I��
                            Write #FileNo, StrConv(TANAREC.Soko_No, vbUnicode) & "-" & StrConv(TANAREC.Retu, vbUnicode) & "-" & StrConv(TANAREC.Ren, vbUnicode) & "-" & StrConv(TANAREC.Dan, vbUnicode),
                                '�����O
                            If SAVE_NAIGAI = NAIGAI_NAI Then
                                Write #FileNo, NAIGAI1,
                            Else
                                Write #FileNo, NAIGAI2,
                            End If
                                '�i��
                            Write #FileNo, SAVE_HIN_GAI,
                                '�i��
                            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
                            Call UniCode_Conv(K0_ITEM.NAIGAI, SAVE_NAIGAI)
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, SAVE_HIN_GAI)
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                            Select Case sts
                                Case BtNoErr
                                    Write #FileNo, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)),
                                
                                
                                
                                
                                Case BtErrKeyNotFound
                                
                                
                                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                                    Write #FileNo, ,
                                
                                
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                    Exit Function
                            End Select
                                '���ד�
                            Write #FileNo, Left(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 4) & "/" & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & Right(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 2),
                                '�i�ԁi�����j
                            Write #FileNo, Left(StrConv(ZAIKOREC.HIN_NAI, vbUnicode), 13),
                                
                            If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = "0" Then
                                Write #FileNo, "(��)",
                            Else
                                Write #FileNo, "(��)",
                            End If
                                                        
                                '�L���݌ɐ�
                            Write #FileNo, Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#,##0"),
                                '�݌v�L���݌ɐ�
                            Write #FileNo, Format(Sum_Yuko_Z_Qty, "#,##0"),
                                '�W���I��
                            Write #FileNo, StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode)
                            
                            
                    
                        End If
                    End If
                    
                    
                    
                    
                    
                    ZAIKO_COM = BtOpGetNext
                
                Loop
                
                
                TANA_COM = BtOpGetNext

            Loop

        End If
    
        Soko_COM = BtOpGetNext
    
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


Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
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
    
    Combo(Index).Clear
    
    For i = 0 To UBound(P_KBN_TBL)
    
        If KBN = P_KBN_TBL(i).KBN_CD Then
            Key_Len = P_KBN_TBL(i).KBN_Len
            Exit For
        End If
    
    Next i
    
    If i > UBound(P_KBN_TBL) Then
        Exit Function
    End If
    
    If Mode = 1 Then
        Combo(Index).AddItem Space(Key_Len)
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
        
        
        
        Combo(Index).AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    



End Function


