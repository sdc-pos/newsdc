VERSION 5.00
Begin VB.Form F9999991 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�g�p���J��"
   ClientHeight    =   6960
   ClientLeft      =   2025
   ClientTop       =   2265
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
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   11295
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   4200
      TabIndex        =   19
      Top             =   840
      Width           =   4935
      Begin VB.OptionButton Option1 
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   24
         Top             =   2040
         Width           =   375
      End
      Begin VB.OptionButton Option1 
         Height          =   495
         Index           =   0
         Left            =   480
         TabIndex        =   23
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   2400
         MaxLength       =   8
         TabIndex        =   0
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   1
         Left            =   2400
         MaxLength       =   13
         TabIndex        =   1
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   2
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   2
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   3
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   3
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   4
         Left            =   3360
         MaxLength       =   2
         TabIndex        =   4
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   5
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   5
         Top             =   2640
         Width           =   375
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4560
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "�h�c�ԍ�"
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   22
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "�i�@�@��"
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   21
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "�I�@�@��"
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   20
         Top             =   2640
         Width           =   1335
      End
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
      Index           =   10
      Left            =   9480
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
      Index           =   9
      Left            =   8640
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
      Index           =   8
      Left            =   7800
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
      Index           =   7
      Left            =   6480
      TabIndex        =   13
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
      Index           =   5
      Left            =   4800
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
      Index           =   4
      Left            =   3960
      TabIndex        =   10
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
      Index           =   2
      Left            =   1800
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
      Index           =   1
      Left            =   960
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�X�V"
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
      TabIndex        =   6
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
      Top             =   6480
      Width           =   180
   End
End
Attribute VB_Name = "F9999991"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        Case 0                              '�f�[�^
            
'            Beep
'            ans = MsgBox("�J�����܂����H", vbYesNo + vbQuestion, "�m�F����")
'            If ans = vbYes Then
            
                Call Update_Proc
            
'            End If
                    
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

Private Sub Form_Load()

Dim i   As Integer
Dim c   As String * 128
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


                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�o�ח\��f�[�^�n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Option1(0).Value = True
    Text1(0).SetFocus
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
                                            '�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F9999991 = Nothing

    End
End Sub

Public Sub Update_Proc()


Dim sts     As Integer
Dim ans     As Integer
Dim com     As Integer

Dim OK_FLG  As Integer

    
    If Option1(0).Value Then
    
        Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Format(CLng(Text1(0).Text), "00000000"))
    
        Do
            sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            Select Case sts
                Case BtNoErr
                    If Len(Trim(StrConv(Y_SYUREC.WEL_ID, vbUnicode))) = 0 Then
                    Else
                        ans = MsgBox("�J�����܂����H[" & StrConv(Y_SYUREC.WEL_ID, vbUnicode) & "][" & StrConv(Y_SYUREC.PRG_ID, vbUnicode) & "]", vbYesNo, "�m�F����")
                        If ans = vbYes Then
                            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
                            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
                                                
                                                
                            Do
                                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case Else
                                        MsgBox "�ُ픭���I�I�@sts = " & sts
                                        Unload Me
                                End Select
                            Loop
                        End If
                    End If
            
                    Exit Do
            
                Case BtErrKeyNotFound
                    MsgBox "�Y���\��Ȃ��I�I"
                    Exit Sub
                Case Else
                    MsgBox "�ُ픭���I�I�@sts = " & sts
                    Unload Me
            End Select
        
        Loop
    End If

    If Option1(1).Value Then
        Call UniCode_Conv(K4_ZAIKO.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K4_ZAIKO.NAIGAI, "1")
        Call UniCode_Conv(K4_ZAIKO.HIN_GAI, Text1(1).Text)
        Call UniCode_Conv(K4_ZAIKO.Soko_No, Text1(2).Text)
        Call UniCode_Conv(K4_ZAIKO.Retu, Text1(3).Text)
        Call UniCode_Conv(K4_ZAIKO.Ren, Text1(4).Text)
        Call UniCode_Conv(K4_ZAIKO.Dan, Text1(5).Text)
        
        com = BtOpGetGreaterEqual
        
        OK_FLG = 0
        
        Do
            DoEvents
            Do
                sts = BTRV(com + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)
                Select Case sts
                    Case BtNoErr
                        
                        
                        
                        If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                            StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> "1" Or _
                            Trim(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)) <> Trim(Text1(1).Text) Or _
                            StrConv(ZAIKOREC.Soko_No, vbUnicode) <> Text1(2).Text Or _
                            StrConv(ZAIKOREC.Retu, vbUnicode) <> Text1(3).Text Or _
                            StrConv(ZAIKOREC.Ren, vbUnicode) <> Text1(4).Text Or _
                            StrConv(ZAIKOREC.Dan, vbUnicode) <> Text1(5).Text Then
                            sts = BtErrEOF
                            
                            Exit Do
                        
                        End If
                        
                        If Len(Trim(StrConv(ZAIKOREC.WEL_ID, vbUnicode))) = 0 Then
                        Else
                            ans = MsgBox("�J�����܂����H[" & StrConv(ZAIKOREC.WEL_ID, vbUnicode) & "][" & StrConv(ZAIKOREC.PRG_ID, vbUnicode) & "]", vbYesNo, "�m�F����")
                            If ans = vbYes Then
                                OK_FLG = 1
                                Call UniCode_Conv(ZAIKOREC.WEL_ID, "")
                                Call UniCode_Conv(ZAIKOREC.PRG_ID, "")
                                Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)
                                Do
                                    sts = BTRV(BtOpUpdate, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case Else
                                            MsgBox "�ُ픭���I�I�@sts = " & sts
                                            Unload Me
                                    End Select
                                Loop
                            End If
                        End If
                
                        Exit Do
                
                    Case BtErrEOF
                        Exit Do
                    Case Else
                        MsgBox "�ُ픭���I�I�@sts = " & sts
                        Unload Me
                End Select
            
            Loop
            If sts = BtErrEOF Then
        
                Exit Do
            End If
        
            com = BtOpGetNext
        Loop
    
    End If
End Sub

