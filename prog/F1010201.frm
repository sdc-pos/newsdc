VERSION 5.00
Begin VB.Form F1010201 
   BackColor       =   &H00C0C0C0&
   Caption         =   "���t�@�C���Z�b�g�A�b�v"
   ClientHeight    =   4710
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   7320
   ControlBox      =   0   'False
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
   ScaleHeight     =   4710
   ScaleWidth      =   7320
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label Lab_dsp 
      Height          =   252
      Index           =   7
      Left            =   4920
      TabIndex        =   15
      Top             =   3960
      Width           =   972
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   2  '��������
      Caption         =   "����"
      Height          =   252
      Index           =   5
      Left            =   2400
      TabIndex        =   14
      Top             =   3960
      Width           =   492
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   2  '��������
      Caption         =   "�ő�"
      Height          =   252
      Index           =   4
      Left            =   2400
      TabIndex        =   13
      Top             =   3480
      Width           =   492
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '��������
      Height          =   252
      Index           =   6
      Left            =   4320
      TabIndex        =   12
      Top             =   3960
      Width           =   252
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   2  '��������
      Caption         =   "-"
      Height          =   252
      Index           =   3
      Left            =   4080
      TabIndex        =   11
      Top             =   3960
      Width           =   132
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '��������
      Height          =   252
      Index           =   5
      Left            =   3720
      TabIndex        =   10
      Top             =   3960
      Width           =   252
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   2  '��������
      Caption         =   "-"
      Height          =   252
      Index           =   2
      Left            =   3480
      TabIndex        =   9
      Top             =   3960
      Width           =   132
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '��������
      Height          =   252
      Index           =   4
      Left            =   3120
      TabIndex        =   8
      Top             =   3960
      Width           =   252
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   2  '��������
      Caption         =   "-"
      Height          =   252
      Index           =   1
      Left            =   4080
      TabIndex        =   7
      Top             =   3480
      Width           =   132
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '��������
      Height          =   252
      Index           =   3
      Left            =   4320
      TabIndex        =   6
      Top             =   3480
      Width           =   252
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '��������
      Height          =   252
      Index           =   2
      Left            =   3720
      TabIndex        =   5
      Top             =   3480
      Width           =   252
   End
   Begin VB.Label Lab_Fix 
      Alignment       =   2  '��������
      Caption         =   "-"
      Height          =   252
      Index           =   0
      Left            =   3480
      TabIndex        =   4
      Top             =   3480
      Width           =   132
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '��������
      Height          =   252
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   3480
      Width           =   252
   End
   Begin VB.Label Lab_dsp 
      Alignment       =   2  '��������
      Height          =   252
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   2880
      Width           =   2532
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�I�f�[�^�X�V���I"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "���t�@�C���Z�b�g�A�b�v"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   5760
   End
End
Attribute VB_Name = "F1010201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
                                            '�I�}�X�^�̒ǉ��^����
Private Function Update_Proc() As Integer
Dim sts As Integer
Dim Upd_com As Integer
Dim com As Integer
Dim Retu, Ren, Dan As Integer
Dim ans As Integer
        
    MsgLab(0).Visible = True
    MsgLab(1).Visible = True
    Me.MousePointer = vbHourglass
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�q�Ƀ}�X�^")
                Update_Proc = True
        End Select
        Lab_dsp(0) = StrConv(SOKOREC.SOKO_NAME, vbUnicode)
        Lab_dsp(1) = StrConv(SOKOREC.RETU_END, vbUnicode)
        Lab_dsp(2) = StrConv(SOKOREC.REN_END, vbUnicode)
        Lab_dsp(3) = StrConv(SOKOREC.DAN_END, vbUnicode)
        
        
                                            '���z�͏������Ȃ�
'        If StrConv(SOKOREC.SOKO_KBN, vbUnicode) <> KBN_KASO$ Then
                                            '��̃��[�v
            For Retu = Val(StrConv(SOKOREC.RETU_START, vbUnicode)) To Val(StrConv(SOKOREC.RETU_END, vbUnicode))
                                            '�A�̃��[�v
                For Ren = Val(StrConv(SOKOREC.REN_START, vbUnicode)) To Val(StrConv(SOKOREC.REN_END, vbUnicode))
                                            '�i�̃��[�v
                    For Dan = Val(StrConv(SOKOREC.DAN_START, vbUnicode)) To Val(StrConv(SOKOREC.DAN_END, vbUnicode))
                                            '�ꉞ�ǉ����ł���l��
                            Lab_dsp(4) = Format$(Retu, "00")
                            Lab_dsp(5) = Format$(Ren, "00")
                            Lab_dsp(6) = Format$(Dan, "00")
                            DoEvents            '������Ƒ��v���Z�X�̗l�q������
                    
                            Call UniCode_Conv(K0_TANA.Soko_No, StrConv(SOKOREC.Soko_No, vbUnicode))
                            Call UniCode_Conv(K0_TANA.Retu, Format$(Retu, "00"))
                            Call UniCode_Conv(K0_TANA.Ren, Format$(Ren, "00"))
                            Call UniCode_Conv(K0_TANA.Dan, Format$(Dan, "00"))
                            Do
                                sts = BTRV(BtOpGetEqual + BtSNoWait, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Upd_com = BtOpUpdate
                                        Exit Do
                                    Case BtErrKeyNotFound
                                        Upd_com = BtOpInsert
                                        Exit Do
                                            '����͖���
                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                        Beep
                                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                        If ans = vbCancel Then
                                            Exit Function
                                        End If
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                        Update_Proc = True
                                        Exit Function
                                End Select
                            Loop
                                            '�I�f�[�^�X�V�^�ǉ�
'                           Call UniCode_Conv(TANAREC.JGYOBU, StrConv(SOKOREC.JGYOBU, vbUnicode))
                            Call UniCode_Conv(TANAREC.Soko_No, StrConv(SOKOREC.Soko_No, vbUnicode))
                            Call UniCode_Conv(TANAREC.Retu, Format$(Retu, "00"))
                            Call UniCode_Conv(TANAREC.Ren, Format$(Ren, "00"))
                            Call UniCode_Conv(TANAREC.Dan, Format$(Dan, "00"))
                            Call UniCode_Conv(TANAREC.KAHI_KBN, StrConv(SOKOREC.KAHI_KBN, vbUnicode))
 '                          Call UniCode_Conv(TANAREC.NAIGAI, StrConv(SOKOREC.NAIGAI, vbUnicode))
                            Call UniCode_Conv(TANAREC.TANA_COND, "0")
  '                         Call UniCode_Conv(TANAREC.KONS_KBN, StrConv(SOKOREC.KONS_KBN, vbUnicode))
                            
                            
                            Call UniCode_Conv(TANAREC.ZAIKO_SHOGO_FLG, ZAIKO_SHOGO_FLG_OK) '�݌ɏƍ��@�u�O�v   2004.02
                            
                            Call UniCode_Conv(TANAREC.FILLER, "")
                            If Upd_com = BtOpInsert Then
                                Lab_dsp(7) = "�ǉ�"
                            Else
                                Lab_dsp(7) = ""
                            End If
                            Do
                                sts = BTRV(Upd_com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
                                Select Case sts
                                    Case BtNoErr
                                        Exit Do
                                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                        Beep
                                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<TANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                        If ans = vbCancel Then
                                            Exit Function
                                        End If
                                    Case Else
                                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�I�}�X�^")
                                        Update_Proc = True
                                        Exit Function
                                End Select
                            Loop
                    Next Dan
                Next Ren
            Next Retu
'        End If
        
        com = BtOpGetNext
    
    Loop
End Function

Private Sub Form_Activate()

Dim ans As Integer
                                
                                '�����I��
    Beep
    ans = MsgBox("���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
    If ans = vbYes Then
        If Update_Proc() Then
            Unload Me
        End If
    End If
    MsgBox "�I�����܂����B"
    Unload Me

End Sub

Private Sub Form_DblClick()
    PrintForm
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
                                

                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�I�}�X�^�n�o�d�m
    If TANA_Open(BtOpenNomal) Then
        Unload Me
    End If
    
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
    
    sts = BTRV(BtOpReset, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1010201 = Nothing

    End
End Sub

