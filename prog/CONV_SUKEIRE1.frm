VERSION 5.00
Begin VB.Form CONV_SUKEIRE1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�f�[�^�R���o�[�g����"
   ClientHeight    =   7230
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   10095
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
   ScaleHeight     =   7230
   ScaleWidth      =   10095
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   4515
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�������"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   3
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
      Width           =   240
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�f�[�^�R���o�[�g����"
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
      Width           =   4800
   End
End
Attribute VB_Name = "CONV_SUKEIRE1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Update_Proc() As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim Count           As Long

Dim upd_count       As Long

Dim DISP_INTERVAL   As Long



Dim c               As String * 128

    Update_Proc = True


'---------------------------------------------  ��������f�[�^�̃R���o�[�g
    MsgLab(1) = "��������f�[�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    upd_count = 0
    DISP_INTERVAL = 0
    Cnt(2).Caption = Format(Count, "#0")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        
        sts = BTRV(com, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�������")
                Exit Function
        End Select
        
        
        Count = Count + 1
        Cnt(2).Caption = Format(Count, "#0")
        
        
        
        If Left(StrConv(P_SUKEIRE_REC.UPD_DATETIME, vbUnicode), 8) = "20080820" Then
        
        
            upd_count = upd_count + 1
            Cnt(0).Caption = Format(upd_count, "#0")
        
        
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(9).NIN, StrConv(P_SUKEIRE_REC.GENKA_TBL(8).NIN, vbUnicode))
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(9).TIMES, StrConv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, vbUnicode))
        
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(8).NIN, StrConv(P_SUKEIRE_REC.GENKA_TBL(7).NIN, vbUnicode))
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(8).TIMES, StrConv(P_SUKEIRE_REC.GENKA_TBL(7).TIMES, vbUnicode))
        
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(7).NIN, StrConv(P_SUKEIRE_REC.GENKA_TBL(6).NIN, vbUnicode))
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(7).TIMES, StrConv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, vbUnicode))
        
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(6).NIN, StrConv(P_SUKEIRE_REC.GENKA_TBL(5).NIN, vbUnicode))
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(6).TIMES, StrConv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, vbUnicode))
        
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(5).NIN, StrConv(P_SUKEIRE_REC.GENKA_TBL(4).NIN, vbUnicode))
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(5).TIMES, StrConv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, vbUnicode))
        
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).NIN, StrConv(P_SUKEIRE_REC.GENKA_TBL(3).NIN, vbUnicode))
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(4).TIMES, StrConv(P_SUKEIRE_REC.GENKA_TBL(3).TIMES, vbUnicode))
        
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(3).NIN, "")
            Call UniCode_Conv(P_SUKEIRE_REC.GENKA_TBL(3).TIMES, "")
        
        
        
            Do
                sts = BTRV(BtOpUpdate, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_SUKEIRE.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�������")
                        Exit Function
                End Select
            Loop
        
        
        End If
        
        
        com = BtOpGetNext
    
    Loop

    Cnt(2).Caption = Format(Count, "#0")


'---------------------------------------------  �I��
Update_End:
    
    Update_Proc = False

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
    
    If P_SUKEIRE_Open(BtOpenNomal) Then
        Unload Me
    End If
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    sts = BTRV(BtOpClose, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�������")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, P_SUKEIRE_POS, P_SUKEIRE_REC, Len(P_SUKEIRE_REC), K0_P_SUKEIRE, Len(K0_P_SUKEIRE), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV_SUKEIRE1 = Nothing

    End
End Sub

