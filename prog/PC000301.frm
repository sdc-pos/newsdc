VERSION 5.00
Begin VB.Form PC000301 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�N���X�}�X�^�R���o�[�g����"
   ClientHeight    =   7230
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   9120
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
   ScaleWidth      =   9120
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label Cnt 
      Alignment       =   1  '�E����
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�N���X�}�X�^��"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   3000
      Width           =   1695
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
Attribute VB_Name = "PC000301"
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

Dim DISP_INTERVAL   As Long


Dim FileNo          As Long
Dim fileName        As String


Dim CLASS_REC     As Variant
Dim RecordBuf       As String

Dim c               As String * 128

    Update_Proc = True

    FileNo = FreeFile
    
                                '���O�t�@�C������荞��
    If GetIni("FILE", "CLASS_TXT", "CONV2006", c) Then
        Beep
        MsgBox "[CLASS_TXT]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
    fileName = RTrim(c)
    
        
    Open fileName For Input As FileNo
    
    
    
    
    
    MsgLab(1) = "�N���X�}�X�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
                                        
                                        
                                        
    Do Until EOF(FileNo)
        
        DoEvents
        
        Line Input #FileNo, RecordBuf
        
        CLASS_REC = Split(RecordBuf, vbTab, -1)
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(P_CLASSREC.SHIMUKE_CODE, "01")                                 '�d�����溰�ށi�|���@�͌Œ�j
        Call UniCode_Conv(P_CLASSREC.CLASS_CODE, CStr(CLASS_REC(0)))                '�N���X
        Call UniCode_Conv(P_CLASSREC.CLASS_NAME, CStr(CLASS_REC(3)))                '�Ăі�
                                                                                    '���i�����i
        
        If IsNumeric(CLASS_REC(1)) Then
            Call UniCode_Conv(P_CLASSREC.TANKA, Format(CDbl(CLASS_REC(1)), "00000000.00"))
        Else
            Call UniCode_Conv(P_CLASSREC.TANKA, "00000000.00")
        End If
                                                                                    
        If IsNumeric(CLASS_REC(9)) Then                                             '�H��
            Call UniCode_Conv(P_CLASSREC.KOUSU, Format(CDbl(CLASS_REC(9)), "000.000"))
        Else
            Call UniCode_Conv(P_CLASSREC.KOUSU, "000.000")
        End If
                                                                                    '�H��
        If IsNumeric(CLASS_REC(10)) Then
            Call UniCode_Conv(P_CLASSREC.KOURYOU, Format(CDbl(CLASS_REC(10)), "00000000.00"))
        Else
            Call UniCode_Conv(P_CLASSREC.KOURYOU, "00000000.00")
        End If
                                            
        Call UniCode_Conv(P_CLASSREC.ETC, "00000000.00")                            '���̑�
        
        
        
        Call UniCode_Conv(P_CLASSREC.UPD_TANTO, "CONV")                             '�X�V�S����
                                                                                    '�X�V����
        Call UniCode_Conv(P_CLASSREC.UPD_DATETIME, Format(Now, "YYYYMMDD") & Format(Now, "HHMMSS"))
        
        
        Do
            sts = BTRV(BtOpInsert, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_CLASS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case BtErrDuplicates
                    
                    Exit Do
                
                Case Else
                    
                    
                    
                    Call File_Error(sts, BtOpInsert, "�N���X�}�X�^")
                    Exit Function
            End Select
        Loop
        
    
    Loop
'---------------------------------------------  �I��

    Cnt(0).Caption = Format(Count, "#0")
    
    Close #FileNo

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
                                '�N���X�}�X�^�n�o�d�m
    If P_Class_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�N���X�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�N���X�}�X�^")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, P_CLASS_POS, P_CLASSREC, Len(P_CLASSREC), K0_P_CLASS, Len(K0_P_CLASS), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set PC000301 = Nothing

    End
End Sub

