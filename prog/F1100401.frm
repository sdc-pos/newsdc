VERSION 5.00
Begin VB.Form F1100401 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�V�X�e���N������"
   ClientHeight    =   4710
   ClientLeft      =   1920
   ClientTop       =   2430
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
   MousePointer    =   11  '�����v
   ScaleHeight     =   4710
   ScaleWidth      =   7320
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   19.5
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   396
      Left            =   3000
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   1236
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "�V�X�e���N���������ł��B                ���΂炭���҂��������B"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   22.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1092
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   7452
   End
End
Attribute VB_Name = "F1100401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WS_NO       As String * 2           '���[���ԍ�
Dim SERVER_ID   As String * 2           '�T�[�o�[�h�c

 
Private Sub Form_Activate()
Dim sts  As Integer
    
    
    If WS_NO = SERVER_ID Then
'---------------------------'�T�[�o�[��̏���
        MsgBox "���Ӌ@��̓d����Ԃ��m�F��A�u�d���������v�L�[�������ĉ������B", vbSystemModal
                            '�o�ח\��̊J��
'''        sts = Y_SYUKA_UNLOCK_PROC()
'''        If sts Then
'''            End
'''        End If
                            '�݌ɂ̊J��
'''        sts = Zaiko_UNLOCK_Proc()
'''        If sts Then
'''            End
'''        End If

        DoEvents

        sts = BTRV(BtOpReset, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        If sts Then
            Call File_Error(sts, BtOpReset, "")
        End If

        
        sts = Shell("..\exe\f110010.exe", vbNormalFocus)
        If sts = 0 Then
            MsgBox "[F110010]�X�L���i����̋N���Ɏ��s���܂���� "
            Call Log_Out(LOG_F, "[F110010]�X�L���i����̋N���Ɏ��s���܂����")
        End If
        
'        sts = Shell("..\exe\f120050.exe", vbNormalFocus)
'        If sts = 0 Then
'            MsgBox "[F120050]�����Ϗo�א��Z�o�����̋N���Ɏ��s���܂���� "
'            Call Log_Out(LOG_F, "[F120050]�����Ϗo�א��Z�o�����̋N���Ɏ��s���܂����")
'            End
'        End If

    Else
'---------------------------'�N���C�A���g��̏���
        MsgBox "�T�[�o�[�o�b�̗������m�F��A�u�d���������v�L�[�������ĉ������B", vbSystemModal
        Call FILE_BACKUP_PROC
    End If

    Unload Me
End Sub


Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()

Dim c           As String * 128
Dim sts         As Integer
    
Dim sBuffer     As String * 255
Dim com         As String
    
    
    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
  
                                
    Label1.Visible = True
'���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
    
'���[���ԍ���荞��
    sBuffer = Space(255)
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        com = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        com = "??"
    End If
    WS_NO = StrConv(RTrim(com), vbUpperCase)

'�T�[�o�[�h�c��荞��
    If GetIni("SYSTEM", "SERVER_ID", "SYS", c) Then
        Beep
        MsgBox "�T�[�o�[�h�c�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Call Log_Out(LOG_F, "[SYS.INI] [SYSTEM] [SERVER_ID] READ ERROR")
        End
    End If
    SERVER_ID = StrConv(RTrim(c), vbUpperCase)
    
    
    
    End Sub

Private Sub Form_Unload(CANCEL As Integer)
    
    Set F1100401 = Nothing

    End
End Sub


Public Function Y_SYUKA_UNLOCK_PROC() As Integer
        
Dim sts As Integer
Dim com As Integer
       
Dim ans As Integer
       
    Y_SYUKA_UNLOCK_PROC = False
                                
    If Y_SYU_Open(BtOpenNomal) Then                 '�o�ח\��f�[�^
        Exit Function
    End If
        
        
    Call UniCode_Conv(K4_Y_SYU.WEL_ID, "")
    Call UniCode_Conv(K4_Y_SYU.PRG_ID, "")
        
    com = BtOpGetGreater

    Do
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'���̎��_�ł̃t�@�C���g�p���͖������[�v�Ƃ���B�L�����Z���ňُ�I��
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Y_SYUKA_UNLOCK_PROC = True
                        Exit Do
                    End If
                
                    com = BtOpGetEqual
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�o�ח\��f�[�^")
                    Y_SYUKA_UNLOCK_PROC = True
                    Exit Do
            End Select
        Loop
    
        If sts = BtNoErr Then
            
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")          '�g�p�q�@ID
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")          '�g�p��۸���
    
            Do
                DoEvents
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), BtNCC)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Y_SYUKA_UNLOCK_PROC = True
                            Exit Do
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "�o�ח\��")
                        Y_SYUKA_UNLOCK_PROC = True
                        Exit Do
                End Select
            Loop
                
        End If
            
        If sts <> BtNoErr Then
            Exit Do
        End If
            
        com = BtOpGetNext
    
    Loop
    
                                                '�o�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpClose, "�o�ח\��f�[�^")
        Y_SYUKA_UNLOCK_PROC = True
        Exit Function
    End If



End Function

Private Function Zaiko_UNLOCK_Proc() As Integer
        
Dim sts As Integer
Dim com As Integer
       
Dim ans As Integer
       
    Zaiko_UNLOCK_Proc = False
                                
    If ZAIKO_Open(BtOpenNomal) Then                 '�݌Ƀf�[�^
        Exit Function
    End If
        
        
    Call UniCode_Conv(K3_ZAIKO.WEL_ID, "")
    Call UniCode_Conv(K3_ZAIKO.PRG_ID, "")
        
    com = BtOpGetGreater

    Do
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), 3)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
'���̎��_�ł̃t�@�C���g�p���͖������[�v�Ƃ���B�L�����Z���ňُ�I��
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Zaiko_UNLOCK_Proc = True
                        Exit Do
                    End If
                
                    com = BtOpGetEqual
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�݌Ƀf�[�^")
                    Zaiko_UNLOCK_Proc = True
                    Exit Do
            End Select
        Loop
    
        If sts = BtNoErr Then
            
            Call UniCode_Conv(ZAIKOREC.LOCK_F, LOCK_OFF)    '�r���׸ށiOFF�j
            Call UniCode_Conv(ZAIKOREC.WEL_ID, "")          '�g�p�q�@ID
            Call UniCode_Conv(ZAIKOREC.PRG_ID, "")          '�g�p��۸���
    
            Do
                DoEvents
                sts = BTRV(BtOpUpdate, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K3_ZAIKO, Len(K3_ZAIKO), BtNCC)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ZAIKO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Zaiko_UNLOCK_Proc = True
                            Exit Do
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "�݌Ƀf�[�^")
                        Zaiko_UNLOCK_Proc = True
                        Exit Do
                End Select
            Loop
                
        End If
            
        If sts <> BtNoErr Then
            Exit Do
        End If
            
        com = BtOpGetNext
    
    Loop
    
                                                '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        Zaiko_UNLOCK_Proc = True
        Exit Function
    End If


End Function

Public Sub FILE_BACKUP_PROC()

Dim FROM_DIR    As String
Dim TO_DIR      As String
Dim FILE_NAME   As String
Dim c           As String * 128
                                    '�o�b�N�A�b�v���t�H���_��荞��
    If GetIni("FILE", "BACK_FROM", "SYS", c) Then
        Beep
        MsgBox "�o�b�N�A�b�v���t�H���_�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Exit Sub
    End If
    FROM_DIR = RTrim(c)

    If GetIni("FILE", "BACK_TO", "SYS", c) Then
        Beep
        MsgBox "�o�b�N�A�b�v��t�H���_�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        Exit Sub
    End If
    TO_DIR = RTrim(c)

    Label2.Visible = True
    
    On Error GoTo Err_Proc
    ChDir FROM_DIR

    FILE_NAME = Dir(FROM_DIR, vbNormal)

    Do While FILE_NAME <> ""
        DoEvents
        Label2.Caption = "�u" & FILE_NAME & "�v�o�b�N�A�b�v���ł��B"
        On Error Resume Next
        FileCopy FROM_DIR & FILE_NAME, TO_DIR & FILE_NAME
        FILE_NAME = Dir
    Loop
    Exit Sub

Err_Proc:
    If Err.Number = 76 Then
        MsgBox "�l�b�g���[�N�ւ̐ڑ����s���ł��B�ė��グ���s���Ă��������B"
        Exit Sub
    Else
        Resume Next
    End If
End Sub
