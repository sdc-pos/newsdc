VERSION 5.00
Begin VB.Form CONV2004_YOIN1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�v���}�X�^�Z�b�g�A�b�v����"
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
   Begin VB.Label In_Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�v���b�r�u��"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1200
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
      Caption         =   "�v���}�X�^�Z�b�g�A�b�v����"
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
      Width           =   6240
   End
End
Attribute VB_Name = "CONV2004_YOIN1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Update_Proc() As Integer

Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

Dim In_Count        As Long
Dim Out_Count       As Long

Dim DISP_INTERVAL   As Long

Dim fileName        As String
Dim FileNo          As Integer

Dim c               As String * 128


Dim In_CODE_TYPE            As String       '��o�[�R�[�h�^�C�v
Dim In_YOIN_CODE            As String       '�v��
Dim In_YOIN_DNAME           As String       '��ƕ\������
Dim In_SUM_KBN              As String       '�W�v�敪
Dim In_SYSTEM_F             As String       '�V�X�e���\���׸�
Dim In_REGI_F               As String       '�o�^���׸�
Dim In_PARAM_F              As String       '�t�����Ұ�
Dim In_Soko_No              As String       '�q�ɇ��i���z�j
Dim In_EOD                  As String

    Update_Proc = True
'---------------------------------------------  �v���}�X�^�ǉ����ڃZ�b�g�A�b�v
    MsgLab(1) = "�v���}�X�^�Z�b�g�A�b�v�������I�I"
    Me.MousePointer = vbHourglass
                                                '������b�r�u�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", "YOIN_CSV", "SETUP", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SETUP.INI [YOIN_CSV]�ǂݍ��݃G���[ ")
        Exit Function
    End If
    fileName = Trim(c)
    
        
    
    On Error GoTo Error_Proc
        
    FileNo = FreeFile
    Open fileName For Input As #FileNo
    
    On Error GoTo 0
    
    
    
    In_Count = 0
    DISP_INTERVAL = 0
    In_Cnt(0).Caption = Format(In_Count, "#0")
                                        
                                        
    Do
        
        DoEvents
            
        On Error GoTo Error_Proc
        
        
        
        
        Input #FileNo, In_CODE_TYPE, In_YOIN_CODE, In_YOIN_DNAME, In_SUM_KBN, _
                        In_SYSTEM_F, In_REGI_F, In_PARAM_F, In_Soko_No, In_EOD
        On Error GoTo 0
        
        In_Count = In_Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            In_Cnt(0).Caption = Format(In_Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(YOINREC.CODE_TYPE, In_CODE_TYPE)
        Call UniCode_Conv(YOINREC.YOIN_CODE, In_YOIN_CODE)
        Call UniCode_Conv(YOINREC.YOIN_DNAME, In_YOIN_DNAME)
        Call UniCode_Conv(YOINREC.SUM_KBN, In_SUM_KBN)
        Call UniCode_Conv(YOINREC.SYSTEM_F, In_SYSTEM_F)
        Call UniCode_Conv(YOINREC.REGI_F, In_REGI_F)
        Call UniCode_Conv(YOINREC.PARAM_F, In_PARAM_F)
        Call UniCode_Conv(YOINREC.Soko_No, In_Soko_No)
        
        
        Call UniCode_Conv(YOINREC.FILLER, "")
        
        Do
            sts = BTRV(BtOpInsert, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<YOIN.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpInsert, "�v���}�X�^")
                    Exit Function
            End Select
        Loop
        
            
    
    Loop

    In_Cnt(0).Caption = Format(In_Count, "#0")

    MsgBox "����I�����܂���"
'---------------------------------------------  �I��
    Update_Proc = False
    
    Exit Function

Error_Proc:
Const ErrDiskNotReady = 71, ErrDeviceUnavailable = 68, ErrNotFound = 53
    Select Case Err.Number
        Case 62
            In_Cnt(0).Caption = Format(In_Count, "#0")
            MsgBox "����I�����܂���"
            Update_Proc = False
            Exit Function
        Case ErrDiskNotReady
            Beep
            ans = MsgBox("�h���C�u���m�F���ĉ�����", vbYesNo + vbExclamation + vbDefaultButton1, "�m�F����")
            If ans = vbYes Then
                Resume
            End If
        Case ErrDeviceUnavailable
            Beep
            ans = MsgBox("�h���C�u��������܂���" & fileName, vbExclamation)
        Case ErrNotFound
            Beep
            ans = MsgBox("�t�@�C����������܂���" & fileName, vbExclamation)
        Case 76
            Beep
            ans = MsgBox("�t�@�C���p�X��������܂���" & fileName, vbExclamation)
        Case Else
            Beep
            ans = MsgBox("�G���[ [PACKING_CSV Open : " & Str(Err.Number) & "] " & Err.Description, vbCritical)
    End Select

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
                                '�v���}�X�^�n�o�d�m
    If YOIN_Open(BtOpenNomal) Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�v���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�v���}�X�^")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, YOIN_POS, YOINREC, Len(YOINREC), K0_YOIN, Len(K0_YOIN), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV2004_YOIN1 = Nothing

    End
End Sub

