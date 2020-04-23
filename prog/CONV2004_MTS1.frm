VERSION 5.00
Begin VB.Form CONV2004_MTS1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "������Ǘ��}�X�^�Z�b�g�A�b�v����"
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
      Caption         =   "������b�r�u��"
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
      Caption         =   "������Ǘ��}�X�^�Z�b�g�A�b�v����"
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
      Width           =   7680
   End
End
Attribute VB_Name = "CONV2004_MTS1"
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

Dim In_NAIGAI               As String       '�����O
Dim In_DATA_KBN             As String       '�f�[�^�敪
Dim In_MUKE_CODE            As String       '���Ӑ�R�[�h
Dim In_SS_CODE              As String       '�q�Ɂ^�r�r�R�[�h
Dim In_MUKE_NAME            As String       '���Ӑ於��
Dim In_SS_NAME              As String       '�r�r����
Dim In_MUKE_DNAME           As String       '�\������
Dim In_DISPLAY_RANKING      As String       '�\������
Dim In_EOD                  As String

    Update_Proc = True
'---------------------------------------------  ������}�X�^�ǉ����ڃZ�b�g�A�b�v
    MsgLab(1) = "������}�X�^�Z�b�g�A�b�v�������I�I"
    Me.MousePointer = vbHourglass
                                                '������b�r�u�f�[�^�t���p�X�捞��
    sts = GetIni("FILE", "MTS_CSV", "SETUP", c)
    If sts <> False Then
        Call Log_Out(LOG_F, "SETUP.INI [MTS_CSV]�ǂݍ��݃G���[ ")
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
        
        
        Input #FileNo, In_NAIGAI, In_DATA_KBN, In_MUKE_CODE, In_SS_CODE, _
                        In_MUKE_NAME, In_SS_NAME, In_MUKE_DNAME, In_DISPLAY_RANKING, In_EOD
        On Error GoTo 0
        
        In_Count = In_Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            In_Cnt(0).Caption = Format(In_Count, "#0")
            DISP_INTERVAL = 0
        End If
        
        Call UniCode_Conv(MTSREC.NAIGAI, In_NAIGAI)
        Call UniCode_Conv(MTSREC.DATA_KBN, In_DATA_KBN)
        Call UniCode_Conv(MTSREC.MUKE_CODE, In_MUKE_CODE)
        Call UniCode_Conv(MTSREC.SS_CODE, In_SS_CODE)
        Call UniCode_Conv(MTSREC.MUKE_NAME, In_MUKE_NAME)
        Call UniCode_Conv(MTSREC.SS_NAME, In_SS_NAME)
        Call UniCode_Conv(MTSREC.MUKE_DNAME, In_MUKE_DNAME)
        
        If Not IsNumeric(In_DISPLAY_RANKING) Then
            Call UniCode_Conv(MTSREC.DISPLAY_RANKING, "999")
        Else
            Call UniCode_Conv(MTSREC.DISPLAY_RANKING, Format(CInt(In_DISPLAY_RANKING), "000"))
        End If
        
        
        Call UniCode_Conv(MTSREC.FILLER, "")
        
        Do
            sts = BTRV(BtOpInsert, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<MTS.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpUpdate, "������}�X�^")
                    Exit Function
            End Select
        Loop
        
            
    
    Loop

    In_Cnt(0).Caption = Format(In_Count, "#0")
    DoEvents

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
                                '������}�X�^�n�o�d�m
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '������}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "������}�X�^")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV2004_MTS1 = Nothing

    End
End Sub

