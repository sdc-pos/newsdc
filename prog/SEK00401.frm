VERSION 5.00
Begin VB.Form SEK00401 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�@�ʒ����f�[�^�J�z���� [SEK0040] 2011.06.29 14:00"
   ClientHeight    =   4704
   ClientLeft      =   1920
   ClientTop       =   2436
   ClientWidth     =   7860
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
   ScaleHeight     =   4704
   ScaleWidth      =   7860
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�@�ʒ����f�[�^"
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
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   3360
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�u�J�z�v�X�V���I"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   3840
   End
End
Attribute VB_Name = "SEK00401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'�u���Ұ�1 0:��ʊm�F����(�f�t�H���g)�@1;�L��v
'�u���Ұ�2 0:���ƍ��͎c��(�f�t�H���g)�@1:���ƍ��ł��폜�v
'�u���Ұ�3 0:�����i�͎c��(�f�t�H���g)�@1:�����i�ł��폜�v
'�u���Ұ�4 ����`(�f�t�H���g)�@YYYYMMDD:��`�����ꍇ���̓��ȑO�̃f�[�^�쐬�������폜�v

Private Option_Mode As Variant


Private Sub Form_Activate()
Dim ans As Integer

    
    
    If Option_Mode(0) = 1 Then
                        '�蓮���s
        Beep
        ans = MsgBox("�u�@�ʒ����f�[�^�J�z�����@���s���܂����H", vbYesNo + vbDefaultButton2, "�m�F����")
        If ans = vbYes Then
            
            
            Call Y_SYU_TEI_DEL_PROC
        End If
    
    Else
                        '�������s
        Call Y_SYU_TEI_DEL_PROC
    End If

    Unload Me



End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()

Dim c As String
Dim Today       As String * 8


    c = Command

    
    Today = Format(Now, "YYYYMMDD")
''    Today = "99999999"
    
    If Trim(c) = "" Then
        c = "0,0,0," & Today
    End If


    If Len(Trim(c)) = 1 Then
        c = Trim(c) & ",0,0," & Today
    End If

    If Len(Trim(c)) = 3 Then
        c = Trim(c) & ",0," & Today
    End If

    If Len(Trim(c)) = 5 Then
        c = Trim(c) & "," & Today
    End If


    Option_Mode = Split(Trim(c), ",", -1)


    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                


    Show
End Sub
Private Sub Form_Unload(CANCEL As Integer)
    
    Set SEK00401 = Nothing
        
    End
End Sub

Private Sub Y_SYU_TEI_DEL_PROC()

Dim sts         As Integer
Dim com         As Integer
        
Dim ans         As Integer
        
Dim Undo        As Boolean
Dim i           As Integer
        
        
        
        
    MsgLab(0).Visible = True
    MsgLab(1).Visible = True
        
    DoEvents
    If Y_SYU_TEI_Open(BtOpenNomal) Then                 '�o�ח\��f�[�^
        Exit Sub
    End If
        
    If DEL_SYU_TEI_Open(BtOpenNomal) Then               '�폜�Ϗo�ח\��f�[�^
        Exit Sub
    End If
        
    com = BtOpGetFirst

    Do
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Do
                    End If
                
                    com = BtOpGetEqual
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�@�ʒ����f�[�^")
                    Exit Do
            End Select
        Loop
    
        If sts = BtNoErr Then
            
            
            '�f�[�^�쐬���`�F�b�N
            If Option_Mode(3) >= StrConv(Y_SYU_TEI_REC.SND_YMD, vbUnicode) Then
                
                Undo = False
                
                '�ƍ��`�F�b�N
                If Option_Mode(1) = 0 Then
                    If Trim(StrConv(Y_SYU_TEI_REC.SHOGO_TANTO, vbUnicode)) = "" Then
                        Undo = True
                    End If
                End If
                '���i�`�F�b�N
                If Option_Mode(2) = 0 Then
                    If Trim(StrConv(Y_SYU_TEI_REC.KENPIN_TANTO, vbUnicode)) = "" Then
                        Undo = True
                    End If
                End If
            
            
                If StrConv(Y_SYU_TEI_REC.CANCEL_F, vbUnicode) = "1" Then
                    Undo = False
                End If
            
                If Undo Then
                Else
                    Do
                        DoEvents
                        sts = BTRV(BtOpInsert, DEL_SYU_TEI_POS, Y_SYU_TEI_REC, Len(DEL_SYU_TEI_REC), K0_DEL_SYU_TEI, Len(K0_DEL_SYU_TEI), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<DEL_SYUKA_TEI.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Exit Do
                                End If
                            Case Else
                                Call File_Error(sts, BtOpInsert, "�폜�ϓ@�ʒ����f�[�^")
                                Exit Do
                        End Select
                    Loop
                        
                    Do
                        DoEvents
                        sts = BTRV(BtOpDelete, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
                        Select Case sts
                            Case BtNoErr
                                Exit Do
                            Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                                Beep
                                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                If ans = vbCancel Then
                                    Exit Do
                                End If
                            Case Else
                                Call File_Error(sts, BtOpDelete, "�@�ʒ����f�[�^")
                                Exit Do
                        End Select
                    Loop
                    
                    
                End If
            End If
        End If
            
        If sts <> BtNoErr Then
            Exit Do
        End If
            
        com = BtOpGetNext
    
    Loop
        
                                                    '�@�ʒ����f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts Then
        Call File_Error(sts, BtOpClose, "�@�ʒ����f�[�^")
    End If
                                                    '�폜�Ϗo�ח\��f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, DEL_SYU_TEI_POS, DEL_SYU_TEI_REC, Len(DEL_SYU_TEI_REC), K0_DEL_SYU_TEI, Len(K0_DEL_SYU_TEI), 0)
    If sts Then
        Call File_Error(sts, BtOpClose, "�폜�ϓ@�ʒ����f�[�^")
    End If
        
    
    
    sts = BTRV(BtOpReset, Y_SYU_TEI_POS, Y_SYU_TEI_REC, Len(Y_SYU_TEI_REC), K0_Y_SYU_TEI, Len(K0_Y_SYU_TEI), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

End Sub

