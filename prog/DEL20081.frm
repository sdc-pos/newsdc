VERSION 5.00
Begin VB.Form DEL20081 
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
   Begin VB.Label Label1 
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   4410
      TabIndex        =   5
      Top             =   3000
      Width           =   510
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   4935
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2940
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�\���}�X�^��"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   630
      TabIndex        =   2
      Top             =   3000
      Width           =   2295
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
Attribute VB_Name = "DEL20081"
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
Dim DEL_Count       As Long

Dim DISP_INTERVAL   As Long


Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8

Dim c               As String * 128

    Update_Proc = True

'---------------------------------------------  �\���}�X�^�̍폜
    MsgLab(1) = "�\���}�X�^�폜�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    
    Count = 0
    DEL_Count = 0
    
    Cnt(0).Caption = Format(Count, "#0")
    Cnt(1).Caption = Format(DEL_Count, "#0")
                                        
                                        
    Call UniCode_Conv(K0_P_COMPO.JGYOBU, "B")
    Call UniCode_Conv(K0_P_COMPO.NAIGAI, "")
    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, "")
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        sts = BTRV(com, P_COMPO_POS, P_COMPO_K_REC, Len(P_COMPO_K_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
        Select Case sts
            Case BtNoErr
            
                If StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) <> "B" Then
                    Exit Do
                
                End If
                
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�\���}�X�^")
                Exit Function
        End Select
        
        
        
        If StrConv(P_COMPO_K_REC.SEQNO, vbUnicode) = "000" Then
        Else
            
            
            If StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = "1" Or _
                StrConv(P_COMPO_K_REC.DATA_KBN, vbUnicode) = "2" Then
            
                Count = Count + 1
                DISP_INTERVAL = DISP_INTERVAL + 1
                Cnt(0).Caption = Format(Count, "#0")
        
            
                Call UniCode_Conv(K0_ITEM.JGYOBU, "S")
                Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                            
            
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                    Case BtErrKeyNotFound
                    
                    
                    
                        Call Log_Out(LOG_F, StrConv(P_COMPO_K_REC.JGYOBU, vbUnicode) & "�|" & _
                                            StrConv(P_COMPO_K_REC.NAIGAI, vbUnicode) & "�|" & _
                                            StrConv(P_COMPO_K_REC.HIN_GAI, vbUnicode) & "�|" & _
                                            StrConv(P_COMPO_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    
                                            
                        Do
                        
                            sts = BTRV(BtOpDelete, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                            Select Case sts
                                Case BtNoErr
                                    Exit Do
                                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                    Beep
                                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_COMPO.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                                    If ans = vbCancel Then
                                        Exit Function
                                    End If
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                                    Exit Function
                            End Select
                        
                        
                        Loop
                    
                        DEL_Count = DEL_Count + 1
                        Cnt(1).Caption = Format(DEL_Count, "#0")
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        Exit Function
                End Select
            
            
            
            End If
        
        End If
        
        com = BtOpGetNext
    
    Loop



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
                                '�\���}�X�^�n�o�d�m
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�\���}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    
                                            '�\���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�\���}�X�^")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set DEL20081 = Nothing

    End
End Sub

