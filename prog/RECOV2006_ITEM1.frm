VERSION 5.00
Begin VB.Form RECOV2006_ITEM1 
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
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2415
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1680
      Width           =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�I�@��"
      Height          =   495
      Left            =   7350
      TabIndex        =   8
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�J�@�n"
      Height          =   495
      Left            =   5355
      TabIndex        =   7
      Top             =   3960
      Width           =   1905
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   2415
      MaxLength       =   1
      TabIndex        =   5
      Top             =   2280
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "���ƕ���"
      Height          =   375
      Index           =   2
      Left            =   1260
      TabIndex        =   9
      Top             =   1680
      Width           =   1170
   End
   Begin VB.Label Label2 
      Caption         =   "(0:���i���v�@1:���i���s�v�j"
      Height          =   375
      Index           =   1
      Left            =   2625
      TabIndex        =   6
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "���i��F��"
      Height          =   375
      Index           =   0
      Left            =   1260
      TabIndex        =   4
      Top             =   2280
      Width           =   1170
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   3000
      TabIndex        =   3
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "�i�ڃ}�X�^��"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   3840
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
      Top             =   3000
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
Attribute VB_Name = "RECOV2006_ITEM1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Update_Proc() As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim com_Zaiko       As Integer
Dim ans             As Integer
Dim Count           As Long

Dim DISP_INTERVAL   As Long


Dim MTS_CODE        As String * 8
Dim SS_CODE         As String * 8

Dim c               As String * 128

    Update_Proc = True

'---------------------------------------------  �i�ڃ}�X�^�̃R���o�[�g
    MsgLab(1) = "�i�ڃ}�X�^�R���o�[�g�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
    Call UniCode_Conv(K0_ITEM.JGYOBU, Text1(1).Text)
    Call UniCode_Conv(K0_ITEM.NAIGAI, "")
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")
                                        
                                        
                                        
    com = BtOpGetGreater
    Do
        
        DoEvents
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            
            
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Trim(Text1(1).Text) Then
                    Exit Do
                End If
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                Exit Function
        End Select
        
        
        Count = Count + 1
        DISP_INTERVAL = DISP_INTERVAL + 1
        If DISP_INTERVAL = 100 Then
            Cnt(0).Caption = Format(Count, "#0")
            DISP_INTERVAL = 0
        End If

'        If StrConv(ITEMREC.ZAIKO_F, vbUnicode) = "0" Then
'            Call UniCode_Conv(ITEMREC.ZAIKO_F, "1")
'        Else
        Call UniCode_Conv(ITEMREC.GOODS_KBN, Text1(0).Text)
'        End If
        
        Do
            sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<ITEM.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                    Exit Function
            End Select
        Loop
        
        com = BtOpGetNext
    
    Loop

    Cnt(0).Caption = Format(Count, "#0")
'---------------------------------------------  �I��
Update_End:
    
    Update_Proc = False

End Function

Private Sub Command1_Click()
Dim ans As Integer
                                
                                
                                
    If Text1(0).Text <> "0" And Text1(0).Text <> "1" Then
        MsgBox "���i��F�G���["
        Text1(0).SetFocus
        Exit Sub
    End If
                                
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

Private Sub Command2_Click()
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
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�i�ڃ}�X�^�b�k�n�r�d
   sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
   If sts Then
       If sts <> BtErrNoOpen Then
           Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
       End If
   End If
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set RECOV2006_ITEM1 = Nothing

    End
End Sub
