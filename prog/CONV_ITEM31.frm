VERSION 5.00
Begin VB.Form CONV_ITEM31 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�f�[�^���o�����iCONV_ITEM3 2010.08)"
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
   Begin VB.CommandButton Command1 
      Caption         =   "�I��"
      Height          =   435
      Index           =   1
      Left            =   7290
      TabIndex        =   14
      Top             =   1860
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���s"
      Height          =   435
      Index           =   0
      Left            =   7290
      TabIndex        =   11
      Top             =   1080
      Width           =   1860
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   4725
      TabIndex        =   10
      Top             =   1620
      Width           =   2265
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   2250
      TabIndex        =   9
      Top             =   1620
      Width           =   2265
   End
   Begin VB.Label Label4 
      Caption         =   "�w����t���ŕύX�̗L�����i�ڂ����O�o��"
      Height          =   375
      Left            =   135
      TabIndex        =   15
      Top             =   120
      Width           =   4740
   End
   Begin VB.Label Label3 
      Height          =   315
      Index           =   1
      Left            =   3735
      TabIndex        =   13
      Top             =   5880
      Width           =   2445
   End
   Begin VB.Label Label3 
      Height          =   315
      Index           =   0
      Left            =   1170
      TabIndex        =   12
      Top             =   5880
      Width           =   2445
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   5940
      TabIndex        =   8
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "���o����"
      Height          =   315
      Index           =   2
      Left            =   5940
      TabIndex        =   7
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "�Ώی���"
      Height          =   315
      Index           =   1
      Left            =   4455
      TabIndex        =   6
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "�Ǎ��݌���"
      Height          =   315
      Index           =   0
      Left            =   2925
      TabIndex        =   5
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2925
      TabIndex        =   4
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "�i�ڃ}�X�^"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1350
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   4455
      TabIndex        =   2
      Top             =   4800
      Width           =   1410
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
      Caption         =   "�f�[�^���o����"
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
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   3360
   End
End
Attribute VB_Name = "CONV_ITEM31"
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
Dim sel_count       As Long
Dim upd_count       As Long

Dim DISP_INTERVAL   As Long


Dim c               As String * 128

Dim i               As Integer


Dim Ins_Flg         As Boolean

Dim Start_Now       As String

Dim wk              As String








    Update_Proc = True


    Label3(0).Caption = Format(Now)
    Label3(1).Caption = ""


'---------------------------------------------  ��������f�[�^�̃R���o�[�g
    MsgLab(1) = "�i�ڃ}�X�^���o�������I�I"
    Me.MousePointer = vbHourglass
    Count = 0
    sel_count = 0
    upd_count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
    Cnt(1).Caption = Format(Count, "#0")
    Cnt(2).Caption = Format(Count, "#0")
                                        
    Start_Now = Format(Now, "YYYYMMDDHHMMSS")
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) = "S" Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                Exit Function
        End Select
        
        
        Count = Count + 1
        Cnt(0).Caption = Format(Count, "#0")
        
        
        If StrConv(ITEMREC.UPD_DATETIME, vbUnicode) >= Trim(Text1(0).Text) And _
            StrConv(ITEMREC.UPD_DATETIME, vbUnicode) <= Trim(Text1(1).Text) Then
        
        
            sel_count = sel_count + 1
            Cnt(1).Caption = Format(sel_count, "#0")
                
            '���o����
            
            Call LOG_OUT(Start_Now & "item.txt", "," & StrConv(ITEMREC.JGYOBU, vbUnicode) & "," & _
                        StrConv(ITEMREC.NAIGAI, vbUnicode) & "," & _
                        StrConv(ITEMREC.HIN_GAI, vbUnicode) & "," & _
                        "�X�V�S��=" & StrConv(ITEMREC.UPD_TANTO, vbUnicode) & "," & _
                        "�X�V����=" & StrConv(ITEMREC.UPD_DATETIME, vbUnicode))
                
        
        End If
        
        
        com = BtOpGetNext
    
    Loop

    Cnt(0).Caption = Format(Count, "#0")

    Label3(1).Caption = Format(Now)
    Me.MousePointer = vbDefault

'---------------------------------------------  �I��
Update_End:
    
    Update_Proc = False

End Function




Private Sub Command1_Click(Index As Integer)
    
Dim ans As Integer
    
Dim sts As Integer
    
Dim FullPath    As String
Dim c           As String * 128


    
    
    
    
    
    Select Case Index
        Case 0
            ans = MsgBox("���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                
                
                
                
                If Update_Proc() Then
                    Unload Me
                End If
            
            
            
            
                MsgBox "�I�����܂���"
            
            End If



        Case 1
            
            Unload Me
    
    End Select

End Sub

Private Sub Form_Activate()

Dim ans As Integer
                                
                                
    Text1(0).Text = "20100716164000"
    Text1(1).Text = "20101231235959"
                                
                                

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
    
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                    
                    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i��Ͻ�")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV_ITEM31 = Nothing

    End
End Sub

