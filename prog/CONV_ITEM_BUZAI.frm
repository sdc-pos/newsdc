VERSION 5.00
Begin VB.Form CONV_ITEM_BUZAI1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "�i�ڃ}�X�^���ރf�[�^�쐬�����iCONV_ITEM_BUZAI 2012.03.23 17:00)"
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
      Caption         =   "�����O�폜"
      Height          =   435
      Index           =   2
      Left            =   7320
      TabIndex        =   12
      Top             =   2880
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I��"
      Height          =   435
      Index           =   1
      Left            =   7290
      TabIndex        =   11
      Top             =   1680
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�J�n"
      Height          =   435
      Index           =   0
      Left            =   7290
      TabIndex        =   10
      Top             =   1080
      Width           =   1860
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   4725
      TabIndex        =   9
      Top             =   1620
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   2250
      TabIndex        =   8
      Top             =   1620
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   5940
      TabIndex        =   7
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "��������"
      Height          =   315
      Index           =   2
      Left            =   5940
      TabIndex        =   6
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "�Ώی���"
      Height          =   315
      Index           =   1
      Left            =   4455
      TabIndex        =   5
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "�Ǎ��݌���"
      Height          =   315
      Index           =   0
      Left            =   2925
      TabIndex        =   4
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2925
      TabIndex        =   3
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "�i�ڃ}�X�^"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1350
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   4455
      TabIndex        =   1
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "�i�ڃ}�X�^���ރf�[�^�쐬����"
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
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   6720
   End
End
Attribute VB_Name = "CONV_ITEM_BUZAI1"
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

Dim wkKISHU1        As String * 25
Dim wkKISHU2        As String * 52
Dim wkKISHU3        As String * 150
Dim wkKISHU_BIKOU   As String * 450

Dim c               As String * 128

Dim i               As Integer

Dim CHG_FLG         As Boolean

Dim Start_Now       As String

    Update_Proc = True


'---------------------------------------------  ��������f�[�^�̃R���o�[�g
    Me.MousePointer = vbHourglass
    Count = 0
    sel_count = 0
    upd_count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
    Start_Now = Format(Now, "YYYYMMDDHHMMSS")
                                        
                                        
                                        
                                        
    Call UniCode_Conv(K0_ITEM.JGYOBU, "S")
    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")
                                        
                                        
                                        
    com = BtOpGetGreater
    Do
        
        DoEvents
        
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> "S" Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> "1" Then
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
        Cnt(1).Caption = Format(Count, "#0")
        Cnt(2).Caption = Format(Count, "#0")
        
        
        Call UniCode_Conv(ITEMREC.JGYOBU, "C")
        Call UniCode_Conv(ITEMREC.NAIGAI, "1")
        
        
        Do
            sts = BTRV(BtOpInsert, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrDuplicates
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                    Exit Function
            End Select
        Loop
        
        
        
        Call UniCode_Conv(K0_ITEM.JGYOBU, "S")
        Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        
        com = BtOpGetGreater
    
    Loop

    Cnt(0).Caption = Format(Count, "#0")


'---------------------------------------------  �I��
Update_End:
    
    Update_Proc = False

End Function
Private Function Delete_Proc() As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim Count           As Long
Dim sel_count       As Long
Dim upd_count       As Long

Dim DISP_INTERVAL   As Long

Dim wkKISHU1        As String * 25
Dim wkKISHU2        As String * 52
Dim wkKISHU3        As String * 150
Dim wkKISHU_BIKOU   As String * 450

Dim c               As String * 128

Dim i               As Integer

Dim CHG_FLG         As Boolean

Dim Start_Now       As String

    Delete_Proc = True


'---------------------------------------------  ��������f�[�^�̃R���o�[�g
    Me.MousePointer = vbHourglass
    Count = 0
    sel_count = 0
    upd_count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
                                        
    Start_Now = Format(Now, "YYYYMMDDHHMMSS")
                                        
                                        
                                        
                                        
    Call UniCode_Conv(K0_ITEM.JGYOBU, "S")
    Call UniCode_Conv(K0_ITEM.NAIGAI, "2")
    Call UniCode_Conv(K0_ITEM.HIN_GAI, "")
                                        
                                        
                                        
    com = BtOpGetGreater
    Do
        
        DoEvents
        
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> "S" Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> "2" Then
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
        Cnt(1).Caption = Format(Count, "#0")
        Cnt(2).Caption = Format(Count, "#0")
        
        
        Call UniCode_Conv(ITEMREC.JGYOBU, "C")
        Call UniCode_Conv(ITEMREC.NAIGAI, "1")
        
        
        Do
            sts = BTRV(BtOpDelete, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrDuplicates
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Exit Do
                Case Else
                    Call File_Error(sts, BtOpGetEqual + BtSNoWait, "�i�ڃ}�X�^")
                    Exit Function
            End Select
        Loop
        
        
        
'        Call UniCode_Conv(K0_ITEM.JGYOBU, "S")
'        Call UniCode_Conv(K0_ITEM.NAIGAI, "2")
'        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        
        com = BtOpGetNext
    
    Loop

    Cnt(0).Caption = Format(Count, "#0")


'---------------------------------------------  �I��
Update_End:
    
    Delete_Proc = False

End Function

Private Sub Command1_Click(Index As Integer)
    
Dim ans As Integer
    
    Select Case Index
        Case 0
            ans = MsgBox("���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            
                MsgBox "�I�����܂���"
                Unload Me
            
            End If


        Case 1
            Unload Me
    
    
        Case 2
    
    
            ans = MsgBox("���s���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Delete_Proc() Then
                    Unload Me
                End If
            
                MsgBox "�I�����܂���"
                Unload Me
            
            End If
    
    
    End Select

End Sub

Private Sub Form_Activate()

Dim ans As Integer
                                
                                
    Text1(0).Text = "20100716164000"
    Text1(1).Text = "20100802235900"
                                
                                

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyZ Then
        Text1(0).Visible = True
        Text1(1).Visible = True
    End If
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
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set CONV_ITEM_BUZAI1 = Nothing

    End
End Sub
