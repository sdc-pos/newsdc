VERSION 5.00
Begin VB.Form F1011301 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���Y��Ͻ�����ݽ"
   ClientHeight    =   8670
   ClientLeft      =   2130
   ClientTop       =   2730
   ClientWidth     =   12030
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
   MaxButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   12030
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   420
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   9360
      MaxLength       =   14
      TabIndex        =   4
      Top             =   1320
      Width           =   1950
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Appearance      =   0  '�ׯ�
      Height          =   420
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   8880
      TabIndex        =   24
      Top             =   6720
      Width           =   1140
   End
   Begin VB.ListBox List2 
      Height          =   2700
      Left            =   9000
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   3120
      Visible         =   0   'False
      Width           =   2310
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   1920
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   420
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   6240
      MaxLength       =   14
      TabIndex        =   3
      Top             =   1320
      Width           =   1950
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   420
      Index           =   1
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1320
      Width           =   3120
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   420
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   1800
      MaxLength       =   20
      TabIndex        =   1
      Top             =   720
      Width           =   3120
   End
   Begin VB.ListBox List1 
      Height          =   3900
      ItemData        =   "F1011301.frx":0000
      Left            =   960
      List            =   "F1011301.frx":0002
      TabIndex        =   5
      Top             =   2520
      Width           =   8835
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I  ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10440
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9600
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8760
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7920
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6600
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5760
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2760
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�X �V"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�X�V����"
      Height          =   255
      Index           =   8
      Left            =   7800
      TabIndex        =   31
      Top             =   2280
      Width           =   1155
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�^����"
      Height          =   255
      Index           =   7
      Left            =   5880
      TabIndex        =   30
      Top             =   2280
      Width           =   1155
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���Y��"
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   29
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�@��"
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   28
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�X�V����"
      Height          =   255
      Index           =   0
      Left            =   8280
      TabIndex        =   27
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Label LabJIGYO 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   240
      Left            =   240
      TabIndex        =   26
      Top             =   7920
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�\������"
      Height          =   255
      Index           =   4
      Left            =   7800
      TabIndex        =   25
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   255
      Index           =   33
      Left            =   960
      TabIndex        =   22
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�^����"
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   21
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���Y��"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   20
      Top             =   1440
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   19
      Top             =   720
      Width           =   5025
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�@��"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   18
      Top             =   840
      Width           =   855
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1011301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const ptxHIN_GAI% = 0
Private Const ptxGENSANKOKU% = 1
Private Const ptxINS_DateTime% = 2

Private Const ptxUPD_DateTime% = 3              '2013.02.19

Private Const ptxDISP_COUNT% = 4


Private Const Text_Max% = 4

Private Const pcmbNAIGAI% = 0

'Private Const LAST_UPDATE_DAY$ = "(F101130 2018.04.08 14:45)"
Private Const LAST_UPDATE_DAY$ = "(F101130 2018.04.12 13:15)"


Private Function List_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���X�g�{�b�N�X�\������
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer

Dim svHin_Gai   As String
Dim In_Cnt      As Long

Dim wkString    As String


    List_Proc = True
    
    List1.Clear
    List2.Clear
    
    If Not IsNumeric(Text1(ptxDISP_COUNT).Text) Then
        Text1(ptxDISP_COUNT).Text = 100
    End If
    In_Cnt = 0
    
    Call UniCode_Conv(K0_GENSAN.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_GENSAN.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_GENSAN.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_GENSAN.GENSANKOKU, "")
    
    com = BtOpGetGreaterEqual
    svHin_Gai = ""
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�����\���J�n" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    
    
    Do
        DoEvents
        
        sts = BTRV(com, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(GENSANREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(GENSANREC.NAIGAI, vbUnicode) <> Right(Combo1(pcmbNAIGAI).Text, 1) Then
                    Exit Do
                End If
                
                If Trim(Text1(ptxHIN_GAI).Text) <> "" Then
                    If Trim(Text1(ptxHIN_GAI).Text) <> Mid(StrConv(GENSANREC.HIN_GAI, vbUnicode), 1, Len(Trim(Text1(ptxHIN_GAI).Text))) Then
                        Exit Do
                    End If
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���Y���}�X�^")
                Exit Function
        End Select
    
    
        In_Cnt = In_Cnt + 1
        If In_Cnt > Val(Text1(ptxDISP_COUNT).Text) Then
            Exit Do
        End If
        
        If Trim(svHin_Gai) = "" Then
            svHin_Gai = Trim(StrConv(GENSANREC.HIN_GAI, vbUnicode))
        End If
        
        
        If svHin_Gai <> Trim(StrConv(GENSANREC.HIN_GAI, vbUnicode)) Then
            If List2.ListCount <= 0 Then
            Else
                For i = List2.ListCount - 1 To 0 Step -1
                
                    List1.AddItem Mid(List2.List(i), 1, 20) & Mid(List2.List(i), 35, 20) & Mid(List2.List(i), 21, 8) & "-" & Mid(List2.List(i), 29, 6) & " " & Mid(List2.List(i), 55, 8) & "-" & Mid(List2.List(i), 63, 6)
                 
                 
                 
                Next i
            End If
            List2.Clear
        End If
                        
        wkString = ""
        For i = 1 To Len(StrConv(GENSANREC.Ins_DateTime, vbUnicode))
        
        
            If Mid(StrConv(GENSANREC.Ins_DateTime, vbUnicode), i, 1) < " " Then
                
                
                wkString = wkString & " "
                
            Else
                wkString = wkString & Mid(StrConv(GENSANREC.Ins_DateTime, vbUnicode), i, 1)
            End If
        
        
        Next i
        
        
        
'        List2.AddItem StrConv(GENSANREC.HIN_GAI, vbUnicode) & StrConv(GENSANREC.Ins_DateTime, vbUnicode) & StrConv(GENSANREC.GENSANKOKU, vbUnicode) & StrConv(GENSANREC.UPD_DATETIME, vbUnicode)
        List2.AddItem StrConv(GENSANREC.HIN_GAI, vbUnicode) & wkString & StrConv(GENSANREC.GENSANKOKU, vbUnicode) & StrConv(GENSANREC.UPD_DATETIME, vbUnicode)
        
        com = BtOpGetNext
    
    
    Loop
    
    
    
    If List2.ListCount <= 0 Then
    Else
        For i = List2.ListCount - 1 To 0 Step -1
            List1.AddItem Mid(List2.List(i), 1, 20) & Mid(List2.List(i), 35, 20) & Mid(List2.List(i), 21, 8) & "-" & Mid(List2.List(i), 29, 6) & " " & Mid(List2.List(i), 55, 8) & "-" & Mid(List2.List(i), 63, 6)
         
        Next i
    End If
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "�����\���I��" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    List_Proc = False
    
End Function
Private Sub Clear_Field(Mode As Integer)
'----------------------------------------------------------------------------
'                   ��ʏ�������
'----------------------------------------------------------------------------
Dim i As Integer

'    For i = 0 To ptxINS_DateTime
    For i = Mode To ptxUPD_DateTime
        Text1(i).Text = ""
    Next i
    Label1(0).Caption = ""

End Sub
Private Function Error_Check_Proc(Mode As Integer, Chk_Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts As Integer
    
    
    Error_Check_Proc = True
    
    Select Case Mode
    
        Case ptxHIN_GAI    '�i�ڊO��
        
                
            Text1(ptxHIN_GAI).Text = StrConv(Text1(ptxHIN_GAI).Text, vbUpperCase)   '2018.04.09
        
        
            Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Label1(0).Caption = StrConv(ITEMREC.HIN_NAME, vbUnicode)
                Case BtErrKeyNotFound
                    If Chk_Mode = 0 Then
                        Label1(0).Caption = ""
                    Else
                        MsgBox "���͂������ڂ̓G���[�ł��i�i�Ԗ��o�^�j"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                    Exit Function
            End Select
    
    
        Case ptxGENSANKOKU
    
            Text1(ptxGENSANKOKU).Text = StrConv(Text1(ptxGENSANKOKU).Text, vbUpperCase)   '2018.04.09
    
    
            If Chk_Mode = 9 Then
            Else
                If Trim(Text1(Mode).Text) = "" Then
                    MsgBox "���͂������ڂ̓G���[�ł��i���Y���j"
                    Text1(Mode).SetFocus
                    Exit Function
                End If
            End If
    
        Case ptxINS_DateTime
    
        
            If Chk_Mode = 9 Then
            Else
                
                
                
                            
                
                If Trim(Text1(Mode).Text) = "" Then
                    Text1(Mode).Text = Format(Now, "YYYYMMDDHHMMSS")
                End If
                
                
                If Len(Trim(Text1(Mode).Text)) <> 14 Then
                    MsgBox "���͂������ڂ̓G���[�ł��i�o�^�����j"
                    Text1(Mode).SetFocus
                    Exit Function
                Else
                    If Not IsDate(Mid(Text1(Mode).Text, 1, 4) & "/" & Mid(Text1(Mode).Text, 5, 2) & "/" & Mid(Text1(Mode).Text, 7, 2)) Then
                        MsgBox "���͂������ڂ̓G���[�ł��i�o�^�����j"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
            
                    If "00" > Mid(Text1(Mode).Text, 9, 2) Or Mid(Text1(Mode).Text, 9, 2) > "23" Then
                        MsgBox "���͂������ڂ̓G���[�ł��i�o�^�����j"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
            
                    If "00" > Mid(Text1(Mode).Text, 11, 2) Or Mid(Text1(Mode).Text, 11, 2) > "59" Then
                        MsgBox "���͂������ڂ̓G���[�ł��i�o�^�����j"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
            
                    If "00" > Mid(Text1(Mode).Text, 13, 2) Or Mid(Text1(Mode).Text, 13, 2) > "59" Then
                        MsgBox "���͂������ڂ̓G���[�ł��i�o�^�����j"
                        Text1(Mode).SetFocus
                        Exit Function
                    End If
            
            
            
            
            
                End If
                
            End If
        
    End Select
        
    Error_Check_Proc = False
End Function
Private Function Dislpay_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���R�[�h���e�̕\��
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim i       As Integer

    Dislpay_Proc = True

    
    
    
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text1(ptxHIN_GAI).Text)

    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
            Label1(0).Caption = StrConv(ITEMREC.HIN_NAME, vbUnicode)
        Case BtErrKeyNotFound
            Label1(0).Caption = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
            Exit Function
    End Select
    
    
    
    
    
    
    Call UniCode_Conv(K0_GENSAN.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_GENSAN.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
    
    Call UniCode_Conv(K0_GENSAN.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_GENSAN.GENSANKOKU, Text1(ptxGENSANKOKU).Text)
    
    sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
        
            Call UniCode_Conv(GENSANREC.Ins_DateTime, "")
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "���Y���}�X�^")
            Exit Function
    End Select
    
    
    Text1(ptxINS_DateTime).Text = Trim(StrConv(GENSANREC.Ins_DateTime, vbUnicode))
    
    
    Text1(ptxUPD_DateTime).Text = Trim(StrConv(GENSANREC.UPD_DATETIME, vbUnicode))      '2013.02.19
                
    
    Dislpay_Proc = False
End Function
Private Function Update_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   �ǉ��^�ύX����
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

Dim i               As Integer

    Update_Proc = True

    Select Case Mode
        Case 0
        
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[" & Trim(Text1(ptxHIN_GAI).Text) & "-" & Trim(Text1(ptxGENSANKOKU).Text) & "]" & "�ǉ������J�n" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
        
        
        
            
            
            Call UniCode_Conv(GENSANREC.JGYOBU, Last_JGYOBU)
            Call UniCode_Conv(GENSANREC.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
            
            Call UniCode_Conv(GENSANREC.HIN_GAI, Text1(ptxHIN_GAI).Text)
            Call UniCode_Conv(GENSANREC.GENSANKOKU, Text1(ptxGENSANKOKU).Text)
            
            Call UniCode_Conv(GENSANREC.INS_TANTO, "1130")
            
            If Trim(Text1(ptxINS_DateTime).Text) = "" Then
                Call UniCode_Conv(GENSANREC.Ins_DateTime, Format(Now, "YYYYMMDDHHMMSS"))
            Else
                Call UniCode_Conv(GENSANREC.Ins_DateTime, Trim(Text1(ptxINS_DateTime).Text))
            End If
            Call UniCode_Conv(GENSANREC.UPD_TANTO, "")
            Call UniCode_Conv(GENSANREC.UPD_DATETIME, "")
            Call UniCode_Conv(GENSANREC.FILLER, "")
                    
            
            
            sts = BTRV(BtOpInsert, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
            Select Case sts
                Case BtNoErr
                
                Case BtErrDuplicates
                    
                    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[" & Trim(Text1(ptxHIN_GAI).Text) & "-" & Trim(Text1(ptxGENSANKOKU).Text) & "]" & "�ǉ�������~" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
                    
                    
                    MsgBox "���[�����珑���������܂����B�m�F���Ă��������B"
                    Update_Proc = False
                    Exit Function
                
                Case Else
                    Call File_Error(sts, BtOpInsert, "���Y���}�X�^")
                    Exit Function
            End Select
        
        
        
        
        
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[" & Trim(Text1(ptxHIN_GAI).Text) & "-" & Trim(Text1(ptxGENSANKOKU).Text) & "]" & "�ǉ������I��" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
        
        
        
        
        
        
        Case 1
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[" & Trim(Text1(ptxHIN_GAI).Text) & "-" & Trim(Text1(ptxGENSANKOKU).Text) & "]" & "�X�V�����J�n" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    
            Call UniCode_Conv(GENSANREC.Ins_DateTime, Text1(ptxINS_DateTime).Text)
            
            Call UniCode_Conv(GENSANREC.UPD_TANTO, "1130")
            
            
            
            'Call UniCode_Conv(GENSANREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))                       2013.02.19
            If Trim(Text1(ptxUPD_DateTime).Text) = "" Then                                                  '2013.02.19
                Call UniCode_Conv(GENSANREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))                    '2013.02.19
            Else                                                                                            '2013.02.19
                If Trim(Text1(ptxUPD_DateTime).Text) = StrConv(GENSANREC.UPD_DATETIME, vbUnicode) Then      '2013.02.19
                    Call UniCode_Conv(GENSANREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))                '2013.02.19
                Else                                                                                        '2013.02.19
                    Call UniCode_Conv(GENSANREC.UPD_DATETIME, Trim(Text1(ptxUPD_DateTime).Text))            '2013.02.19
                End If                                                                                      '2013.02.19
            End If                                                                                          '2013.02.19
                    
            
            
            sts = BTRV(BtOpUpdate, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
            Select Case sts
                Case BtNoErr
                
                Case BtErrKeyNotFound
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[" & Trim(Text1(ptxHIN_GAI).Text) & "-" & Trim(Text1(ptxGENSANKOKU).Text) & "]" & "�X�V������~" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
                    MsgBox "���[�����珑���������܂����B�m�F���Ă��������B"
                    Update_Proc = False
                    Exit Function
                
                Case Else
                    Call File_Error(sts, BtOpInsert, "���Y���}�X�^")
                    Exit Function
            End Select
    
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[" & Trim(Text1(ptxHIN_GAI).Text) & "-" & Trim(Text1(ptxGENSANKOKU).Text) & "]" & "�X�V�����I��" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
    
    End Select





    Update_Proc = False

End Function


Private Function DELETE_Proc() As Integer
'----------------------------------------------------------------------------
'                   �폜����
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

Dim i               As Integer

    DELETE_Proc = True

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[" & Trim(Text1(ptxHIN_GAI).Text) & "-" & Trim(Text1(ptxGENSANKOKU).Text) & "]" & "�폜�����J�n" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
        
    Call UniCode_Conv(K0_GENSAN.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_GENSAN.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
    
    Call UniCode_Conv(K0_GENSAN.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_GENSAN.GENSANKOKU, Text1(ptxGENSANKOKU).Text)
    
    sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
    Select Case sts
        Case BtNoErr
        
        
        
        Case BtErrKeyNotFound

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[" & Trim(Text1(ptxHIN_GAI).Text) & "-" & Trim(Text1(ptxGENSANKOKU).Text) & "]" & "�폜������~" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
            DELETE_Proc = False
            Exit Function
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "���Y���}�X�^")
            Exit Function
    End Select

            
    
    
    sts = BTRV(BtOpDelete, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[" & Trim(Text1(ptxHIN_GAI).Text) & "-" & Trim(Text1(ptxGENSANKOKU).Text) & "]" & "�폜������~" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)
        
        Case Else
            Call File_Error(sts, BtOpInsert, "���Y���}�X�^")
            Exit Function
    End Select

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "[" & Trim(Text1(ptxHIN_GAI).Text) & "-" & Trim(Text1(ptxGENSANKOKU).Text) & "]" & "�폜�����I��" & Format(Now, "YYYY/MM/DD HH:MM:SS"), Me.hwnd, 0)

    DELETE_Proc = False

End Function


Private Sub Command_Click(Index As Integer)

Dim yn  As Integer
Dim sts As Integer
Dim i   As Integer

    Select Case Index
        Case 0, 1
                                            
            
            For i = 0 To ptxINS_DateTime
                If Error_Check_Proc(i, Index) Then    '�G���[�`�F�b�N
                    Exit Sub
                End If
            Next i
                
            If Index = 0 Then
                
                
                
                
                
                Call UniCode_Conv(K0_GENSAN.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_GENSAN.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
                
                Call UniCode_Conv(K0_GENSAN.HIN_GAI, Text1(ptxHIN_GAI).Text)
                Call UniCode_Conv(K0_GENSAN.GENSANKOKU, Text1(ptxGENSANKOKU).Text)
                
                sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                Select Case sts
                    Case BtNoErr
                    
                        MsgBox "�w��̃R�[�h�́A���R�[�h�o�^�ςł��B"
                        Exit Sub
                    
                    Case BtErrKeyNotFound
                    
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "���Y���}�X�^")
                        Unload Me
                End Select
                
                
                
                
                
                
                
                
                
                yn = MsgBox("�ǉ����܂����H", vbYesNo + vbQuestion, "�m�F����")
            Else
                
                
                
                
                
                Call UniCode_Conv(K0_GENSAN.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_GENSAN.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
                
                Call UniCode_Conv(K0_GENSAN.HIN_GAI, Text1(ptxHIN_GAI).Text)
                Call UniCode_Conv(K0_GENSAN.GENSANKOKU, Text1(ptxGENSANKOKU).Text)
                
                sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                    
                    
                        MsgBox "�w��̃R�[�h�́A���R�[�h���o�^�ł��B"
                        Exit Sub
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "���Y���}�X�^")
                        Unload Me
                End Select
                
                
                
                
                
                
                
                
                
                
                
                
                
                yn = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            End If
            
            If yn = vbYes Then
                Call Input_Lock
                If Update_Proc(Index) Then
                    Unload Me
                End If
                Call Input_UnLock
                
'                Call Clear_Field(0)            '2108.04.09
                Call Clear_Field(1)             '2018.04.09
                
                Call Input_Lock                 '2018.04.09
                If List_Proc() Then             '2018.04.09
                    Unload Me                   '2018.04.09
                End If                          '2018.04.09
                Call Input_UnLock               '2018.04.09
                
                
                List1.SetFocus
            End If
            
        Case 3
        
        
                Call UniCode_Conv(K0_GENSAN.JGYOBU, Last_JGYOBU)
                Call UniCode_Conv(K0_GENSAN.NAIGAI, Right(Combo1(pcmbNAIGAI).Text, 1))
                
                Call UniCode_Conv(K0_GENSAN.HIN_GAI, Text1(ptxHIN_GAI).Text)
                Call UniCode_Conv(K0_GENSAN.GENSANKOKU, Text1(ptxGENSANKOKU).Text)
                
                sts = BTRV(BtOpGetEqual, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
                Select Case sts
                    Case BtNoErr
                    
                    Case BtErrKeyNotFound
                    
                    
                        MsgBox "�w��̃R�[�h�́A���R�[�h���o�^�ł��B"
                        Exit Sub
                    
                    
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "���Y���}�X�^")
                        Unload Me
                End Select
        
        
        
        
            yn = MsgBox("�폜���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If yn = vbYes Then
                Call Input_Lock
                If DELETE_Proc() Then
                    Unload Me
                End If
                Call Input_UnLock
                Text1(ptxGENSANKOKU).SetFocus
            
'                Call Clear_Field(0)            '2108.04.09
                Call Clear_Field(1)             '2018.04.09
                
                Call Input_Lock                 '2018.04.09
                If List_Proc() Then             '2018.04.09
                    Unload Me                   '2018.04.09
                End If                          '2018.04.09
                Call Input_UnLock               '2018.04.09
                
                List1.SetFocus
            
            
            End If
        
        Case 4
        
            Call Input_Lock
            If List_Proc() Then
                Unload Me
            End If
            Call Input_UnLock
            List1.SetFocus
        
        Case 11
            Unload Me
        Case Else
            Beep
    End Select
    

End Sub


Private Sub Form_DblClick()
'    PrintForm          '2018.04.09
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
Dim i   As Integer
    
    Select Case KeyCode
        
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
        Case vbKeyZ
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim c       As String * 128
Dim sts     As Integer
Dim i       As Integer


'    If App.PrevInstance Then                       2018.04.09
'        Beep                                       2018.04.09
'        MsgBox "����v���O�������s���ł��B"        2018.04.09
'        End                                        2018.04.09
'    End If                                         2018.04.09
                                
                                
                                
                                
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "���Y��Ͻ�����ݽ" & LAST_UPDATE_DAY, Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
                                
                                
                                
                                '���O�t�@�C������荞��
'    If GetIni("FILE", "LOGF", "SYS", c) Then
    If GetIni(App.EXEName, "LOGF", App.EXEName, c) Then
        If GetIni("FILE", "LOGF", "SYS", c) Then
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
            End
        End If
    End If
    LOG_F = RTrim(c)

                                
                                
                                
                                
                                
                                
                                
                                
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
        
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1011301.Caption = "���Y��Ͻ�����ݽ�i" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
                                
    Combo1(pcmbNAIGAI).AddItem NAIGAI1$ & " " & NAIGAI_NAI$
    Combo1(pcmbNAIGAI).AddItem NAIGAI2$ & " " & NAIGAI_GAI$
    Combo1(pcmbNAIGAI).ListIndex = 0
                                
                                
                                
                                
    Text1(ptxDISP_COUNT).Text = 100
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '���Y���}�X�^�n�o�d�m
    If GENSAN_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Show
                                
                                
    
    End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '���Y���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���Y���}�X�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, GENSAN_POS, GENSANREC, Len(GENSANREC), K0_GENSAN, Len(K0_GENSAN), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set F1011301 = Nothing
    End
End Sub

Private Sub List1_DblClick()

        Text1(ptxHIN_GAI).Text = Mid(List1.List(List1.ListIndex), 1, 20)
        Text1(ptxGENSANKOKU).Text = Mid(List1.List(List1.ListIndex), 21, 20)

        If Dislpay_Proc() Then
            Unload Me
        End If

        Text1(ptxGENSANKOKU).SetFocus

End Sub

Private Sub List1_GotFocus()
    
    If List1.ListCount > 0 Then
        If List1.ListIndex <= 0 Then
            List1.ListIndex = 0
        End If
    End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sts As Integer

    Select Case KeyCode
        Case vbKeyReturn
            
            Call List1_DblClick
    
    End Select

End Sub





Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------
Dim i As Integer

    F1011301.MousePointer = vbHourglass

    Call Ctrl_Lock(F1011301)

End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1011301)

    F1011301.MousePointer = vbDefault

End Sub


Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    
                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If
    
    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1011301.Caption = "���Y��Ͻ�����ݽ�i" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index, 0) Then    '�G���[�`�F�b�N
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

Private Sub Text1_LostFocus(Index As Integer)

'>>>>>>>    2018.04.09
    Select Case Index
            
        Case ptxHIN_GAI
            Text1(ptxHIN_GAI).Text = StrConv(Text1(ptxHIN_GAI).Text, vbUpperCase)   '2018.04.09


        Case ptxGENSANKOKU
            Text1(ptxGENSANKOKU).Text = StrConv(Text1(ptxGENSANKOKU).Text, vbUpperCase)   '2018.04.09


    End Select
'>>>>>>>    2018.04.09


End Sub
