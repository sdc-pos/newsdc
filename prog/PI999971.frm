VERSION 5.00
Begin VB.Form PI999971 
   Caption         =   "���i���x���ꊇ���s ([PI99997]�@2011.08.05 11:00"
   ClientHeight    =   6264
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   10188
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
   ScaleHeight     =   6264
   ScaleWidth      =   10188
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text2 
      Height          =   3855
      Left            =   420
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   5412
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�ر�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   420
      TabIndex        =   10
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      Index           =   2
      Left            =   8880
      TabIndex        =   9
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      Index           =   1
      Left            =   8880
      TabIndex        =   8
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      Index           =   0
      Left            =   4620
      TabIndex        =   5
      Top             =   5160
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   3888
      Left            =   6156
      TabIndex        =   2
      Top             =   1200
      Width           =   3585
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8772
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7416
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "�m�f����"
      Height          =   252
      Index           =   4
      Left            =   7932
      TabIndex        =   7
      Top             =   5760
      Width           =   1068
   End
   Begin VB.Label Label1 
      Caption         =   "�n�j����"
      Height          =   252
      Index           =   3
      Left            =   7932
      TabIndex        =   6
      Top             =   5280
      Width           =   1068
   End
   Begin VB.Label Label1 
      Caption         =   "�Ǎ��݌���"
      Height          =   252
      Index           =   2
      Left            =   3360
      TabIndex        =   4
      Top             =   5280
      Width           =   1272
   End
   Begin VB.Label Label1 
      Caption         =   "�������"
      Height          =   252
      Index           =   1
      Left            =   6156
      TabIndex        =   3
      Top             =   960
      Width           =   960
   End
End
Attribute VB_Name = "PI999971"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IN_cnt  As Integer
Private OK_cnt  As Integer
Private NG_cnt  As Integer


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PI999971.MousePointer = vbHourglass

    Call Ctrl_Lock(PI999971)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI999971)


    PI999971.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer, chk As Integer, flg) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts         As Integer
Dim yn          As Integer
    
    
    Error_Check_Proc = True
    
        
        
    Error_Check_Proc = False
    

End Function

Private Sub Command1_Click(Index As Integer)

Dim ans             As Integer
Dim i               As Integer




Dim com             As Integer
Dim sts             As Integer

Dim Skip_F          As Boolean

Dim wkLine          As Variant
Dim wkItem          As Variant


Dim Parts_F             As Integer
Dim Gaisou_F            As Integer
Dim Kishu_F             As Integer
Dim GAISOU_QTY          As Long
Dim GAISOU_SHIJI_QYU    As Long



'=============================== 2011.08.04 =====
Dim Parts       As String   '�i��
Dim ID          As Long     '�w����

Dim PartsLabel  As Integer  '�i������ 0:�Ȃ� �ȊO�F����
Dim KisyuLabel  As Integer  '�@������ 0:�Ȃ�
Dim JanLabel    As Integer  'JAN���� 0:�Ȃ�
Dim GLabel      As Integer  '�O������ 0:�Ȃ�
Dim ItemLabel   As Integer  '�������ٖ���

Dim OrderNo     As String
Dim ItemNo      As String

Dim Pri_Date    As String

Dim L_QTY       As Long

Dim wkDate      As String * 10
Dim NGItem      As String * 20
'=============================== 2011.08.04 =====




Dim objAccess       As Access.Application
Dim strAccessPath   As String



    Select Case Index
        Case 0              '���
            
            
            
            Beep
            ans = MsgBox("������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                
                
                
                
                List2.Clear
                
                OK_cnt = 0
                NG_cnt = 0
                Text1(1).Text = Format(OK_cnt, "#,##0")
                Text1(2).Text = Format(NG_cnt, "#,##0")
                
                    
                    
                                    
                wkLine = Split(Text2.Text, vbCrLf, -1)
                                    
                                    
                For i = 0 To UBound(wkLine)
                    
                    
                    wkItem = Split(wkLine(i), vbTab, -1)
                    
                    Skip_F = False
                    
                    If UBound(wkItem) < 2 Then
                    Else
                        
                        
                        If Not IsNumeric(wkItem(2)) Then
                            Skip_F = True
                        Else
                            If CInt(wkItem(2)) = 0 Then
                                Skip_F = True
                            End If
                        End If
                        
                        wkDate = ""
                        If UBound(wkItem) >= 3 Then
                            If Trim(wkItem(3)) <> "" Then
                                If Not IsDate(wkItem(3)) Then
                                    Skip_F = True
                                Else
                                    wkDate = Format(wkItem(3), "YYYY/MM/DD")
                                End If
                            End If
                        End If
                                                                        
                                                                        
                        If Not Skip_F Then
                        
                        
                            Call UniCode_Conv(K0_ITEM.JGYOBU, Format(wkItem(0)))
                            Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                            Call UniCode_Conv(K0_ITEM.HIN_GAI, Format(wkItem(1)))
                    
                    
                            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                
                            Select Case sts
                                Case BtNoErr
                                Case BtErrKeyNotFound
                                    
                                    Skip_F = True
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                    Exit Sub
                            
                            End Select
                        End If
                
                
                
                
                        If Skip_F Then
                            
                            NGItem = wkItem(1)
                            List2.AddItem wkItem(0) & " " & NGItem & " " & "NG"
                            
                            NG_cnt = NG_cnt + 1
                            Text1(2).Text = Format(NG_cnt, "#,##0")
                        Else
                        
                        
                                On Error Resume Next
                                Set objAccess = GetObject(, "Access.Application")
                                If Err().Number <> 0 Then
                                    MsgBox "���̒[���ł͏��i���x�����s�͍s���܂���B"
            '                        MsgBox "GetObject(Access.Application)" & Err().Number & " " & Err().Description
                                Else
            '                        MsgBox Err.Number
                                    
                                    strAccessPath = App.Path
                                    If Right(strAccessPath, 1) <> "\" Then
                                        strAccessPath = strAccessPath & "\"
                                    End If
                                    
                                    strAccessPath = strAccessPath & "litem.mdb"
                                    Set objAccess = GetObject(strAccessPath)
                
                                
                        
                                    Parts_F = 1
                                    
                                    
                                    Gaisou_F = 0
                                    
                                    Kishu_F = 0
                                    
                                    GAISOU_QTY = 0
                                    
                                    GAISOU_SHIJI_QYU = 0
                                    
                                    com = BtOpGetFirst
                                    Do
                                    
                                    
                                    
                                        sts = BTRV(com, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                                        Select Case sts
                                            Case BtNoErr
                                            
                                                sts = BTRV(BtOpDelete, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                                                
                                                Select Case sts
                                                
                                                    Case BtNoErr
                                                    Case Else
                                                        Call File_Error(sts, com, "���ٗp�i��Ͻ�")
                                                        Exit Sub
                                                End Select
                                            
                                            Case BtErrEOF
                                                Exit Do
                                            Case Else
                                                Call File_Error(sts, com, "���ٗp�i��Ͻ�")
                                                Exit Sub
                                        End Select
                                        
                                        com = BtOpGetNext
                                    
                                    
                                    Loop
                                    
                                    Call UniCode_Conv(K0_ITEM.JGYOBU, Format(wkItem(0)))
                                    Call UniCode_Conv(K0_ITEM.NAIGAI, "1")
                                    Call UniCode_Conv(K0_ITEM.HIN_GAI, Format(wkItem(1)))
                            
                            
                                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                                    Select Case sts
                                        Case BtNoErr
                                        
                                            
                                            Call UniCode_Conv(ITEMREC.L_IRI_QTY, "00000000")
                                            
                                            
                                            sts = BTRV(BtOpInsert, L_ITEM_POS, ITEMREC, Len(ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
                                            Select Case sts
                                                Case BtNoErr
                                        
                                                    
                                                    
                                                    '=================================================2011.08.04
'                                                    objAccess.Run "PosPrintLabel", Trim(Format(wkItem(1))), CLng(wkItem(2)), Parts_F, Gaisou_F, Kishu_F, GAISOU_QTY, GAISOU_SHIJI_QYU, 0
                                        
                                        
                                        
                                                    PartsLabel = 0
                                                    KisyuLabel = 0
                                                    JanLabel = 0
                                                    GLabel = 0
                                                    ItemLabel = 0


                                                    '�i�ڃR�[�h
                                                    Parts = wkItem(1)
                                                    '�p�[�c���x��
                                                    PartsLabel = CLng(wkItem(2))
                                                    'ID
                                                    ID = 0
                                                    '�A�C�e�����x��
                                                    ItemLabel = 0
                                                    '�I�[�_�[��
                                                    OrderNo = ""
                                                    '�A�C�e����
                                                    ItemNo = ""
                                                    '������t
                                                    Pri_Date = wkDate
                                                    '����
                                                    L_QTY = 1
                                                    objAccess.Run "NewPosPrintLabel", _
                                                                        Trim(Parts), _
                                                                        PartsLabel, _
                                                                        KisyuLabel, _
                                                                        JanLabel, _
                                                                        GLabel, _
                                                                        ID, _
                                                                        ItemLabel, _
                                                                        Trim(OrderNo), _
                                                                        Trim(ItemNo), _
                                                                        Pri_Date, _
                                                                        L_QTY
                                                    '=================================================2011.08.04
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                                Case Else
                                                    Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                                    Exit Sub
                                                
                                        
                                            End Select
                                        
                                        Case BtErrKeyNotFound
                                            
                                        Case Else
                                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                                            Exit Sub
                                    
                                    End Select
                                    
                                    
                                    
                                    
                                                            
                                    Set objAccess = Nothing
                                End If
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                        
                            OK_cnt = OK_cnt + 1
                            Text1(1).Text = Format(OK_cnt, "#,##0")
                        
                        
                        End If
            
                    End If
                Next i
            
                MsgBox "������I�����܂����B"
            
            
            End If
        Case 1              '�I��
            Unload Me
    
        Case 2
            Text2.Text = ""
    End Select

ErrHandler:
    
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

Dim c           As String * 128
Dim sts         As Integer
Dim i           As Integer

Dim MUKE_CODE   As Variant


    If App.PrevInstance Then
        MsgBox "����v���O�������s���ł��B"
        End
    End If

                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
    LOG_F = RTrim(c)
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
                                '���x���p�i�ڃ}�X�^�n�o�d�m
    If L_ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    
    
    
    
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
                                            
                                            
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
    
                                            '���x���p�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, L_ITEM_POS, L_ITEMREC, Len(L_ITEMREC), K0_L_ITEM, Len(K0_L_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���x���p�i�ڃ}�X�^")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PI999971 = Nothing

    End
End Sub


Private Sub RichTextBox1_Change()

End Sub

