VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form ITEM_MAKE1 
   Caption         =   "品目ﾏｽﾀ作成"
   ClientHeight    =   9510
   ClientLeft      =   2025
   ClientTop       =   -5445
   ClientWidth     =   15210
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  '手動
   ScaleHeight     =   9510
   ScaleWidth      =   15210
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton Command1 
      Caption         =   "終 了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   7
      ToolTipText     =   "処理を終了します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      OLEDragMode     =   1  '自動
      OLEDropMode     =   1  '手動
      TabIndex        =   6
      Top             =   1200
      Width           =   6975
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7335
      Left            =   1680
      TabIndex        =   2
      Top             =   1680
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   12938
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "№"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "事"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "国"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "対内品番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "品名"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "棚番"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1005"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=900"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=661"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=556"
      Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=0"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=900"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=794"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=5159"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=5054"
      Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=0"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=4339"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=4233"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(5).Width=3281"
      Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=3175"
      Splits(0)._ColumnProps(26)=   "Column(5)._ColStyle=1"
      Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   0
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      OLEDropMode     =   1
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ ゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1,.bgcolor=&HFFFF00&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(27)  =   ":id=68,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(28)  =   ":id=68,.fontname=ＭＳ ゴシック"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=69,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=71,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=74,.parent=9,.bgcolor=&HFFFF00&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=75,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=82,.parent=67,.alignment=3"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=68"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=69"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=71"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=94,.parent=67,.alignment=0"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=91,.parent=68"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=92,.parent=69"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=93,.parent=71"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=24,.parent=67"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=21,.parent=68"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=22,.parent=69"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=23,.parent=71"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=98,.parent=67,.alignment=0"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=95,.parent=68"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=96,.parent=69"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=97,.parent=71"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=20,.parent=67"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=17,.parent=68"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=18,.parent=69"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=19,.parent=71"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=102,.parent=67,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=99,.parent=68"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=100,.parent=69"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=101,.parent=71"
      _StyleDefs(62)  =   "Named:id=33:Normal"
      _StyleDefs(63)  =   ":id=33,.parent=0"
      _StyleDefs(64)  =   "Named:id=34:Heading"
      _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(66)  =   ":id=34,.wraptext=-1"
      _StyleDefs(67)  =   "Named:id=35:Footing"
      _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(69)  =   "Named:id=36:Selected"
      _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(71)  =   "Named:id=37:Caption"
      _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(73)  =   "Named:id=38:HighlightRow"
      _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(75)  =   "Named:id=39:EvenRow"
      _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(77)  =   "Named:id=40:OddRow"
      _StyleDefs(78)  =   ":id=40,.parent=33"
      _StyleDefs(79)  =   "Named:id=41:RecordSelector"
      _StyleDefs(80)  =   ":id=41,.parent=34"
      _StyleDefs(81)  =   "Named:id=42:FilterBar"
      _StyleDefs(82)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新 規"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "入荷予定データを登録します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "読 込"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "ファイル名"
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblDisp_Count 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9960
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "読込件数"
      Height          =   255
      Index           =   1
      Left            =   8880
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "ITEM_MAKE1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim PLN_Y_NYUKA         As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 5              '最大列数

Private Const colNO% = 0                '№
Private Const colJGYOBU% = 1            '品番(外部)
Private Const colNAIGAI% = 2            '品番(外部)

Private Const colHIN_GAI% = 3           '品番(外部)

Private Const colHIN_NAME% = 4          '品名


Private Const colST_TANABAN% = 5              '入荷予定日




Private EXCEL_DATA  As Variant





Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Long
Dim j           As Long


    Select Case Index



        Case 0          '読込み



            '取込みﾃﾞｰﾀ表示
            
            If List_Disp_SA_Proc() Then
                Unload Me
            End If



        Case 1          '登録-->新規    2012.02.13


            If Update_Proc(1) Then
                Unload Me
            End If




        Case 3          '終了   2-->3   2012.02.13

            Unload Me
    
    End Select



'    Command1(Index).SetFocus


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    
    Select Case KeyCode
        Case vbKeyF12
            Unload Me
    End Select
    
    
    

End Sub

Private Sub Form_Load()


Dim c       As String * 128






    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)

                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(0) Then
        Unload Me
    End If




End Sub


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Text1.Text = Trim(Data.Files(1))


    Command1(0).Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    
    
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    Set ITEM_MAKE1 = Nothing



    End

End Sub

Private Sub SHORI_Click(Index As Integer)

    Select Case Index
    
        Case 0
            Command1(0).Value = True
        Case 1
            Command1(1).Value = True
        Case 2
            Command1(2).Value = True
    End Select



End Sub

Private Sub TDBGrid1_OLEDragDrop(ByVal Data As TrueDBGrid80.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Text1.Text = Trim(Data.Files(0))
'    Text1.Text = Data.GetData(0)


    Command1(0).Value = True


End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Text1.Text = Trim(Data.Files(1))
    
    Command1(0).Value = True


'    If Data.GetFormat(vbCFText) Then
'        Text1.Text = Data.GetData(vbCFText)
'        Command1(0).Value = True
'    End If

End Sub

Private Function Update_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   「商品化用入荷予定ファイル」登録処理
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    
Dim INS_NOW         As String * 14
    
Dim KEY_NO          As Long

Dim Row             As Long

'Dim c               As String * 128
'Dim FullPath        As String
    
Dim wkTanaban       As String * 8
    
    
    
    If PLN_Y_NYUKA.Count(1) = 0 Then
        Exit Function
    End If



    Update_Proc = True
    
    Call Input_Lock

                                    
    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
                                    
                                    'テーブルリセット
    
    For Row = 1 To PLN_Y_NYUKA.UpperBound(1)
        
        
        DoEvents
        
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.13
        Call Rclr_ITEMREC
            
            
        Call UniCode_Conv(ITEMREC.JGYOBU, PLN_Y_NYUKA(Row, colJGYOBU))
        Call UniCode_Conv(ITEMREC.NAIGAI, PLN_Y_NYUKA(Row, colNAIGAI))
        Call UniCode_Conv(ITEMREC.HIN_GAI, PLN_Y_NYUKA(Row, colHIN_GAI))
        Call UniCode_Conv(ITEMREC.HIN_NAME, PLN_Y_NYUKA(Row, colHIN_NAME))
            
        wkTanaban = PLN_Y_NYUKA(Row, colST_TANABAN)
        
        Call UniCode_Conv(ITEMREC.ST_SOKO, Mid(wkTanaban, 1, 2))
        Call UniCode_Conv(ITEMREC.ST_RETU, Mid(wkTanaban, 3, 2))
        Call UniCode_Conv(ITEMREC.ST_REN, Mid(wkTanaban, 5, 2))
        Call UniCode_Conv(ITEMREC.ST_DAN, Mid(wkTanaban, 7, 2))
                    
        
        Call UniCode_Conv(ITEMREC.ST_SET_DT, Format(Now, "YYYYMMDD"))

        Call UniCode_Conv(ITEMREC.INS_TANTO, App.EXEName)
        Call UniCode_Conv(ITEMREC.Ins_DateTime, INS_NOW)
                        
        Do
            sts = BTRV(BtOpInsert, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                
                
                
                Case BtErrDuplicates
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    
                    Beep
                    ans = MsgBox("「品目マスタ」他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Call Input_UnLock
                        Exit Function
                    End If
                
                Case Else
                    Call Input_UnLock
                    Call File_Error(sts, BtOpInsert, "品目マスタ")
                    Exit Function
            End Select
                
        Loop
            
    
        Set TDBGrid1.Array = PLN_Y_NYUKA
        TDBGrid1.ReBind
        
        TDBGrid1.Update
        TDBGrid1.Bookmark = Row
        
    Next Row


    Set TDBGrid1.Array = PLN_Y_NYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst







    Update_Proc = False
    Call Input_UnLock
    Exit Function

Error_Proc:
    
    MsgBox "Err.Number= " & Err.Number & " " & Err.Description
    Call Input_UnLock

End Function

Private Function List_Disp_SA_Proc() As Integer
'----------------------------------------------------------------------------
'                   「商品化用入荷予定ファイル」読込み処理 サファイヤ
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
    
Dim INS_NOW         As String * 14
    
    
    

Dim xlApp           As Object
Dim xlBook          As Object
Dim xlSheet         As Object

Dim Row             As Long

Dim END_GYO         As Integer
Dim SKIP_F          As Boolean

Dim wkJGYOBU        As String * 1
Dim wkNAIGAI        As String * 1
Dim wkHIN_GAI       As String * 20
Dim wkHIN_NAME      As String * 40
Dim wkST_TANABAN    As String * 8



Dim i               As Long

    List_Disp_SA_Proc = True

    Call Input_Lock







    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlApp.Worksheets(1)
    
    
    On Error GoTo Error_Proc

    '2011.12.03
    xlApp.Workbooks.Open (Text1.Text), ReadOnly:=True, UpdateLinks:=0

    On Error GoTo 0




    Set PLN_Y_NYUKA = Nothing
    
    Row = Min_Row - 1
    lblDisp_Count.Caption = ""





    END_GYO = 0
    For i = 1 To 1048576
        
        SKIP_F = False
        
        
        On Error GoTo Error_Proc        '2011.12.03

            
        If Trim(xlSheet.Application.Cells(i, 1)) = "" And _
            Trim(xlSheet.Application.Cells(i, 2)) = "" And _
            Trim(xlSheet.Application.Cells(i, 3)) = "" Then
        
            SKIP_F = True
            END_GYO = END_GYO + 1
            
            If END_GYO > 5 Then
                Exit For
            End If
        Else
            
            
            
            END_GYO = 0
        
            If Trim(xlSheet.Application.Cells(i, 1)) = "" And _
                Trim(xlSheet.Application.Cells(i, 2)) = "" And _
                Trim(xlSheet.Application.Cells(i, 3)) = "" Then
        
                SKIP_F = True
            
            Else
                    
                    
                '品番
                wkJGYOBU = Trim(xlSheet.Application.Cells(i, 1))
                    
                '品番
                wkNAIGAI = Trim(xlSheet.Application.Cells(i, 2))
                    
                '品番
                wkHIN_GAI = Trim(xlSheet.Application.Cells(i, 3))
                '品番
                wkHIN_NAME = Trim(xlSheet.Application.Cells(i, 4))
                '品番
                wkST_TANABAN = Trim(xlSheet.Application.Cells(i, 5))
                
                
                If Not SKIP_F Then
            
                    Row = Row + 1
                    PLN_Y_NYUKA.ReDim Min_Row, Row, Min_Col, Max_Col
                    
                    PLN_Y_NYUKA(Row, colNO) = Row
                
                
                    PLN_Y_NYUKA(Row, colJGYOBU) = wkJGYOBU
                    PLN_Y_NYUKA(Row, colNAIGAI) = wkNAIGAI
                
                
    
                    PLN_Y_NYUKA(Row, colHIN_GAI) = Trim(wkHIN_GAI)
                    PLN_Y_NYUKA(Row, colHIN_NAME) = Trim(wkHIN_NAME)
                    PLN_Y_NYUKA(Row, colST_TANABAN) = Trim(wkST_TANABAN)
    
                End If
            End If
        End If
    
    
    
    Next i

    On Error GoTo 0                 '2011.12.03


    Set TDBGrid1.Array = PLN_Y_NYUKA
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst






    lblDisp_Count.Caption = Format(Row, "#0") & "件"











    Call Input_UnLock

    xlApp.DisplayAlerts = False
    xlBook.Close False
    xlApp.Quit 'EXCELを閉じる
    Set xlApp = Nothing

    List_Disp_SA_Proc = False
    Exit Function

Error_Proc:
    

    Select Case Err.Number
        
        '52 ファイル名または番号が不正です。
        '53 ファイルが見つかりません。
        '54 ファイル モードが不正です。
        '55 ファイルは既に開かれています。
        '57 デバイス I/O エラーです。
        '59 レコード長が一致しません。
        '61 ディスクの空き容量が不足しています。
        '62 ファイルにこれ以上データがありません。
        '63 レコード番号が不正です。
        '68 デバイスが準備されていません。
        '70 書き込みできません。
        '71 ディスクが準備されていません。
        '75 パス名が無効です。
        '76 パスが見つかりません。
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "指定のファイルが見つかりません。" & Chr(13) & Chr(10) & "正しいファイル名を入力してください。"
            
            
            
            List_Disp_SA_Proc = False      '


    '2011.12.03
        Case 13
        
            MsgBox "読込み対象のEXCELデータに異常が有ります。内容を確認後、再実行してください。"
            
            xlApp.DisplayAlerts = False
            xlBook.Close False
            xlApp.Quit 'EXCELを閉じる
            Set xlApp = Nothing
            
            
            
            
            
            
            List_Disp_SA_Proc = False      '
            


        Case Else
            MsgBox Err.Description
    
    '2011.12.03
    End Select
    
    On Error GoTo 0
    
    
    Call Input_UnLock

End Function




Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    ITEM_MAKE1.MousePointer = vbHourglass

    Call Ctrl_Lock(ITEM_MAKE1)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(ITEM_MAKE1)


    ITEM_MAKE1.MousePointer = vbDefault

End Sub



