VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form B_ITEM_MAKE1 
   Caption         =   "[美的品番管理]美的品番ﾃﾞｰﾀ作成処理[B_ITEM_MAKE] 2013.10.22 16:30"
   ClientHeight    =   12300
   ClientLeft      =   2025
   ClientTop       =   -5445
   ClientWidth     =   14865
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
   ScaleHeight     =   12300
   ScaleWidth      =   14865
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
      Left            =   2040
      TabIndex        =   7
      ToolTipText     =   "処理を終了します"
      Top             =   0
      Width           =   1380
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   9975
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   17595
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
      Columns(3).Caption=   "対外品番"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "対内品番"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "美的品番"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "品　　　名"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=741"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=635"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=2"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=661"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=556"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=556"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=450"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=4128"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=4022"
      Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=0"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(4).Width=4128"
      Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=4022"
      Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=0"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(5).Width=8387"
      Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=8281"
      Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=0"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(30)=   "Column(6).Width=5953"
      Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=5847"
      Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
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
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=82,.parent=67,.alignment=1"
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
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=20,.parent=67,.alignment=0"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=17,.parent=68"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=18,.parent=69"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=19,.parent=71"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=102,.parent=67,.alignment=0"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=99,.parent=68"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=100,.parent=69"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=101,.parent=71"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=16,.parent=67"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=13,.parent=68"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=14,.parent=69"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=15,.parent=71"
      _StyleDefs(66)  =   "Named:id=33:Normal"
      _StyleDefs(67)  =   ":id=33,.parent=0"
      _StyleDefs(68)  =   "Named:id=34:Heading"
      _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   ":id=34,.wraptext=-1"
      _StyleDefs(71)  =   "Named:id=35:Footing"
      _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   "Named:id=36:Selected"
      _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(75)  =   "Named:id=37:Caption"
      _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(77)  =   "Named:id=38:HighlightRow"
      _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(79)  =   "Named:id=39:EvenRow"
      _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(81)  =   "Named:id=40:OddRow"
      _StyleDefs(82)  =   ":id=40,.parent=33"
      _StyleDefs(83)  =   "Named:id=41:RecordSelector"
      _StyleDefs(84)  =   ":id=41,.parent=34"
      _StyleDefs(85)  =   "Named:id=42:FilterBar"
      _StyleDefs(86)  =   ":id=42,.parent=33"
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
      Left            =   360
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
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      OLEDragMode     =   1  '自動
      OLEDropMode     =   1  '手動
      TabIndex        =   6
      Top             =   720
      Width           =   12135
   End
   Begin VB.Label Label1 
      Caption         =   "ファイル名"
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblDisp_Count 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "読込件数"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
End
Attribute VB_Name = "B_ITEM_MAKE1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim B_ITEM      As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 6              '最大列数

Private Const colNO% = 0                '№
Private Const colJGYOBU% = 1            '事業部
Private Const colNAIGAI% = 2            '国内外

Private Const colHIN_GAI% = 3           '品番(外部)
Private Const colHIN_NAI% = 4           '品番(内部)

Private Const colB_HIN_CODE% = 5        '美的品番

Private Const colHIN_NAME% = 6          '品名




Dim SET_JGYOBU  As String * 1           '事業部
Dim SET_NAIGAI  As String * 1           '国内外

Dim TITLE_GYO   As Long                 'タイトル行数

Private EXCEL_DATA  As Variant





Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Long
Dim j           As Long


    Select Case Index



        Case 0          '読込み



            '取込みﾃﾞｰﾀ表示
            If List_Disp_B_ITEM_Proc() Then
                Unload Me
            End If



        Case 1          '新規


            If Update_Proc(1) Then
                Unload Me
            End If




        Case 3          '終了

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
                                
                                
                                
                                '事業部     取り込み
    If GetIni(App.EXEName, "JGYOBU", App.EXEName, c) Then
        MsgBox "[B_ITEM_MAKE.INI][JGYOBU] 設定ｴﾗｰ"
        End
    Else
        If Trim(c) = "" Then
            MsgBox "[B_ITEM_MAKE.INI][JGYOBU] 設定ｴﾗｰ"
            End
        End If
        SET_JGYOBU = Trim(c)
    
    End If
                                
                                '内外     取り込み
    If GetIni(App.EXEName, "NAIGAI", App.EXEName, c) Then
        MsgBox "[B_ITEM_MAKE.INI][NAIGAI] 設定ｴﾗｰ"
        End
    Else
        If Trim(c) <> "1" And Trim(c) <> "2" Then
            MsgBox "[B_ITEM_MAKE.INI][NAIGAI] 設定ｴﾗｰ"
            End
        End If
        SET_NAIGAI = Trim(c)
    
    End If
                                
                                
                                'タイトル行数     取り込み
    If GetIni(App.EXEName, "TITLE", App.EXEName, c) Then
            TITLE_GYO = 0
    Else
        If Not IsNumeric(Trim(c)) Then
            TITLE_GYO = 0
        Else
            TITLE_GYO = Val(c)
        End If
    End If
                                
                                
                                
                                
                                
                                
                                
                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '美的品番管理ﾃﾞｰﾀＯＰＥＮ
    If B_ITEM_Open(BtOpenNomal) Then
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
    
    
    
    sts = BTRV(BtOpClose, B_ITEM_POS, B_ITEMREC, Len(B_ITEMREC), K0_B_ITEM, Len(K0_B_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "美的品番管理ﾃﾞｰﾀ")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    Set B_ITEM_MAKE1 = Nothing



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


    Command1(0).Value = True


End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Text1.Text = Trim(Data.Files(1))
    
    Command1(0).Value = True



End Sub

Private Function Update_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   「品目マスタ/美的品番管理ﾃﾞｰﾀ」登録処理
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
    
    
Dim cnt             As Long
Dim cnt1             As Long
    
    
    If B_ITEM.Count(1) = 0 Then
        Exit Function
    End If



    Update_Proc = True
    
    Call Input_Lock

                                    
    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
                                    
                                    'テーブルリセット
    cnt = 0
    cnt1 = 0
    For Row = 1 To B_ITEM.UpperBound(1)
        
        
        DoEvents
        
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  2012.04.13
        Call Rclr_ITEMREC
            
            
        Call UniCode_Conv(ITEMREC.JGYOBU, B_ITEM(Row, colJGYOBU))
        Call UniCode_Conv(ITEMREC.NAIGAI, B_ITEM(Row, colNAIGAI))
        Call UniCode_Conv(ITEMREC.HIN_GAI, B_ITEM(Row, colHIN_GAI))
        Call UniCode_Conv(ITEMREC.HIN_NAI, B_ITEM(Row, colHIN_NAI))
        Call UniCode_Conv(ITEMREC.HIN_NAME, B_ITEM(Row, colHIN_NAME))
            
        Call UniCode_Conv(ITEMREC.ST_SOKO, "**")
        Call UniCode_Conv(ITEMREC.ST_RETU, "**")
        Call UniCode_Conv(ITEMREC.ST_REN, "**")
        Call UniCode_Conv(ITEMREC.ST_DAN, "**")

        Call UniCode_Conv(ITEMREC.INS_TANTO, App.EXEName)
        Call UniCode_Conv(ITEMREC.Ins_DateTime, INS_NOW)
                        
        Do
            sts = BTRV(BtOpInsert, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                
                
                
                Case BtErrDuplicates
                    cnt = cnt + 1
                    
                    
                    Call LOG_OUT(LOG_F, "ITEM " & Format(cnt, "000") & " " & B_ITEM(Row, colHIN_GAI))
                    
                    
                    
                    
                    
                    
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
            
            
        If Trim(B_ITEM(Row, colB_HIN_CODE)) <> "" Then
            Call UniCode_Conv(B_ITEMREC.JGYOBU, B_ITEM(Row, colJGYOBU))
            Call UniCode_Conv(B_ITEMREC.NAIGAI, B_ITEM(Row, colNAIGAI))
            Call UniCode_Conv(B_ITEMREC.HIN_GAI, B_ITEM(Row, colHIN_GAI))
            Call UniCode_Conv(B_ITEMREC.B_HIN_CODE, B_ITEM(Row, colB_HIN_CODE))
                
            Call UniCode_Conv(B_ITEMREC.FILLER, "")
                
            Call UniCode_Conv(B_ITEMREC.INS_TANTO, App.EXEName)
            Call UniCode_Conv(B_ITEMREC.Ins_DateTime, INS_NOW)
                
            Call UniCode_Conv(B_ITEMREC.UPD_TANTO, "")
            Call UniCode_Conv(B_ITEMREC.UPD_DATETIME, "")
                
                
                
            Do
                sts = BTRV(BtOpInsert, B_ITEM_POS, B_ITEMREC, Len(B_ITEMREC), K0_B_ITEM, Len(K0_B_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    
                    
                    
                    Case BtErrDuplicates
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        
                        Beep
                        ans = MsgBox("「美的品番管理ﾃﾞｰﾀ」他端末でデータ使用中です。<ITEM.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Call Input_UnLock
                            Exit Function
                        End If
                    
                    Case Else
                        Call Input_UnLock
                        Call File_Error(sts, BtOpInsert, "美的品番管理ﾃﾞｰﾀ")
                        Exit Function
                End Select
                    
            Loop
        Else
            cnt1 = cnt1 + 1
            
            
            Call LOG_OUT(LOG_F, "B_ITEM " & Format(cnt1, "000") & " " & B_ITEM(Row, colHIN_GAI))
        End If
    
    
    
        Set TDBGrid1.Array = B_ITEM
        TDBGrid1.ReBind
        
        TDBGrid1.Update
        TDBGrid1.Bookmark = Row
        
    Next Row

    MsgBox "「美的品番」登録処理が終了しました。"

    Set TDBGrid1.Array = B_ITEM
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

Private Function List_Disp_B_ITEM_Proc() As Integer
'----------------------------------------------------------------------------
'                   「美的品番データ（EXCEL）」読込み処理
'----------------------------------------------------------------------------
Dim sts                 As Integer
Dim ans                 As Integer
    
Dim INS_NOW             As String * 14
    
    
    

Dim xlApp               As Object
Dim xlBook              As Object
Dim xlSheet             As Object

Dim Row                 As Long

Dim END_GYO             As Integer
Dim SKIP_F              As Boolean

Dim wkJGYOBU            As String * 1
Dim wkNAIGAI            As String * 1
Dim wkHIN_GAI           As String * 20
Dim wkHIN_NAI           As String * 20
Dim wkHIN_NAME          As String * 40
Dim wkB_HIN_CODE        As String * 70

Dim wkB_HIN_CODE_TBL    As Variant



Dim i                   As Long
            
Dim j                   As Long

    List_Disp_B_ITEM_Proc = True

    Call Input_Lock







    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlApp.Worksheets(1)
    
    
    On Error GoTo Error_Proc

    '2011.12.03
    xlApp.Workbooks.Open (Text1.Text), ReadOnly:=True, UpdateLinks:=0

    On Error GoTo 0




    Set B_ITEM = Nothing
    
    Row = Min_Row - 1
    lblDisp_Count.Caption = ""





    
    END_GYO = 0
    For i = 1 + TITLE_GYO To 1048576
        
        
'        On Error GoTo Error_Proc        '2011.12.03

            
        If Trim(xlSheet.Application.Cells(i, 2)) = "" And _
            Trim(xlSheet.Application.Cells(i, 3)) = "" Then
        
            SKIP_F = True
            END_GYO = END_GYO + 1
            
            If END_GYO > 5 Then
                Exit For
            End If
        Else
            
            
            
            END_GYO = 0
        
            If Trim(xlSheet.Application.Cells(i, 2)) = "" Or _
                Trim(xlSheet.Application.Cells(i, 3)) = "" Then
        
            
            Else
                    
                    
                '事業部
                wkJGYOBU = SET_JGYOBU
                '内外
                wkNAIGAI = SET_NAIGAI
                '外部品番
                wkHIN_GAI = Trim(xlSheet.Application.Cells(i, 2))
                '内部品番
                wkHIN_NAI = Trim(xlSheet.Application.Cells(i, 3))
                '美的品番
                
                If Trim(xlSheet.Application.Cells(i, 4)) = "" Then
                    wkB_HIN_CODE_TBL = Split(" ", Chr(&HA), -1)
                Else
                    wkB_HIN_CODE_TBL = Split(Trim(xlSheet.Application.Cells(i, 4)), Chr(&HA), -1)
                End If
                wkB_HIN_CODE = ""
                For j = 0 To UBound(wkB_HIN_CODE_TBL)
                    wkB_HIN_CODE = Trim(wkB_HIN_CODE) & wkB_HIN_CODE_TBL(j)
                Next j
                '品名
                wkHIN_NAME = Trim(xlSheet.Application.Cells(i, 5))
                
                
            
                Row = Row + 1
                B_ITEM.ReDim Min_Row, Row, Min_Col, Max_Col
                
                B_ITEM(Row, colNO) = Row
            
            
                B_ITEM(Row, colJGYOBU) = wkJGYOBU
                B_ITEM(Row, colNAIGAI) = wkNAIGAI
            
                B_ITEM(Row, colHIN_GAI) = wkHIN_GAI
                B_ITEM(Row, colHIN_NAI) = wkHIN_NAI
                
                B_ITEM(Row, colB_HIN_CODE) = wkB_HIN_CODE
            
                B_ITEM(Row, colHIN_NAME) = wkHIN_NAME
            

    
            End If
        End If
    
    
    
    Next i

    On Error GoTo 0



DISP:

    Set TDBGrid1.Array = B_ITEM
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst






    lblDisp_Count.Caption = Format(Row, "#0") & "件"











    Call Input_UnLock

    xlApp.DisplayAlerts = False
    xlBook.Close False
    xlApp.Quit 'EXCELを閉じる
    Set xlApp = Nothing

    List_Disp_B_ITEM_Proc = False
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
            
            
            MsgBox Err.Description & "(" & Err.Number & ")"
            MsgBox "指定のファイルが見つかりません。" & Chr(13) & Chr(10) & "正しいファイル名を入力してください。"
            
            
            
            List_Disp_B_ITEM_Proc = False      '


        Case 13
        
            MsgBox Err.Description & "(" & Err.Number & ")"
            MsgBox "読込み対象のEXCELデータに異常が有ります。内容を確認後、再実行してください。"
            
           GoTo DISP
            
            
            xlApp.DisplayAlerts = False
            xlBook.Close False
            xlApp.Quit 'EXCELを閉じる
            Set xlApp = Nothing
            
            
            
            
            
            
            List_Disp_B_ITEM_Proc = False      '
            

        Case 9
            GoTo DISP

        Case Else
            MsgBox Err.Description & "(" & Err.Number & ")"
    
    End Select
    
    On Error GoTo 0
    
    
    Call Input_UnLock

End Function




Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    B_ITEM_MAKE1.MousePointer = vbHourglass

    Call Ctrl_Lock(B_ITEM_MAKE1)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(B_ITEM_MAKE1)


    B_ITEM_MAKE1.MousePointer = vbDefault

End Sub



