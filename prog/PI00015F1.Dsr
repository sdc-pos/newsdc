VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PI00015F1 
   ClientHeight    =   15615
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   28560
   StartUpPosition =   3  'Windows の既定値
   _ExtentX        =   50377
   _ExtentY        =   27543
   SectionData     =   "PI00015F1.dsx":0000
End
Attribute VB_Name = "PI00015F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Doukon_com      As Integer      '構成／同梱のBtrieve Operation
Private Doukon_eof      As Integer      '構成／同梱 Eof

Private Doukon_cnt      As Integer      '構成／同梱のLINE COUNT


Private SHIJI_QTY       As Double       '今回指示数
Private Function fStrCut(ByRef CutTxt As String, _
                         ByVal CutLen As Long) As String
'半角・全角の混在する文字列を半角換算文字長で取り出し
    Dim myLen As Long, SysCodeTxt As String
    SysCodeTxt = StrConv(CutTxt, vbFromUnicode)     '文字列を変換
    myLen = LenB(SysCodeTxt)    '半角換算のバイト数を取得
    If myLen <= CutLen Then     '指定の長さより短い場合
        fStrCut = CutTxt & Space$(CutLen - myLen)   '足りない分はスペースで
    Else    '該当の文字列の方が長い場合、指定のバイトでカットする
        fStrCut = StrConv(LeftB$(SysCodeTxt, CutLen), vbUnicode)
        If InStr(fStrCut, vbNullChar) > 0 Then
            '漢字１バイト目で分断された場合の処理
            fStrCut = Left$(fStrCut, InStr(fStrCut, vbNullChar) - 1) & " "
        End If
    End If
End Function


Private Sub ActiveReport_DataInitialize()
    
    Me.Fields.Add "KO_NO"               'No
    Me.Fields.Add "KO_HIN_GAI"          '品番
    Me.Fields.Add "KO_SYUBETSU"         '種別
    Me.Fields.Add "KO_QTY"              '員数
    Me.Fields.Add "KO_SHIJI_QTY"        '数量

    Me.Fields.Add "KO_ST_LOCATION"      '棚番
    Me.Fields.Add "KO_ZAIKO_QTY"        '理論在庫
    Me.Fields.Add "KO_ID_NO"            'ID_NO
    Me.Fields.Add "KO_ID_BCR"           'ID_NOﾊﾞｰｺｰﾄﾞ
    Me.Fields.Add "KO_BIKOU"            '備考
    Me.Fields.Add "KO_HIN_NAME"         '品名



End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer
    
Dim Sumi_Qty    As Long
Dim Mi_Qty      As Long
    
Dim SURYO       As String

Dim ST_SOKO     As String
Dim c           As String * 128
    
Dim wkJgyobu    As String * 1
    
Dim ST_LOCATION As String * 8   '2013.03.31
    
Dim IN_WORD     As String       '2017.12.15
Dim OUT_WORD    As String       '2017.12.15
    


'    If Doukon_cnt > 19 Then                '2013.11.21
'        Exit Sub                           '2013.11.21
'    End If                                 '2013.11.21
    
    
    If Doukon_cnt > 19 Then                     '2013.11.21
        If Doukon_eof Then                      '2013.11.21
            Exit Sub                            '2013.11.21
        Else                                    '2013.11.21
            Doukon_cnt = 0                      '2013.11.21
        End If                                  '2013.11.21
    End If                                      '2013.11.21
    
    
    
    
    
    
    Me.Fields("ko_no").Value = Doukon_Tbl_No(Doukon_cnt)
    
    If Doukon_eof Then
        Me.Fields("KO_HIN_GAI") = ""        '品番
        Me.Fields("KO_SYUBETSU") = ""       '種別
        Me.Fields("KO_QTY") = ""            '員数
        Me.Fields("KO_SHIJI_QTY") = ""      '数量
        Me.Fields("KO_ST_LOCATION") = ""    '棚番
        Me.Fields("KO_ZAIKO_QTY") = ""      '理論在庫
        Me.Fields("KO_ID_NO") = ""          'ID_NO
        Me.Fields("KO_ID_BCR") = ""         'ID_NOﾊﾞｰｺｰﾄﾞ
        Me.Fields("KO_BIKOU") = ""          '備考
    
    
    Else
'--------------------------------------------------- 大阪  部材対応　2012.03.18
'        sts = BTRV(Doukon_com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
        sts = BTRV(Doukon_com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K3_P_SSHIJI_K, Len(K3_P_SSHIJI_K), 3)
'--------------------------------------------------- 大阪  部材対応　2012.03.18
        Select Case sts
            Case BtNoErr
                If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> Taget_Key Or _
                    StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode) <> P_DOUKON Then
                    Doukon_eof = True
                End If
            
            
                If Doukon_cnt = 0 Then              '2016.01.14
                    If Doukon_eof Then              '2016.01.14
                        Doukon_cnt = Doukon_cnt + 1 '2016.01.14
'                        eof = False                '2016.01.14
                        Exit Sub                   '2016.01.14
                    End If                          '2016.01.14
                End If                              '2016.01.14
            
            
            
            
            Case BtErrEOF
                
                Doukon_eof = True
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "対象指図票ﾃﾞｰﾀ（親）")
                Exit Sub
        
        End Select
                                            
        If Doukon_eof Then
            Me.Fields("KO_HIN_GAI") = ""        '品番
            Me.Fields("KO_SYUBETSU") = ""       '種別
            Me.Fields("KO_QTY") = ""            '員数
            Me.Fields("KO_SHIJI_QTY") = ""      '数量
            Me.Fields("KO_ST_LOCATION") = ""    '棚番
            Me.Fields("KO_ZAIKO_QTY") = ""      '理論在庫
            Me.Fields("KO_ID_NO") = ""          'ID_NO
            Me.Fields("KO_ID_BCR") = ""         'ID_NOﾊﾞｰｺｰﾄﾞ
            Me.Fields("KO_BIKOU") = ""          '備考
            Me.Fields("KO_HIN_NAME") = ""       '品名
                                            
                                            
                                            
        Else
                                                '品番
            Me.Fields("KO_HIN_GAI") = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                                                '種別
            Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN06_CD)
            Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SSHIJI_K_REC.KO_SYUBETSU, vbUnicode))
            sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    Call UniCode_Conv(P_CODEREC.C_RNAME, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                    Exit Sub
            
            End Select
'            Me.Fields("KO_SYUBETSU") = StrConv(P_CODEREC.C_RNAME, vbUnicode)               '2017.12.15
            Me.Fields("KO_SYUBETSU") = fStrCut(StrConv(P_CODEREC.C_RNAME, vbUnicode), 6)    '2017.12.15
                                                
                                                
                                                '員数
            If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                Me.Fields("KO_QTY") = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
            Else
                Me.Fields("KO_QTY") = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
            End If
                                                '数量
'            If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                Me.Fields("KO_SHIJI_QTY") = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'            Else
'                Me.Fields("KO_SHIJI_QTY") = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'            End If
        
            SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
            If CInt(Right(SURYO, 2)) = 0 Then
                Me.Fields("KO_SHIJI_QTY") = Format(CDbl(SURYO), "#0")
            Else
                Me.Fields("KO_SHIJI_QTY") = Format(CDbl(SURYO), "#0.00")
            End If
        
        
        
            '品目マスタ読み込み
            
            
'>>>>>  2016.01.27 読み替え廃止
'            If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) = SHIZAI Then                           '2013.03.31
'                Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)                                            '2013.03.31
'            Else                                                                                    '2013.03.31
'                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
'            End If                                                                                  '2013.03.31
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
'>>>>>  2016.01.27 読み替え廃止
            Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))
            Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
            sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
            Select Case sts
                Case BtNoErr
                Case BtErrKeyNotFound
                    
                    
                    Call UniCode_Conv(ITEMREC.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
                    Call UniCode_Conv(ITEMREC.NAIGAI, StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))
                    Call UniCode_Conv(ITEMREC.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
                    
                    
                    Call UniCode_Conv(ITEMREC.HIN_NAME, "")     '2008.02.27
                    
                    
                    Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                    Call UniCode_Conv(ITEMREC.ST_RETU, "")
                    Call UniCode_Conv(ITEMREC.ST_REN, "")
                    Call UniCode_Conv(ITEMREC.ST_DAN, "")
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                    Exit Sub
    
            End Select
        
        
        
            If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                Me.Fields("KO_ST_LOCATION") = ""
            Else
                '標準棚番
                
                ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
'                If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then         '2016.01.13
                If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then      '2016.01.13
                Else
                    ST_SOKO = RTrim(c)
                End If
                
                
                
                Me.Fields("KO_ST_LOCATION") = Trim(ST_SOKO) & "-" & _
                                StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                StrConv(ITEMREC.ST_DAN, vbUnicode)
            End If
        
'--------------------------------------------------- 大阪  部材対応　2012.03.18
            '在庫数
            
            
            
            
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.03.31
'            If StrConv(ITEMREC.JGYOBU, vbUnicode) = SHIZAI Then
'                wkJgyobu = BUZAI
'            Else
'                'wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)  2012.04.04
'                wkJgyobu = YUKO_JGYOBU                          '2012.04.04
'            End If
            
 '>>>>>>>   読み替え廃止 2016.01.27
'            Select Case StrConv(ITEMREC.JGYOBU, vbUnicode)
'                Case SHIZAI
'                    wkJgyobu = BUZAI
'                Case SETSUBI
'                    wkJgyobu = YUKO_JGYOBU
'                Case Else
'                    wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
'            End Select
            wkJgyobu = StrConv(ITEMREC.JGYOBU, vbUnicode)
 '>>>>>>>   読み替え廃止 2016.01.27
            
            
            
'            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
'                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
'                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), , , , Jyogai_Soko_umu) Then
                
                
            ST_LOCATION = StrConv(ITEMREC.ST_SOKO, vbUnicode) & _
                                StrConv(ITEMREC.ST_RETU, vbUnicode) & _
                                StrConv(ITEMREC.ST_REN, vbUnicode) & _
                                StrConv(ITEMREC.ST_DAN, vbUnicode)
            If Zaiko_Syukei_Proc(Sumi_Qty, Mi_Qty, wkJgyobu, _
                                                    StrConv(ITEMREC.NAIGAI, vbUnicode), _
                                                    StrConv(ITEMREC.HIN_GAI, vbUnicode), ST_LOCATION, , , Jyogai_Soko_umu) Then
'>>>>>>>>>>>>>>>>>>>>>>>>>> 2013.03.31
                
                
                
                Exit Sub
            
            End If
'--------------------------------------------------- 大阪  部材対応　2012.03.18
            Me.Fields("KO_ZAIKO_QTY") = Format(Sumi_Qty + Mi_Qty, "#0")
            '備考OR出荷ﾊﾞｰｺｰド
            
            
        
            Select Case PRI_BIKOU_BCR
                Case 0          '備考
                    Me.Fields("KO_BIKOU") = Trim(StrConv(P_SSHIJI_K_REC.KO_BIKOU, vbUnicode))
            
                Case 1          'ID_NOﾊﾞｰｺｰﾄﾞ
            
                    If Trim(StrConv(P_SSHIJI_K_REC.KO_ID_NO, vbUnicode)) = "" Then
                                                                                            'ID_NO
                        Me.Fields("KO_ID_NO") = ""
                                                                                            'ID_NOﾊﾞｰｺｰﾄﾞ
                        Me.Fields("KO_ID_BCR") = ""
                    Else
                                                                                            'ID_NO
                        Me.Fields("KO_ID_NO") = StrConv(ITEMREC.JGYOBU, vbUnicode) & StrConv(P_SSHIJI_K_REC.KO_ID_NO, vbUnicode)
                                                                                                'ID_NOﾊﾞｰｺｰﾄﾞ
                        Me.Fields("KO_ID_BCR") = "*" & StrConv(ITEMREC.JGYOBU, vbUnicode) & StrConv(P_SSHIJI_K_REC.KO_ID_NO, vbUnicode) & "*"
                    End If
            
                Case 2
                    Me.Fields("KO_HIN_NAME") = Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode))
                
            End Select
        End If
            
    
    
    
    
    
    
        Doukon_com = BtOpGetNext
    End If
    
    
    
    Doukon_cnt = Doukon_cnt + 1
    
            
    eof = False
    
    
    


End Sub

Private Sub ActiveReport_Initialize()

Dim sts             As Integer

Dim cnt             As Integer
Dim com             As Integer


Dim i               As Integer
Dim Total_Times     As Double
Dim AVE             As Double


Dim SURYO           As String

Dim ST_SOKO         As String
Dim c               As String * 128

Dim Target          As Double

Dim wkValue         As String
Dim wkEDIT_NIN      As String
Dim wkEDIT_TIMES    As String
Dim wkAVE           As String



    '対象指図票ﾃﾞｰﾀ（親）の読み込み
    Call UniCode_Conv(K0_P_SSHIJI_O.SHIJI_No, Taget_Key)
    sts = BTRV(BtOpGetEqual, P_SSHIJI_O_POS, P_SSHIJI_O_REC, Len(P_SSHIJI_O_REC), K0_P_SSHIJI_O, Len(K0_P_SSHIJI_O), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "指図票ﾃﾞｰﾀ（親）")
            Exit Sub
    
    End Select

    '仕向け先名
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN04_CD)
    Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_CODEREC.C_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "コードマスタ")
            Exit Sub
    
    End Select
       
    Field1.text = Trim(StrConv(P_CODEREC.C_RNAME, vbUnicode))             '仕向け先名
    
    If CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)) = 0 Then
        Field2.text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode)   '指図票№
    Else
        Field2.text = StrConv(P_SSHIJI_O_REC.SHIJI_No, vbUnicode) & "-" & _
                        Format(CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)) + 1, "#")
    End If
    Field3.text = Format(Now, "YYYY/MM/DD HH:MM")                   '発行日時

'    Field3.Text = Mid(StrConv(P_SSHIJI_O_REC.PRINT_DATETIME, vbUnicode), 1, 4) & "/" & _
'                    Mid(StrConv(P_SSHIJI_O_REC.PRINT_DATETIME, vbUnicode), 5, 2) & "/" & _
'                    Mid(StrConv(P_SSHIJI_O_REC.PRINT_DATETIME, vbUnicode), 7, 2) & " " & _
'                    Mid(StrConv(P_SSHIJI_O_REC.PRINT_DATETIME, vbUnicode), 9, 2) & ":" & _
'                    Mid(StrConv(P_SSHIJI_O_REC.PRINT_DATETIME, vbUnicode), 11, 2)

    '承認者
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(P_SSHIJI_O_REC.SHONIN_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Sub
    
    End Select
    Field4.text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)             '承認者
    
    '担当者
    Call UniCode_Conv(K0_TANTO.TANTO_CODE, StrConv(P_SSHIJI_O_REC.TANTO_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(TANTOREC.TANTO_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
            Exit Sub
    
    End Select
    Field5.text = StrConv(TANTOREC.TANTO_NAME, vbUnicode)             '担当者
'--------------------------------------------------- 大阪  部材対応　2012.03.18
    If Trim(StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode)) = "" Then
        Field61.text = "注文なし"
    Else
        Field61.text = "注文№:" & Trim(StrConv(P_SSHIJI_O_REC.ORDER_DT, vbUnicode)) & Trim(StrConv(P_SSHIJI_O_REC.ORDER_DT_SEQ, vbUnicode))
    End If
'--------------------------------------------------- 大阪  部材対応　2012.03.18
    
    '収単／担当者
    lblS_Tanto1.Visible = PRI_S_TANTO
    fldS_Tanto.Visible = PRI_S_TANTO
    speS_tanto1.Visible = PRI_S_TANTO
    l_S_Tanto1.Visible = PRI_S_TANTO
    If PRI_S_TANTO Then
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN05_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SSHIJI_O_REC.S_TANTO, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case BtErrKeyNotFound
                Call UniCode_Conv(P_CODEREC.C_RNAME, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "コードマスタ")
                Exit Sub
        
        End Select
        fldS_Tanto.text = StrConv(P_CODEREC.C_RNAME, vbUnicode)         '収単／担当者
    End If
    
    
    Select Case StrConv(P_SSHIJI_O_REC.SHIJI_F, vbUnicode)              '2007.11.08 指示形態
        Case P_SHIJI_F_NORMAL           '事前
            lblSHIJI_F.Caption = " 事　前 "
        Case P_SHIJI_F_SPOT             'ｽﾎﾟｯﾄ
            lblSHIJI_F.Caption = "スポット"
        Case P_SHIJI_F_KEPPIN           '欠品解除
            lblSHIJI_F.Caption = "欠品解除"
        Case P_SHIJI_F_SAIKON           '再梱包 2007.11.09
            lblSHIJI_F.Caption = "再梱包"
        Case Else
            lblSHIJI_F.Caption = ""
    End Select
    
    
    
    
    
    
    
    Field7.text = StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode)            '品番
                                                                        '数量
    SHIJI_QTY = CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)) - CLng(StrConv(P_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode))
    Field8.text = Format(SHIJI_QTY, "#0")
    '品名／棚番
    Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
            Call UniCode_Conv(ITEMREC.ST_SOKO, "")
            Call UniCode_Conv(ITEMREC.ST_RETU, "")
            Call UniCode_Conv(ITEMREC.ST_REN, "")
            Call UniCode_Conv(ITEMREC.ST_DAN, "")
        
            Call UniCode_Conv(ITEMREC.G_LABEL_NON, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Exit Sub
    
    End Select
    Field9.text = StrConv(ITEMREC.HIN_NAME, vbUnicode)                      '品名

    If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
        Field10.text = ""                                                   '標準棚番
    Else
        ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
'        If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then         '2016.01.13
        If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then      '2016.01.13
        Else
            ST_SOKO = RTrim(c)
        End If
        
        
        Field10.text = Trim(ST_SOKO) & "-" & _
                        StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                        StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                        StrConv(ITEMREC.ST_DAN, vbUnicode)
    End If

    Field11.text = Trim(StrConv(P_SSHIJI_O_REC.S_CLASS_CODE, vbUnicode))    '商品化ｸﾗｽ
    Field12.text = Trim(StrConv(P_SSHIJI_O_REC.F_CLASS_CODE, vbUnicode))    '付加ｸﾗｽ
    Field13.text = Trim(StrConv(P_SSHIJI_O_REC.N_CLASS_CODE, vbUnicode))    '内職ｸﾗｽ


    'ラベル貼付計上有無
    If StrConv(ITEMREC.G_LABEL_NON, vbUnicode) = P_G_LABEL_OFF Then
        lblLabel_NIN.Caption = "******"
        lblLabel_TIMES.Caption = "******"
    Else
        lblLabel_NIN.Caption = ""
        lblLabel_TIMES.Caption = ""
    End If


    '受払先
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SSHIJI_O_REC.UKEHARAI_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
        
    Select Case sts
        Case BtNoErr
        
        Case BtErrKeyNotFound
            Call UniCode_Conv(P_UKEHARAIREC.UKEHARAI_RNAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先マスタ")
            Exit Sub
    
    End Select
    Field14.text = Trim(StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode))   '商品化手配先
    

    '個装資材のループ
    cnt = 0

    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Taget_Key)
    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, P_KOSOU)
    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")

    com = BtOpGetGreaterEqual

    Do
    
        sts = BTRV(com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> Taget_Key Or _
                    StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode) <> P_KOSOU Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "指図票ﾃﾞｰﾀ（子）")
                Exit Sub
        
        End Select
        '品目マスタ読み込み
        If StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode) = SHIZAI Then                           '2013.03.31
            Call UniCode_Conv(K0_ITEM.JGYOBU, BUZAI)                                            '2013.03.31
        Else                                                                                    '2013.03.31
            Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
        End If                                                                                  '2013.03.31
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Sub

        End Select
    
    
    
    
        cnt = cnt + 1
    
        Select Case cnt
            Case 1
            
                '個装資材№
                Field15.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '個装資材　員数
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field16.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field16.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '個装資材　数量
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field17.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field17.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                
                
                
                SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field17.text = Format(CDbl(SURYO), "#0")
                Else
                    Field17.text = Format(CDbl(SURYO), "#0.00")
                End If
                
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field18.text = ""
                Else
                    
                    
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then         '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then      '2016.01.13
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    
                    Field18.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If

            
            
            
            Case 2
            
                '個装資材№
                Field19.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '個装資材　員数
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field20.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field20.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '個装資材　数量
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field21.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field21.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                
                
                
                SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field21.text = Format(CDbl(SURYO), "#0")
                Else
                    Field21.text = Format(CDbl(SURYO), "#0.00")
                End If
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field22.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then         '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then      '2016.01.13
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field22.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            Case 3
                '個装資材№
                Field23.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '個装資材　員数
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field24.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field24.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '個装資材　数量
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field25.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field25.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                
                
                SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field25.text = Format(CDbl(SURYO), "#0")
                Else
                    Field25.text = Format(CDbl(SURYO), "#0.00")
                End If
                
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field26.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then     '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then  '2016.01.13
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field26.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            
            Case 4
            
                '個装資材№
                Field27.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '個装資材　員数
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field28.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field28.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '個装資材　数量
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field29.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field29.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                
                SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field29.text = Format(CDbl(SURYO), "#0")
                Else
                    Field29.text = Format(CDbl(SURYO), "#0.00")
                End If
                
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field30.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then     '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then  '2016.01.13
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field30.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            Case 5
                '個装資材№
                Field31.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '個装資材　員数
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field32.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field32.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '個装資材　数量
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field33.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field33.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field33.text = Format(CDbl(SURYO), "#0")
                Else
                    Field33.text = Format(CDbl(SURYO), "#0.00")
                End If
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field34.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then         '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then      '2016.01.13
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field34.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
        
        End Select
        com = BtOpGetNext
    
    Loop


    '外装資材のループ
    cnt = 0

    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Taget_Key)
    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, P_GAISOU)
    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")

    com = BtOpGetGreaterEqual

    Do
    
        sts = BTRV(com, P_SSHIJI_K_POS, P_SSHIJI_K_REC, Len(P_SSHIJI_K_REC), K0_P_SSHIJI_K, Len(K0_P_SSHIJI_K), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(P_SSHIJI_K_REC.SHIJI_No, vbUnicode) <> Taget_Key Or _
                    StrConv(P_SSHIJI_K_REC.DATA_KBN, vbUnicode) <> P_GAISOU Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "指図票ﾃﾞｰﾀ（子）")
                Exit Sub
        
        End Select
        '品目マスタ読み込み
        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(P_SSHIJI_K_REC.KO_JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(P_SSHIJI_K_REC.KO_NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode))
        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrKeyNotFound
                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                Call UniCode_Conv(ITEMREC.ST_SOKO, "")
                Call UniCode_Conv(ITEMREC.ST_RETU, "")
                Call UniCode_Conv(ITEMREC.ST_REN, "")
                Call UniCode_Conv(ITEMREC.ST_DAN, "")
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Sub

        End Select
    
    
    
    
        cnt = cnt + 1
    
        Select Case cnt
            Case 1
            
                '外装資材№
                Field35.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '外装資材　員数
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field36.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field36.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '外装資材　数量
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field37.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field37.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                
                
                
                
                
                SURYO = Format(Int(CDbl(SHIJI_QTY / CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)))), "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field37.text = Format(CDbl(SURYO), "#0")
                Else
                    Field37.text = Format(CDbl(SURYO), "#0.00")
                End If
                
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field38.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then     '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, App.EXEName, c) Then  '2016.01.13
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field38.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If

            
            
            
            Case 2
            
                '外装資材№
                Field39.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '外装資材　員数
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field40.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field40.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '外装資材　数量
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field41.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field41.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                
                SURYO = Format(Int(CDbl(SHIJI_QTY / CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)))), "00000000.00")
'                SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field41.text = Format(CDbl(SURYO), "#0")
                Else
                    Field41.text = Format(CDbl(SURYO), "#0.00")
                End If
                
                
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field42.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then     '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then      '2016.01.13
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field42.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            Case 3
                '外装資材№
                Field43.text = StrConv(P_SSHIJI_K_REC.KO_HIN_GAI, vbUnicode)
                '外装資材　員数
                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode), 2)) = 0 Then
                    Field44.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0")
                Else
                    Field44.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)), "#0.00")
                End If
                '外装資材　数量
'                If CInt(Right(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode), 2)) = 0 Then
'                    Field45.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0")
'                Else
'                    Field45.text = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_SHIJI_QTY, vbUnicode)), "#0.00")
'                End If
                
                
'                SURYO = Format(CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)) * SHIJI_QTY, "00000000.00")
                SURYO = Format(Int(CDbl(SHIJI_QTY / CDbl(StrConv(P_SSHIJI_K_REC.KO_QTY, vbUnicode)))), "00000000.00")
                If CInt(Right(SURYO, 2)) = 0 Then
                    Field45.text = Format(CDbl(SURYO), "#0")
                Else
                    Field45.text = Format(CDbl(SURYO), "#0.00")
                End If
                
                '標準棚番
                If Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode)) = "" Then
                    Field46.text = ""
                Else
                    ST_SOKO = StrConv(ITEMREC.ST_SOKO, vbUnicode)
'                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then     '2016.01.13
                    If GetIni(StrConv(App.EXEName, vbUpperCase), ST_SOKO, "P_SYS", c) Then      '2016.01.13
                    Else
                        ST_SOKO = RTrim(c)
                    End If
                    Field46.text = ST_SOKO & "-" & _
                                    StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                    StrConv(ITEMREC.ST_DAN, vbUnicode)
                End If
            
            
        
        End Select
    
        com = BtOpGetNext
    
    Loop

    Field47.text = Trim(StrConv(P_SSHIJI_O_REC.BIKOU, vbUnicode))               '備考
    
    
    
    
    
    '見本作成の印字有無
    If StrConv(P_SSHIJI_O_REC.SAMPLE_F, vbUnicode) = P_SAMPLE_F_OFF Then        '見本作成
        lblSample.Visible = False
        Shape10.Visible = False
    Else
        lblSample.Visible = True
        Shape10.Visible = True
    End If

    
    'ﾒｲﾝﾊﾞｰｺｰﾄﾞ
    fldMain_Bcr.Visible = PRI_MAIN_BCR
    If PRI_MAIN_BCR Then
        fldMain_Bcr.text = "*" & Trim(Field2.text) & "*"
    End If

    
    '明細備考
    Select Case PRI_BIKOU_BCR
        Case 0
            fldBIKOU.Visible = True
            
            fldSyuka_No.Visible = False
            fldSyuka_Bcr.Visible = False
            fldHin_Name.Visible = False

        Case 1
        
            fldSyuka_No.Visible = True
            fldSyuka_Bcr.Visible = True

            fldBIKOU.Visible = False
            fldHin_Name.Visible = False

        Case 2
            
            fldHin_Name.Visible = True
        
            fldSyuka_No.Visible = False
            fldSyuka_Bcr.Visible = False

            fldBIKOU.Visible = False

        Case Else
            fldHin_Name.Visible = False
        
            fldSyuka_No.Visible = False
            fldSyuka_Bcr.Visible = False

            fldBIKOU.Visible = False
    End Select

    '同梱部品
    lblDOUKON.Visible = PRI_DOUKON
    lblDOUKON_GOUHI.Visible = PRI_DOUKON


'--------------------------------------------------- 大阪  部材対応　2012.03.18
    '構成／同梱
'    Call UniCode_Conv(K0_P_SSHIJI_K.SHIJI_No, Taget_Key)
'    Call UniCode_Conv(K0_P_SSHIJI_K.DATA_KBN, P_DOUKON)
'    Call UniCode_Conv(K0_P_SSHIJI_K.SEQNO, "")


    Call UniCode_Conv(K3_P_SSHIJI_K.SHIJI_No, Taget_Key)
    Call UniCode_Conv(K3_P_SSHIJI_K.DATA_KBN, P_DOUKON)
    Call UniCode_Conv(K3_P_SSHIJI_K.ST_TANABAN, "")

'--------------------------------------------------- 大阪  部材対応　2012.03.18

    '入庫完了印
    l_Nyuko_IN1.Visible = PRI_NYUKO_IN
    l_Nyuko_IN2.Visible = PRI_NYUKO_IN
    l_Nyuko_IN3.Visible = PRI_NYUKO_IN
    l_Nyuko_IN4.Visible = PRI_NYUKO_IN

    lblNyuko_In.Visible = PRI_NYUKO_IN

    '入力完了印
    l_Input_IN1.Visible = PRI_INPUT_IN
    l_Input_IN2.Visible = PRI_INPUT_IN
    l_Input_IN3.Visible = PRI_INPUT_IN
    l_Input_IN4.Visible = PRI_INPUT_IN

    lblInput_In.Visible = PRI_INPUT_IN


    If Not PRI_NYUKO_IN And Not PRI_NYUKO_IN Then
        l_IN_Center.Visible = False
    Else
        l_IN_Center.Visible = True
    End If

    If CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)) = 0 Then
        lblBunnou.Visible = False
    Else
        lblBunnou.Visible = True
    End If

    '自責タイトル
    If CStr(JISEKI_TITLE(0)) = "" Then
    Else
        lblJISEKI_TITLE.Caption = CStr(JISEKI_TITLE(0)) & "/" & CStr(JISEKI_TITLE(1))
        
    End If
    '他責タイトル
    If CStr(TASEKI_TITLE(0)) = "" Then
    Else
        LblTASEKI_TITLE.Caption = CStr(TASEKI_TITLE(0)) & "/" & CStr(TASEKI_TITLE(1))
        
    End If
        
    '前回実績の獲得
    Call UniCode_Conv(K1_wP_SSHIJI_O.KAN_F, P_KAN_ON)   '完了ﾌﾗｸﾞ
                                                        '仕向け先
    Call UniCode_Conv(K1_wP_SSHIJI_O.SHIMUKE_CODE, StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode))
                                                        '事業部
    Call UniCode_Conv(K1_wP_SSHIJI_O.JGYOBU, StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode))
                                                        '国内外
    Call UniCode_Conv(K1_wP_SSHIJI_O.NAIGAI, StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode))
                                                        '品番
    Call UniCode_Conv(K1_wP_SSHIJI_O.HIN_GAI, StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode))
                                                        '完了日
    Call UniCode_Conv(K1_wP_SSHIJI_O.KAN_DT, "zzzzzzzz")
                                                        '指図表№
    Call UniCode_Conv(K1_wP_SSHIJI_O.SHIJI_No, "zzzzzzzz")
    sts = BTRV(BtOpGetLess, wP_SSHIJI_O_POS, wP_SSHIJI_O_REC, Len(wP_SSHIJI_O_REC), K1_wP_SSHIJI_O, Len(K1_wP_SSHIJI_O), 1)
    Select Case sts
        Case BtNoErr
            If StrConv(wP_SSHIJI_O_REC.KAN_F, vbUnicode) <> P_KAN_ON Or _
                StrConv(wP_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) <> StrConv(P_SSHIJI_O_REC.SHIMUKE_CODE, vbUnicode) Or _
                StrConv(wP_SSHIJI_O_REC.JGYOBU, vbUnicode) <> StrConv(P_SSHIJI_O_REC.JGYOBU, vbUnicode) Or _
                StrConv(wP_SSHIJI_O_REC.NAIGAI, vbUnicode) <> StrConv(P_SSHIJI_O_REC.NAIGAI, vbUnicode) Or _
                StrConv(wP_SSHIJI_O_REC.HIN_GAI, vbUnicode) <> StrConv(P_SSHIJI_O_REC.HIN_GAI, vbUnicode) Then
                    lblBEF_JISSEKI.Caption = ""
                    lblBEF_BEFORE1.Caption = ""
                    lblBEF_BEFORE2.Caption = ""
                    lblBEF_BEFORE3.Caption = ""
                    lblBEF_BEFORE4.Caption = ""
                    lblBEF_SAGYO1.Caption = ""
                    lblBEF_SAGYO2.Caption = ""
                    lblBEF_SAGYO3.Caption = ""
                    lblBEF_AFTER1.Caption = ""
                    lblBEF_AFTER2.Caption = ""
                    lblBEF_JISEKI.Caption = ""
                    lblBEF_TASEKI.Caption = ""
            
            Else
                    

                    '作業①
                    
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)) Then
                        lblBEF_SAGYO1.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)) = 0 Then
                            lblBEF_SAGYO1.Caption = ""
                        Else
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(0).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            
                            
                            lblBEF_SAGYO1.Caption = wkEDIT_NIN & "人×" & wkEDIT_TIMES & "分"
                        End If
                    End If
                    '作業②
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)) Then
                        lblBEF_SAGYO2.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)) = 0 Then
                            lblBEF_SAGYO2.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(1).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_SAGYO2.Caption = wkEDIT_NIN & "人×" & wkEDIT_TIMES & "分"
                        End If
                    End If
                    '作業③
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)) Then
                        lblBEF_SAGYO3.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)) = 0 Then
                            lblBEF_SAGYO3.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)), "#0")
                            Else
                                
                                
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(2).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_SAGYO3.Caption = wkEDIT_NIN & "人×" & wkEDIT_TIMES & "分"
                        End If
                    End If
                    '準備①
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).TIMES, vbUnicode)) Then
                        lblBEF_BEFORE1.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(3).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(3).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).NIN, vbUnicode)) = 0 Then
                            lblBEF_BEFORE1.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(3).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_BEFORE1.Caption = wkEDIT_NIN & "人×" & wkEDIT_TIMES & "分"
                        End If
                    End If
                    '準備②
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)) Then
                        lblBEF_BEFORE2.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)) = 0 Then
                            lblBEF_BEFORE2.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(4).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_BEFORE2.Caption = wkEDIT_NIN & "人×" & wkEDIT_TIMES & "分"
                        End If
                    End If
                    '準備③
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)) Then
                        lblBEF_BEFORE3.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)) = 0 Then
                            lblBEF_BEFORE3.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(5).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_BEFORE3.Caption = wkEDIT_NIN & "人×" & wkEDIT_TIMES & "分"
                        End If
                    End If
                    '準備④
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)) Then
                        lblBEF_BEFORE4.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)) = 0 Then
                            lblBEF_BEFORE4.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(6).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_BEFORE4.Caption = wkEDIT_NIN & "人×" & wkEDIT_TIMES & "分"
                        End If
                    End If
                    '後片付け①
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)) Then
                        lblBEF_AFTER1.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)) = 0 Then
                            lblBEF_AFTER1.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(7).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_AFTER1.Caption = wkEDIT_NIN & "人×" & wkEDIT_TIMES & "分"
                        End If
                    End If
                    '後片付け②
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)) Then
                        lblBEF_AFTER2.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)) = 0 Then
                            lblBEF_AFTER2.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)), "#0")
                            Else
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(8).TIMES, vbUnicode)), "#0.00")
                                End If
                            End If
                            
                            lblBEF_AFTER2.Caption = wkEDIT_NIN & "人×" & wkEDIT_TIMES & "分"
                        End If
                    End If
                    
                    '自責
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)) Then
                        lblBEF_JISEKI.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.JISEKI_NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.JISEKI_TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) = 0 Then
                            lblBEF_JISEKI.Caption = ""
                        Else
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)), "#0")
                            Else
                            
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)), "#0.00")
                                End If
                            
                            End If
                            
                            
                            
                            lblBEF_JISEKI.Caption = wkEDIT_NIN & "人×" & _
                                                    wkEDIT_TIMES & "分 " & _
                                                    StrConv(wP_SSHIJI_O_REC.JISEKI_NAME, vbUnicode)
                        End If
                    End If
                    '他責
                    If Not IsNumeric(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) Or _
                        Not IsNumeric(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)) Then
                        lblBEF_TASEKI.Caption = ""
                        Call UniCode_Conv(wP_SSHIJI_O_REC.TASEKI_NIN, "0.0")
                        Call UniCode_Conv(wP_SSHIJI_O_REC.TASEKI_TIMES, "000.00")
                    Else
                        If CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) = 0 Then
                            lblBEF_TASEKI.Caption = ""
                        Else
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)), "#0.0")
                            If Right(wkValue, 1) = "0" Then
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)), "#0")
                            Else
                                wkEDIT_NIN = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)), "#0.0")
                            End If
                            
                            wkValue = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)), "#0.00")
                            If Right(wkValue, 2) = "00" Then
                                wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)), "#0")
                            Else
                                
                                If Right(wkValue, 1) = "0" Then
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)), "#0.0")
                                Else
                                    wkEDIT_TIMES = Format(CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)), "#0.00")
                                End If
                                
                            End If
                            
                            
                            
                            lblBEF_TASEKI.Caption = wkEDIT_NIN & "人×" & _
                                                    wkEDIT_TIMES & "分 " & _
                                                    StrConv(wP_SSHIJI_O_REC.TASEKI_NAME, vbUnicode)
                        End If
                    End If
                    
                    
                    
                    
                    '総計の計算
                    Total_Times = 0
                    For i = 0 To 8
                        Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)) * CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).TIMES, vbUnicode)))
                    Next i
                    
                    Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_NIN, vbUnicode)) * CDbl(StrConv(wP_SSHIJI_O_REC.JISEKI_TIMES, vbUnicode)))
                    Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_NIN, vbUnicode)) * CDbl(StrConv(wP_SSHIJI_O_REC.TASEKI_TIMES, vbUnicode)))
                    
                    If Total_Times = 0 Then
                        AVE = 0
                    Else
                        AVE = Round(CDbl(Total_Times / CDbl(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode))), 1)
                    End If
            
                    wkValue = Format(Total_Times, "#0.00")
                    If Right(wkValue, 2) = "00" Then
                        wkEDIT_TIMES = Format(Total_Times, "#0")
                    Else
                        wkEDIT_TIMES = Format(Total_Times, "#0.00")
                    End If
            
                    lblBEF_JISSEKI.Caption = "前回:" & Mid(StrConv(wP_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(wP_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(wP_SSHIJI_O_REC.HAKKO_DT, vbUnicode), 7, 2) & ":" & _
                                                Format(CLng(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), "#0") & _
                                                "個 " & _
                                                wkEDIT_TIMES & "分(" & Format(AVE, "#0.0") & "分/個)"

                    '目標の計算
                    Total_Times = 0
                    For i = 0 To 2
                        Total_Times = Total_Times + (CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).NIN, vbUnicode)) * CDbl(StrConv(wP_SSHIJI_O_REC.GENKA_TBL(i).TIMES, vbUnicode)))
                    Next i
                    
                    
                    
                    If CLng(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)) = 0 Then
                        AVE = 0
                    Else
                        AVE = Round(Total_Times / CLng(StrConv(wP_SSHIJI_O_REC.UKEIRE_QTY, vbUnicode)), 1)
                    End If
                                        
                    
                    Target = AVE * CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode))
                    lblTarget1.Caption = "今回目標：" & Format(Target, "#0") & "分"
                    
                    wkValue = Format(AVE, "#0.0")
                    If Right(wkValue, 1) = "0" Then
                        wkAVE = Format(AVE, "#0")
                    Else
                        wkAVE = Format(AVE, "#0.0")
                    End If
                    lblTarget2.Caption = wkAVE & "分/個×" & Format(CLng(StrConv(P_SSHIJI_O_REC.SHIJI_QTY, vbUnicode)), "#0") & "個"
                               
            
            
            
            End If
        
        Case BtErrEOF
            lblBEF_JISSEKI.Caption = ""
            lblBEF_BEFORE1.Caption = ""
            lblBEF_BEFORE2.Caption = ""
            lblBEF_BEFORE3.Caption = ""
            lblBEF_BEFORE4.Caption = ""
            lblBEF_SAGYO1.Caption = ""
            lblBEF_SAGYO2.Caption = ""
            lblBEF_SAGYO3.Caption = ""
            lblBEF_AFTER1.Caption = ""
            lblBEF_AFTER2.Caption = ""
            lblBEF_JISEKI.Caption = ""
            lblBEF_TASEKI.Caption = ""
        
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "指図票ﾃﾞｰﾀ（親）")
            Exit Sub

    End Select


    If CInt(StrConv(P_SSHIJI_O_REC.BUNNOU_CNT, vbUnicode)) = 0 Then
        lblBunnou.Visible = False
    
    
        lblTarget1.Visible = True
        lblTarget2.Visible = True
    
    
    
    
    
    Else
        lblBunnou.Visible = True
    
        lblTarget1.Visible = False
        lblTarget2.Visible = False
    
    
    End If

'    Doukon_com = BtOpGetGreater            '2013.03.31
    Doukon_com = BtOpGetGreaterEqual        '2013.03.31
    
    
    
    Doukon_eof = False

    Doukon_cnt = 0



End Sub

Private Sub ActiveReport_ReportStart()
    
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = vbPRORPortrait
        .PaperBin = vbPRBNCassette
    End With
    
    
    
    Me.PageBottomMargin = 10
    Me.PageTopMargin = 10
    Me.PageLeftMargin = 20
    Me.PageRightMargin = 20

    Me.documentName = "商品化指図票："

    DoEvents

End Sub

