VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00025F1 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows の既定値
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "PR00025F1.dsx":0000
End
Attribute VB_Name = "PR00025F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DET_com         As Integer      '明細のBtrieve Operation


'計上年月用添字
Private Const ptxKEIJYO_YM% = 0         '対象年月






Private Sub ActiveReport_DataInitialize()

    Me.Fields.Add "SHIIRE_NAME"         '仕入先名称（ｺｰﾄﾞ含む）
    
    Me.Fields.Add "SHIIRE01"             '資材
    Me.Fields.Add "SHIIRE02"             '工料
    Me.Fields.Add "SHIIRE03"             '家賃
    Me.Fields.Add "SHIIRE04"             'その他
    Me.Fields.Add "SHIIRE05"             '小計
    Me.Fields.Add "SHIIRE06"             '派遣工料
    Me.Fields.Add "SHIIRE07"             '合計
    Me.Fields.Add "SHIIRE08"             '経費仕入
    Me.Fields.Add "SHIIRE09"             '仕入合計
    Me.Fields.Add "SHIIRE10"             '消費税
    Me.Fields.Add "SHIIRE11"             '総合計
    


End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer

Dim TOTAL       As Long
Dim ZEI         As Long
 
Dim YMD         As String * 8
    
    
    sts = BTRV(DET_com, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "資材仕入集計ﾃﾞｰﾀ")
            Exit Sub
    End Select
    '仕入先名称
    Call UniCode_Conv(K0_P_UKEHARAI.UKEHARAI_CODE, StrConv(P_SHSHI_SUM_REC.SHIIRE_CODE, vbUnicode))
    sts = BTRV(BtOpGetEqual, P_UKEHARAI_POS, P_UKEHARAIREC, Len(P_UKEHARAIREC), K0_P_UKEHARAI, Len(K0_P_UKEHARAI), 0)
    Select Case sts
        Case BtNoErr
            Me.Fields("SHIIRE_NAME") = StrConv(P_SHSHI_SUM_REC.SHIIRE_CODE, vbUnicode) & " " & StrConv(P_UKEHARAIREC.UKEHARAI_RNAME, vbUnicode)
        Case BtErrKeyNotFound
            Me.Fields("SHIIRE_NAME") = ""
        Case Else
            Call File_Error(sts, BtOpGetEqual, "受払先ﾏｽﾀ")
            Exit Sub
    End Select

    Me.Fields("SHIIRE01") = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)), "#,##0")
    Me.Fields("SHIIRE02") = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)), "#,##0")
    Me.Fields("SHIIRE03") = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)), "#,##0")
    Me.Fields("SHIIRE04") = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode)), "#,##0")
    
    
        
    TOTAL = CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)) + _
            CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)) + _
            CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)) + _
            CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode))

        
    Me.Fields("SHIIRE05") = Format(TOTAL, "#,##0")

    Me.Fields("SHIIRE06") = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode)), "#,##0")
        
    TOTAL = TOTAL + CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode))

    Me.Fields("SHIIRE07") = Format(TOTAL, "#,##0")
        
    Me.Fields("SHIIRE08") = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode)), "#,##0")

    TOTAL = TOTAL + CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode))

    Me.Fields("SHIIRE09") = Format(TOTAL, "#,##0")
    
'------------------------------------20060301 センター間消費税なし　訂正　岸見
''''''''If StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) = P_TORI_ANOTHER Then
''''''''ZEI = 0
''''''''    Else
''''''''
''''''''
''''''''    YMD = Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 1, 4) & _
''''''''            Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 6, 2) & _
''''''''            StrConv(P_KANRIREC.SHIME_DD, vbUnicode)
''''''''
''''''''
''''''''    If YMD < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
''''''''        ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
''''''''                CDbl(CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 10))
''''''''    Else
''''''''        ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
''''''''                CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
''''''''
''''''''    End If
''''''''    End If


    If StrConv(P_UKEHARAIREC.TORI_KBN, vbUnicode) = P_TORI_ANOTHER Then
        ZEI = 0
    Else
        ZEI = CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, vbUnicode))
    End If

'--------------------------------------------------------------------------
    Me.Fields("SHIIRE10") = Format(ZEI, "#,##0")

    Me.Fields("SHIIRE11") = Format(TOTAL + ZEI, "#,##0")

    
    
    DET_com = BtOpGetNext
    
            
    eof = False
    
    

End Sub

Private Sub ActiveReport_Initialize()

Dim sts         As Integer
Dim SKIP_Flg    As Boolean
 
Dim TOTAL       As Long
Dim ZEI         As Long
 
Dim YMD         As String * 8
 
    Label1.Caption = PR000251.Text1(ptxKEIJYO_YM).Text                  '計上年月
    Label5.Caption = Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode))   'センター名


    '合計ﾚｺｰﾄﾞ（一般仕入分）の読み込み
    Call UniCode_Conv(K0_P_SHSHI_SUM.SHIIRE_CODE, "")
    Call UniCode_Conv(K0_P_SHSHI_SUM.TORI_KBN, P_TORI_GENERAL)

    SKIP_Flg = False

    sts = BTRV(BtOpGetEqual, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
                    
            SKIP_Flg = True
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "資材仕入集計ﾃﾞｰﾀ")
            Exit Sub
    End Select
    

    If SKIP_Flg Then
        SHIIRE_01_01.Text = "0"
        SHIIRE_01_02.Text = "0"
        SHIIRE_01_03.Text = "0"
        SHIIRE_01_04.Text = "0"
        SHIIRE_01_05.Text = "0"
        SHIIRE_01_06.Text = "0"
        SHIIRE_01_07.Text = "0"
        SHIIRE_01_08.Text = "0"
        SHIIRE_01_09.Text = "0"
        SHIIRE_01_10.Text = "0"
        SHIIRE_01_11.Text = "0"
    
    Else
        
        SHIIRE_01_01.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)), "#,##0")
        SHIIRE_01_02.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)), "#,##0")
        SHIIRE_01_03.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)), "#,##0")
        SHIIRE_01_04.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode)), "#,##0")
        
        TOTAL = CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)) + _
                CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)) + _
                CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)) + _
                CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode))

        SHIIRE_01_05.Text = Format(TOTAL, "#,##0")

        SHIIRE_01_06.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode)), "#,##0")
        
        TOTAL = TOTAL + CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode))

        SHIIRE_01_07.Text = Format(TOTAL, "#,##0")


        SHIIRE_01_08.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode)), "#,##0")

        TOTAL = TOTAL + CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode))

        SHIIRE_01_09.Text = Format(TOTAL, "#,##0")


''        YMD = Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 1, 4) & _
''                Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 6, 2) & _
''                StrConv(P_KANRIREC.SHIME_DD, vbUnicode)
''
''
''        If YMD < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
''            ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
''                    CDbl(CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 10))
''        Else
''            ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
''                    CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
''
''        End If

''        SHIIRE_01_10.Text = Format(ZEI, "#,##0")

        ZEI = CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, vbUnicode))

        SHIIRE_01_10.Text = Format(ZEI, "#,##0")


        SHIIRE_01_11.Text = Format(TOTAL + ZEI, "#,##0")
    
        sts = BTRV(BtOpDelete, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "資材仕入集計ﾃﾞｰﾀ")
                Exit Sub
        End Select
    
    
    End If


    '合計ﾚｺｰﾄﾞ（内職分）の読み込み
    Call UniCode_Conv(K0_P_SHSHI_SUM.SHIIRE_CODE, "")
    Call UniCode_Conv(K0_P_SHSHI_SUM.TORI_KBN, P_TORI_NAISYOKU)

    SKIP_Flg = False

    sts = BTRV(BtOpGetEqual, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
                    
            SKIP_Flg = True
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "資材仕入集計ﾃﾞｰﾀ")
            Exit Sub
    End Select
    

    If SKIP_Flg Then
        SHIIRE_02_01.Text = "0"
        SHIIRE_02_02.Text = "0"
        SHIIRE_02_03.Text = "0"
        SHIIRE_02_04.Text = "0"
        SHIIRE_02_05.Text = "0"
        SHIIRE_02_06.Text = "0"
        SHIIRE_02_07.Text = "0"
        SHIIRE_02_08.Text = "0"
        SHIIRE_02_09.Text = "0"
        SHIIRE_02_10.Text = "0"
        SHIIRE_02_11.Text = "0"
    
    Else
        
        SHIIRE_02_01.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)), "#,##0")
        SHIIRE_02_02.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)), "#,##0")
        SHIIRE_02_03.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)), "#,##0")
        SHIIRE_02_04.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode)), "#,##0")
        
        TOTAL = CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)) + _
                CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)) + _
                CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)) + _
                CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode))

        SHIIRE_02_05.Text = Format(TOTAL, "#,##0")

        SHIIRE_02_06.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode)), "#,##0")
        
        TOTAL = TOTAL + CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode))

        SHIIRE_02_07.Text = Format(TOTAL, "#,##0")


        SHIIRE_02_08.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode)), "#,##0")

        TOTAL = TOTAL + CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode))

        SHIIRE_02_09.Text = Format(TOTAL, "#,##0")


YMD = Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 1, 4) & _
        Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 6, 2) & _
        StrConv(P_KANRIREC.SHIME_DD, vbUnicode)


If YMD < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
    ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
            CDbl(CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 10))
Else
    ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
            CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))

End If


        ZEI = CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, vbUnicode))


        SHIIRE_02_10.Text = Format(ZEI, "#,##0")

        SHIIRE_02_11.Text = Format(TOTAL + ZEI, "#,##0")
    
        sts = BTRV(BtOpDelete, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "資材売上集計ﾃﾞｰﾀ")
                Exit Sub
        End Select
    
    
    End If
    '買掛金合計
    SHIIRE_03_01.Text = Format(CLng(SHIIRE_01_01.Text) + CLng(SHIIRE_02_01.Text), "#,##0")
    SHIIRE_03_02.Text = Format(CLng(SHIIRE_01_02.Text) + CLng(SHIIRE_02_02.Text), "#,##0")
    SHIIRE_03_03.Text = Format(CLng(SHIIRE_01_03.Text) + CLng(SHIIRE_02_03.Text), "#,##0")
    SHIIRE_03_04.Text = Format(CLng(SHIIRE_01_04.Text) + CLng(SHIIRE_02_04.Text), "#,##0")
    SHIIRE_03_05.Text = Format(CLng(SHIIRE_01_05.Text) + CLng(SHIIRE_02_05.Text), "#,##0")
    SHIIRE_03_06.Text = Format(CLng(SHIIRE_01_06.Text) + CLng(SHIIRE_02_06.Text), "#,##0")
    SHIIRE_03_07.Text = Format(CLng(SHIIRE_01_07.Text) + CLng(SHIIRE_02_07.Text), "#,##0")
    SHIIRE_03_08.Text = Format(CLng(SHIIRE_01_08.Text) + CLng(SHIIRE_02_08.Text), "#,##0")
    SHIIRE_03_09.Text = Format(CLng(SHIIRE_01_09.Text) + CLng(SHIIRE_02_09.Text), "#,##0")
    SHIIRE_03_10.Text = Format(CLng(SHIIRE_01_10.Text) + CLng(SHIIRE_02_10.Text), "#,##0")
    SHIIRE_03_11.Text = Format(CLng(SHIIRE_01_11.Text) + CLng(SHIIRE_02_11.Text), "#,##0")

    '合計ﾚｺｰﾄﾞ（現金仕入）の読み込み
    Call UniCode_Conv(K0_P_SHSHI_SUM.SHIIRE_CODE, "")
    Call UniCode_Conv(K0_P_SHSHI_SUM.TORI_KBN, P_TORI_GENKIN)

    SKIP_Flg = False

    sts = BTRV(BtOpGetEqual, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
                    
            SKIP_Flg = True
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "資材仕入集計ﾃﾞｰﾀ")
            Exit Sub
    End Select
    

    If SKIP_Flg Then
        SHIIRE_04_01.Text = "0"
        SHIIRE_04_02.Text = "0"
        SHIIRE_04_03.Text = "0"
        SHIIRE_04_04.Text = "0"
        SHIIRE_04_05.Text = "0"
        SHIIRE_04_06.Text = "0"
        SHIIRE_04_07.Text = "0"
        SHIIRE_04_08.Text = "0"
        SHIIRE_04_09.Text = "0"
        SHIIRE_04_10.Text = "0"
        SHIIRE_04_11.Text = "0"
    
    Else
        
        SHIIRE_04_01.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)), "#,##0")
        SHIIRE_04_02.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)), "#,##0")
        SHIIRE_04_03.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)), "#,##0")
        SHIIRE_04_04.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode)), "#,##0")
        
        TOTAL = CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)) + _
                CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)) + _
                CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)) + _
                CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode))

        SHIIRE_04_05.Text = Format(TOTAL, "#,##0")

        SHIIRE_04_06.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode)), "#,##0")
        
        TOTAL = TOTAL + CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode))

        SHIIRE_04_07.Text = Format(TOTAL, "#,##0")


        SHIIRE_04_08.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode)), "#,##0")

        TOTAL = TOTAL + CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode))

        SHIIRE_04_09.Text = Format(TOTAL, "#,##0")


'----------------------20060423
'YMD = Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 1, 4) & _
'        Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 6, 2) & _
'        StrConv(P_KANRIREC.SHIME_DD, vbUnicode)
'
'
'If YMD < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
'    ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
'            CDbl(CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 10))
'Else
'    ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'            CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
'
'End If
'----------------------20060423

        ZEI = CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(6).SHIIRE, vbUnicode))


        SHIIRE_04_10.Text = Format(ZEI, "#,##0")

        SHIIRE_04_11.Text = Format(TOTAL + ZEI, "#,##0")
    
        sts = BTRV(BtOpDelete, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "資材仕入集計ﾃﾞｰﾀ")
                Exit Sub
        End Select
    
    
    End If


    '合計ﾚｺｰﾄﾞ（他ｾﾝﾀｰからの振替）の読み込み
    Call UniCode_Conv(K0_P_SHSHI_SUM.SHIIRE_CODE, "")
    Call UniCode_Conv(K0_P_SHSHI_SUM.TORI_KBN, P_TORI_ANOTHER)

    SKIP_Flg = False

    sts = BTRV(BtOpGetEqual, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
                    
            SKIP_Flg = True
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "資材仕入集計ﾃﾞｰﾀ")
            Exit Sub
    End Select
    

    If SKIP_Flg Then
        SHIIRE_05_01.Text = "0"
        SHIIRE_05_02.Text = "0"
        SHIIRE_05_03.Text = "0"
        SHIIRE_05_04.Text = "0"
        SHIIRE_05_05.Text = "0"
        SHIIRE_05_06.Text = "0"
        SHIIRE_05_07.Text = "0"
        SHIIRE_05_08.Text = "0"
        SHIIRE_05_09.Text = "0"
        SHIIRE_05_10.Text = "0"
        SHIIRE_05_11.Text = "0"
    
    Else
        
        SHIIRE_05_01.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)), "#,##0")
        SHIIRE_05_02.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)), "#,##0")
        SHIIRE_05_03.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)), "#,##0")
        SHIIRE_05_04.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode)), "#,##0")
        
        TOTAL = CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)) + _
                CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)) + _
                CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)) + _
                CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode))

        SHIIRE_05_05.Text = Format(TOTAL, "#,##0")

        SHIIRE_05_06.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode)), "#,##0")
        
        TOTAL = TOTAL + CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode))

        SHIIRE_05_07.Text = Format(TOTAL, "#,##0")


        SHIIRE_05_08.Text = Format(CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode)), "#,##0")

        TOTAL = TOTAL + CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode))

        SHIIRE_05_09.Text = Format(TOTAL, "#,##0")

'----------------------20060301 センター間消費税なし　訂正　岸見

      
'        YMD = Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 1, 4) & _
'                Mid(PR000251.Text1(ptxKEIJYO_YM).Text, 6, 2) & _
'                StrConv(P_KANRIREC.SHIME_DD, vbUnicode)
'
'
'        If YMD < StrConv(P_KANRIREC.ZEI_CHANGE_YMD, vbUnicode) Then
'            ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 100)) + _
'                    CDbl(CDbl(StrConv(P_KANRIREC.NOW_ZEI_RITU, vbUnicode)) / 10))
'        Else
'            ZEI = Int(CDbl(TOTAL * (CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 100)) + _
'                    CDbl(CDbl(StrConv(P_KANRIREC.NEW_ZEI_RITU, vbUnicode)) / 10))
'
'        End If

ZEI = 0

'------------------------------------------------------------------------
        SHIIRE_05_10.Text = Format(ZEI, "#,##0")

        SHIIRE_05_11.Text = Format(TOTAL + ZEI, "#,##0")
    
        sts = BTRV(BtOpDelete, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
        Select Case sts
            Case BtNoErr
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "資材売上集計ﾃﾞｰﾀ")
                Exit Sub
        End Select
    
    
    End If


    '総合計
    SHIIRE_06_01.Text = Format(CLng(SHIIRE_03_01.Text) + CLng(SHIIRE_04_01.Text) + CLng(SHIIRE_05_01.Text), "#,##0")
    SHIIRE_06_02.Text = Format(CLng(SHIIRE_03_02.Text) + CLng(SHIIRE_04_02.Text) + CLng(SHIIRE_05_02.Text), "#,##0")
    SHIIRE_06_03.Text = Format(CLng(SHIIRE_03_03.Text) + CLng(SHIIRE_04_03.Text) + CLng(SHIIRE_05_03.Text), "#,##0")
    SHIIRE_06_04.Text = Format(CLng(SHIIRE_03_04.Text) + CLng(SHIIRE_04_04.Text) + CLng(SHIIRE_05_04.Text), "#,##0")
    SHIIRE_06_05.Text = Format(CLng(SHIIRE_03_05.Text) + CLng(SHIIRE_04_05.Text) + CLng(SHIIRE_05_05.Text), "#,##0")
    SHIIRE_06_06.Text = Format(CLng(SHIIRE_03_06.Text) + CLng(SHIIRE_04_06.Text) + CLng(SHIIRE_05_06.Text), "#,##0")
    SHIIRE_06_07.Text = Format(CLng(SHIIRE_03_07.Text) + CLng(SHIIRE_04_07.Text) + CLng(SHIIRE_05_07.Text), "#,##0")
    SHIIRE_06_08.Text = Format(CLng(SHIIRE_03_08.Text) + CLng(SHIIRE_04_08.Text) + CLng(SHIIRE_05_08.Text), "#,##0")
    SHIIRE_06_09.Text = Format(CLng(SHIIRE_03_09.Text) + CLng(SHIIRE_04_09.Text) + CLng(SHIIRE_05_09.Text), "#,##0")
    SHIIRE_06_10.Text = Format(CLng(SHIIRE_03_10.Text) + CLng(SHIIRE_04_10.Text) + CLng(SHIIRE_05_10.Text), "#,##0")
    SHIIRE_06_11.Text = Format(CLng(SHIIRE_03_11.Text) + CLng(SHIIRE_04_11.Text) + CLng(SHIIRE_05_11.Text), "#,##0")


    DET_com = BtOpGetFirst
    Do
        sts = BTRV(DET_com, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
        Select Case sts
            Case BtNoErr
            
            
                If CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)) = 0 And _
                    CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)) = 0 And _
                    CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)) = 0 And _
                    CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode)) = 0 And _
                    CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode)) = 0 And _
                    CLng(StrConv(P_SHSHI_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode)) = 0 Then
                    sts = BTRV(BtOpDelete, P_SHSHI_SUM_POS, P_SHSHI_SUM_REC, Len(P_SHSHI_SUM_REC), K0_P_SHSHI_SUM, Len(K0_P_SHSHI_SUM), 0)
                    Select Case sts
                        Case BtNoErr
                                    
                            Exit Do
                        
                        Case Else
                            Call File_Error(sts, BtOpDelete, "資材仕入集計ﾃﾞｰﾀ")
                            Exit Sub
                    End Select
                End If
                    
            Case BtErrEOF
                        
                Exit Do
            
            Case Else
                Call File_Error(sts, DET_com, "資材仕入集計ﾃﾞｰﾀ")
                Exit Sub
        End Select
    
        DET_com = BtOpGetNext
    
    Loop






    Call UniCode_Conv(K0_P_SHSHI_SUM.SHIIRE_CODE, "")
    Call UniCode_Conv(K0_P_SHSHI_SUM.TORI_KBN, "")

    DET_com = BtOpGetGreater


End Sub


Private Sub ActiveReport_ReportStart()
    
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = vbPRORLandscape
        .PaperBin = vbPRBNCassette
    End With
    
    
    
    Me.PageBottomMargin = 25
    Me.PageTopMargin = 25
    Me.PageLeftMargin = 0
    Me.PageRightMargin = 0

    Me.documentName = "得意先別仕入集計表："

    DoEvents

End Sub

