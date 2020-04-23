Attribute VB_Name = "ODR2010G"
Option Explicit
'********************************************************************
'*
'*              ＯＤＲ２０１０用　共通変数
'*
'********************************************************************

'Public NAIGAI_CODE()   As String * 1
'Public NAIGAI_NAME()   As String

'Public ODR_KEY_TB()       As String       'Key内容
'Public ODR_QTY_TB()       As String       'Key単位の日別数量

'Public ODR_KEY_TB(1000)       As String       'Key内容
'Public ODR_QTY_TB(1000, 32)   As String       'Key単位の日別数量

Public DIS_ORDR_NO      As String       '親部品　注文№
Public DIS_OYA_ITEM     As String       '親部品コード
Public DIS_ORDR_QTY     As String       '注文数量
Public DIS_SUM_QTY      As String       '合計数
Public DIS_QTY(1 To 31) As String       '０１
Public DIS_KEY          As String       '

