Attribute VB_Name = "PR00090com"
Option Explicit

'Glid用環境---------------------------------

Public SSHIJI   As New XArrayDB

Public Const Min_Row% = 1                   '最小行数
Public Const Min_Col% = 0                  '最小列数
Public Const Max_Col% = 6                  '最大列数

Public Const colHAKKO_DT% = 0              '発行日
Public Const colSHIMUKE_CODE% = 1          '仕向け先
Public Const colSHIJI_NO% = 2              '指図票№
Public Const colHIN_GAI% = 3               '品番
Public Const colSHIJI_QTY% = 4             '指示数
Public Const colDOUKON% = 5                '同梱件数
Public Const colKAN_DT% = 6                '完了日

