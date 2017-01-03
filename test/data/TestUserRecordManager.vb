'
' 日付: 2017/01/02
'
Imports NUnit.Framework

Imports System.Data
Imports System.Linq
Imports Common.Account
Imports Common.Extensions
Imports Common.Util

<TestFixture> _
Public Class TestUserRecordManager
  Dim COL_CNT  As String = WorkItemColumnsInfo.WORKCOUNT_COL_NAME
  Dim COL_TIME As String = WorkItemColumnsInfo.WORKTIME_COL_NAME
  Dim COL_PROD As String = WorkItemColumnsInfo.WORKPRODUCTIVITY_COL_NAME
  
  Dim manager As UserRecordManager
  Dim users   As UserInfo()
  
  ''' <summary>
  ''' ユーザ管理オブジェクトを初期化する。
  ''' </summary>
  <TestFixtureSetUp> _
  Public Sub Init
    Dim p As New ExcelProperties("test/ExcelForTest.properties")
    
    manager = New UserRecordManager(p)
    
    Dim b As UserRecordBuffer = manager.Loader().GetUserRecordBuffer()
    
    users = {
      New UserInfo("john",   "001", "xxx"),
      New UserInfo("paul",   "002", "xxx"),
      New UserInfo("george", "003", "xxx")
    }
    
    For Each user As UserInfo In users
      Dim record As UserRecord = b.GetUserRecord(user)
      InsertData(record, Integer.Parse(user.GetSimpleId))
      
      b.PlusToTotalRecord(user)
    Next
  End Sub
  
  ''' <summary>
  ''' データテーブルに仮データをセットする。
  ''' </summary>
  Private Sub InsertData(record As UserRecord, id As Integer)
    For m As Integer = 10 To 12
      Dim t As DataTable = record.GetRecord(m)
      For d As Integer = 1 To DateTime.DaysInMonth(2016, m)
        If d Mod 4 = 0 Then
          t.Rows(d - 1)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME) = d + m + id
          
          Dim time As Double = 1.0
          If d Mod 8 = 0 Then time = 0
          t.Rows(d - 1)("item1" & WorkItemColumnsInfo.WORKTIME_COL_NAME)  = time
        End If
        
        If d Mod 6 = 0 Then
          t.Rows(d - 1)("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME) = d + m + id
          t.Rows(d - 1)("item2" & WorkItemColumnsInfo.WORKTIME_COL_NAME)  = 1
        End If
        
        If d Mod 8 = 0 Then
          t.Rows(d - 1)("item3" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME) = d * m + id
        End If
        
        If d Mod 10 = 0 Then
          t.Rows(d - 1)("item4" & WorkItemColumnsInfo.WORKTIME_COL_NAME)  = m + id
        End If
      Next
    Next
  End Sub
  
  <Test> _
  Public Sub TestGetDailyRecord
    '
    ' 指定した月の１日ごとのデータをまとめたテーブルを取得するテスト
    '
    Dim table1 As DataTable = Me.manager.GetDailyRecord(users(0), 2016, 10, False)
    Assert.AreEqual(32,    table1.Rows.Count)
    Assert.AreEqual(15,    table1.Rows(3).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(15.0,  table1.Rows(3).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(189,   table1.Rows(31).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(47.25, table1.Rows(31).GetOrDefault("item1" & COL_PROD, 0.0))
    
    AssertDate(table1, 0, 30, 1)
    Assert.AreEqual("合計", table1.Rows(31).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    
    ' 作業時間が０の作業件数が合計に含まれてないことをテスト
    Dim table1b As DataTable = Me.manager.GetDailyRecord(users(0), 2016, 10, True)
    Assert.AreEqual(108,  table1b.Rows(31).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(27.0, table1b.Rows(31).GetOrDefault("item1" & COL_PROD, 0.0))
    
    AssertDate(table1b, 0, 30, 1)
    Assert.AreEqual("合計", table1b.Rows(31).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    
    Dim table2 As DataTable = Me.manager.GetDailyRecord(users(1), 2016, 11, False)
    Assert.AreEqual(31,   table2.Rows.Count)
    Assert.AreEqual(19,   table2.Rows(5).GetOrDefault("item2" & COL_CNT,  0))
    Assert.AreEqual(19.0, table2.Rows(5).GetOrDefault("item2" & COL_PROD, 0.0))
    Assert.AreEqual(155,  table2.Rows(30).GetOrDefault("item2" & COL_CNT,  0))
    Assert.AreEqual(31.0, table2.Rows(30).GetOrDefault("item2" & COL_PROD, 0.0))
    
    AssertDate(table2, 0, 29, 1)
    Assert.AreEqual("合計", table2.Rows(30).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
  End Sub
  
  <Test> _
  Public Sub TestGetGetDailyTotalRecord
    '
    ' 指定した期間の集計データを収めたテーブルを取得するテスト
    '
    Dim term   As DateTerm  = CreateTerm(10, 20, 11, 10)
    Dim table1 As DataTable = Me.manager.GetDailyTotalRecord(term, False)
    Assert.AreEqual(23,    table1.Rows.Count)
    Assert.AreEqual(96,    table1.Rows(0).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(32,    table1.Rows(0).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(438,   table1.Rows(22).GetOrDefault("item1" & COL_CNT, 0))
    Assert.AreEqual(48.67, Math.Round(table1.Rows(22).GetOrDefault("item1" & COL_PROD, 0.0), 2))
    
    AssertDate(table1, 0, 11, 20)
    AssertDate(table1, 12, 21, 1)
    Assert.AreEqual("合計", table1.Rows(22).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    
    ' 作業時間が０の作業件数が合計に含まれてないことをテスト
    Dim table2 As DataTable = Me.manager.GetDailyTotalRecord(term, True)
    Assert.AreEqual(23,    table2.Rows.Count)
    Assert.AreEqual(267,   table2.Rows(22).GetOrDefault("item1" & COL_CNT, 0))
    Assert.AreEqual(29.67, Math.Round(table2.Rows(22).GetOrDefault("item1" & COL_PROD, 0.0), 2))
    
    AssertDate(table2, 0, 11, 20)
    AssertDate(table2, 12, 21, 1)
    Assert.AreEqual("合計", table2.Rows(22).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
  End Sub
  
  ' 日付列が順番に入っているかどうか検証する
  Private Sub AssertDate(table As DataTable, startIdx As Integer, endIdx As Integer, startDay As Integer)
    Dim day As Integer = startDay
    For idx As Integer = startIdx To endIdx
      Dim dayStr As String = day.ToString & "日"
      Assert.AreEqual(dayStr, table.Rows(idx).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
      day += 1
    Next
  End Sub
  
  <Test> _
  Public Sub TestGetWeeklyRecord
    '
    ' 指定した期間のデータを１週間ごとに集計したテーブルを取得するテスト
    '
    Dim term   As DateTerm  = CreateTerm(10, 20, 11, 10)
    Dim table1 As DataTable = Me.manager.GetWeeklyRecord(users(1), term, False)
    Assert.AreEqual(5, table1.Rows.Count)
    Assert.AreEqual(32,    table1.Rows(0).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(32.0,  table1.Rows(0).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(76,    table1.Rows(1).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(76.0,  table1.Rows(1).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(17,    table1.Rows(2).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(21,    table1.Rows(3).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(146,   table1.Rows(4).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(48.67, Math.Round(table1.Rows(4).GetOrDefault("item1" & COL_PROD,  0.0), 2))
    
    Assert.AreEqual("10月第4週", table1.Rows(0).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("10月第5週", table1.Rows(1).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("10月第6週/11月第1週", table1.Rows(2).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("11月第2週", table1.Rows(3).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("合計", table1.Rows(4).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    
    ' 作業時間が０の作業件数を集計に含めない
    Dim table1b As DataTable = Me.manager.GetWeeklyRecord(users(1), term, True)
    Assert.AreEqual(5, table1b.Rows.Count)
    Assert.AreEqual(32,    table1b.Rows(0).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(32.0,  table1b.Rows(0).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(40,    table1b.Rows(1).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(40.0,  table1b.Rows(1).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(17,    table1b.Rows(2).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(0,     table1b.Rows(3).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(89,    table1b.Rows(4).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(29.67, Math.Round(table1b.Rows(4).GetOrDefault("item1" & COL_PROD,  0.0), 2))
    
    Assert.AreEqual("10月第4週", table1b.Rows(0).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("10月第5週", table1b.Rows(1).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("10月第6週/11月第1週", table1b.Rows(2).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("11月第2週", table1b.Rows(3).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("合計",      table1b.Rows(4).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
  End Sub
  
  <Test> _
  Public Sub TestGetWeeklyTotalRecord
    Dim term   As DateTerm  = CreateTerm(10, 20, 11, 10)
    Dim table1 As DataTable = Me.manager.GetWeeklyTotalRecord(term, False)
    Assert.AreEqual(5, table1.Rows.Count)
    Assert.AreEqual(96,    table1.Rows(0).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(32.0,  table1.Rows(0).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(228,   table1.Rows(1).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(76.0,  table1.Rows(1).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(51,    table1.Rows(2).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(63,    table1.Rows(3).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(438,   table1.Rows(4).GetOrDefault("item1" & COL_CNT, 0))
    Assert.AreEqual(48.67, Math.Round(table1.Rows(4).GetOrDefault("item1" & COL_PROD, 0.0), 2))
    
    Assert.AreEqual("10月第4週", table1.Rows(0).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("10月第5週", table1.Rows(1).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("10月第6週/11月第1週", table1.Rows(2).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("11月第2週", table1.Rows(3).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("合計",      table1.Rows(4).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    
    Dim table1b As DataTable = Me.manager.GetWeeklyTotalRecord(term, True)
    Assert.AreEqual(5, table1b.Rows.Count)
    Assert.AreEqual(96,    table1b.Rows(0).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(32.0,  table1b.Rows(0).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(120,   table1b.Rows(1).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(40.0,  table1b.Rows(1).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(51,    table1b.Rows(2).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(0,     table1b.Rows(3).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(267,   table1b.Rows(4).GetOrDefault("item1" & COL_CNT, 0))
    Assert.AreEqual(29.67, Math.Round(table1b.Rows(4).GetOrDefault("item1" & COL_PROD, 0.0), 2))
    
    Assert.AreEqual("10月第4週", table1b.Rows(0).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("10月第5週", table1b.Rows(1).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("10月第6週/11月第1週", table1b.Rows(2).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("11月第2週", table1b.Rows(3).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("合計",      table1b.Rows(4).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
  End Sub
  
  <Test> _
  Public Sub TestGetMonthlyRecord
    '
    ' 指定した期間の集計データを月ごとに集計したテーブルを取得するテスト
    '
    Dim term   As DateTerm  = CreateTerm(10, 20, 11, 10)
    Dim table1 As DataTable = Me.manager.GetMonthlyRecord(users(2), term, False)
    Assert.AreEqual(3, table1.Rows.Count)
    Assert.AreEqual(111,   table1.Rows(0).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(55.5,  table1.Rows(0).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(40,    table1.Rows(1).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(40.0,  table1.Rows(1).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(151,   table1.Rows(2).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(50.33, Math.Round(table1.Rows(2).GetOrDefault("item1" & COL_PROD, 0.0), 2))
    
    Assert.AreEqual("10月", table1.Rows(0).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("11月", table1.Rows(1).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("合計", table1.Rows(2).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    
    Dim table1b As DataTable = Me.manager.GetMonthlyRecord(users(2), term, True)
    Assert.AreEqual(3, table1b.Rows.Count)
    Assert.AreEqual(74,    table1b.Rows(0).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(37,    table1b.Rows(0).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(18,    table1b.Rows(1).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(18.0,  table1b.Rows(1).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(92,    table1b.Rows(2).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(30.67, Math.Round(table1b.Rows(2).GetOrDefault("item1" & COL_PROD, 0.0), 2))
    
    Assert.AreEqual("10月", table1b.Rows(0).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("11月", table1b.Rows(1).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("合計", table1b.Rows(2).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
  End Sub
  
  <Test> _
  Public Sub TestGetMonthlyTotalRecord
    '
    ' 指定した期間のデータを月ごとに集計したテーブルを取得するテスト
    '
    Dim term   As DateTerm  = CreateTerm(10, 20, 11, 10)
    Dim table1 As DataTable = Me.manager.GetMonthlyTotalRecord(term, False)
    Assert.AreEqual(3, table1.Rows.Count)
    Assert.AreEqual(324,   table1.Rows(0).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(54,    table1.Rows(0).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(114,   table1.Rows(1).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(38.0,  table1.Rows(1).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(438,   table1.Rows(2).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(48.67, Math.Round(table1.Rows(2).GetOrDefault("item1" & COL_PROD, 0.0), 2))
    
    Assert.AreEqual("10月", table1.Rows(0).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("11月", table1.Rows(1).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("合計", table1.Rows(2).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    
    Dim table1b As DataTable = Me.manager.GetMonthlyTotalRecord(term, True)
    Assert.AreEqual(3, table1b.Rows.Count)
    Assert.AreEqual(216,   table1b.Rows(0).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(36.0,  table1b.Rows(0).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(51,    table1b.Rows(1).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(17.0,  table1b.Rows(1).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(267,   table1b.Rows(2).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(29.67, Math.Round(table1b.Rows(2).GetOrDefault("item1" & COL_PROD, 0.0), 2))  
    
    Assert.AreEqual("10月", table1b.Rows(0).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("11月", table1b.Rows(1).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("合計", table1b.Rows(2).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
  End Sub
  
  <Test> _
  Public Sub TestGetSumRecord
    '
    ' 指定したユーザの集計テーブルを取得するテスト
    '
    Dim table1 As DataTable = Me.manager.GetSumRecord(users(0), False)
    Assert.AreEqual(18,  table1.Rows.Count)
    Assert.AreEqual(0,   table1.Rows(0).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(34,  table1.Rows(1).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(23,  table1.Rows(2).GetOrDefault("item2" & COL_CNT,  0))
    Assert.AreEqual(161, table1.Rows(3).GetOrDefault("item3" & COL_CNT,  0))
    
    Assert.AreEqual(41, table1.Rows(13).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(43, table1.Rows(13).GetOrDefault("item2" & COL_CNT,  0))
    Assert.AreEqual(0,  table1.Rows(13).GetOrDefault("item3" & COL_CNT,  0))
    
    Assert.AreEqual(189,   table1.Rows(14).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(47.25, table1.Rows(14).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(150,   table1.Rows(15).GetOrDefault("item2" & COL_CNT,  0))
    Assert.AreEqual(30.00, table1.Rows(15).GetOrDefault("item2" & COL_PROD, 0.0))
    Assert.AreEqual(579,   table1.Rows(16).GetOrDefault("item3" & COL_CNT,  0))
    
    Assert.AreEqual(588,   table1.Rows(17).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(49.0,  table1.Rows(17).GetOrDefault("item1" & COL_PROD, 0.0))
    
    Assert.AreEqual("10月第1週", table1.Rows(0).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("10月",      table1.Rows(14).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("合計",      table1.Rows(17).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    
    Dim table1b As DataTable = Me.manager.GetSumRecord(users(1), True)
    Assert.AreEqual(18,  table1b.Rows.Count)
    Assert.AreEqual(112,   table1b.Rows(14).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(28.00, table1b.Rows(14).GetOrDefault("item1" & COL_PROD, 0.0))
    Assert.AreEqual(155,   table1b.Rows(15).GetOrDefault("item2" & COL_CNT,  0))
    Assert.AreEqual(31.00, table1b.Rows(15).GetOrDefault("item2" & COL_PROD, 0.0))
    Assert.AreEqual(582,   table1b.Rows(16).GetOrDefault("item3" & COL_CNT,  0))
    
    Assert.AreEqual(348,   table1b.Rows(17).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(29.0,  table1b.Rows(17).GetOrDefault("item1" & COL_PROD, 0.0))
    
    Assert.AreEqual("10月第1週", table1b.Rows(0).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("10月",      table1b.Rows(14).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
    Assert.AreEqual("合計",      table1b.Rows(17).GetOrDefault(UserRecordColumnsInfo.DATE_COL_NAME, ""))
  End Sub
  
  <Test> _
  Public Sub TestGetAllUserSumRecord
    '
    ' 指定した期間のすべてのユーザ集計データを収めたテーブルを取得するテスト
    '
    Dim term   As DateTerm  = CreateTerm(10, 20, 11, 10)
    Dim table1 As DataTable = Me.manager.GetAllUserSumRecord(term, False)
    
    Assert.AreEqual(4, table1.Rows.Count)
    Assert.AreEqual(141,   table1.Rows(0).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(47.00, table1.Rows(0).GetOrDefault("item1" & COL_PROD,  0.0))
    Assert.AreEqual(97 ,   table1.Rows(1).GetOrDefault("item2" & COL_CNT,  0))
    Assert.AreEqual(334,   table1.Rows(2).GetOrDefault("item3" & COL_CNT,  0))
    
    Assert.AreEqual(438,   table1.Rows(3).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(291,   table1.Rows(3).GetOrDefault("item2" & COL_CNT,  0))
    Assert.AreEqual(996,   table1.Rows(3).GetOrDefault("item3" & COL_CNT,  0))
    
    Assert.AreEqual("001 john",   table1.Rows(0).GetOrDefault(UserRecordColumnsInfo.NAME_COL_NAME, ""))
    Assert.AreEqual("002 paul",   table1.Rows(1).GetOrDefault(UserRecordColumnsInfo.NAME_COL_NAME, ""))
    Assert.AreEqual("003 george", table1.Rows(2).GetOrDefault(UserRecordColumnsInfo.NAME_COL_NAME, ""))
    Assert.AreEqual("合計",       table1.Rows(3).GetOrDefault(UserRecordColumnsInfo.NAME_COL_NAME, ""))
    
    Dim table1b As DataTable = Me.manager.GetAllUserSumRecord(term, True)
    
    Assert.AreEqual(4, table1.Rows.Count)
    Assert.AreEqual(86,    table1b.Rows(0).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(28.67, Math.Round(table1b.Rows(0).GetOrDefault("item1" & COL_PROD,  0.0), 2))
    
    Assert.AreEqual(267,   table1b.Rows(3).GetOrDefault("item1" & COL_CNT,  0))
    Assert.AreEqual(291,   table1b.Rows(3).GetOrDefault("item2" & COL_CNT,  0))
    Assert.AreEqual(996,   table1b.Rows(3).GetOrDefault("item3" & COL_CNT,  0))
    
    Assert.AreEqual("001 john",   table1b.Rows(0).GetOrDefault(UserRecordColumnsInfo.NAME_COL_NAME, ""))
    Assert.AreEqual("002 paul",   table1b.Rows(1).GetOrDefault(UserRecordColumnsInfo.NAME_COL_NAME, ""))
    Assert.AreEqual("003 george", table1b.Rows(2).GetOrDefault(UserRecordColumnsInfo.NAME_COL_NAME, ""))
    Assert.AreEqual("合計",       table1b.Rows(3).GetOrDefault(UserRecordColumnsInfo.NAME_COL_NAME, ""))
  End Sub
  
  Private Function CreateTerm(fromMonth As Integer, fromDay As Integer, toMonth As Integer, toDay As Integer) As DateTerm
    Dim begin As New DateTime(2016, fromMonth, fromDay)
    Dim _end  As New DateTime(2016, toMonth,   toDay)
    
    Return New DateTerm(begin, _end)
  End Function
End Class
