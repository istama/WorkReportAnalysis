'
' 日付: 2016/07/07
'
Imports NUnit.Framework

Imports System.Data
Imports Common.Account
Imports Common.IO
Imports Common.COM
Imports Common.Extensions
Imports Common.Util

<TestFixture> _
Public Class TestUserRecord
  
  ''' <summary>
  ''' ユーザを作成
  ''' </summary>
  Private Function CreateUserRecord(begin As DateTime, _end As DateTime) As UserRecord
    Dim p        As New ExcelProperties("test/ExcelForTest.properties")
    Dim col      As UserRecordColumnsInfo = UserRecordColumnsInfo.Create(p)
    Dim userinfo As New UserInfo("john", "001", "pass")
    
    Return New UserRecord(userinfo, col, New DateTerm(begin, _end))    
  End Function
  
  <Test> _
  Public Sub TestConstruction
    '
    ' ユーザレコードが正しくコンストラクションされることをテスト
    '
    Dim record As UserRecord = CreateUserRecord(New DateTime(2016, 10, 1), New DateTime(2016, 12, 31))
    
    Assert.AreEqual("john",                     record.GetName, "user name")
    Assert.AreEqual("001",                      record.GetIdNumber, "user id")
    Assert.AreEqual(New DateTime(2016, 10, 1),  record.GetRecordDateTerm.BeginDate, "begin date")
    Assert.AreEqual(New DateTime(2016, 12, 31), record.GetRecordDateTerm.EndDate, "end date")
    
    ' Excelを読み込む列ノードが正しく生成されているか
    Dim node As ExcelColumnNode = record.GetExcelCulumnNodeTree()
    Assert.AreEqual("Z", node.GetCol)
    Assert.AreEqual(5, node.GetChilds().Count)
    
    Dim CNT  As String = WorkItemColumnsInfo.WORKCOUNT_COL_NAME
    Dim TIME As String = WorkItemColumnsInfo.WORKTIME_COL_NAME
    Dim PROD As String = WorkItemColumnsInfo.WORKPRODUCTIVITY_COL_NAME
    For month As Integer = 10 To 12
      Dim table As DataTable = record.GetRecord(month)
      ' 各月のテーブルの列が正しく作成されているか
      Assert.AreEqual("item1" & CNT,  table.Columns(0).ColumnName)
      Assert.AreEqual("item1" & TIME, table.Columns(1).ColumnName)
      Assert.AreEqual("item1" & PROD, table.Columns(2).ColumnName)
      Assert.AreEqual("item2" & CNT,  table.Columns(3).ColumnName)
      Assert.AreEqual("item2" & TIME, table.Columns(4).ColumnName)
      Assert.AreEqual("item2" & PROD, table.Columns(5).ColumnName)
      Assert.AreEqual("item3" & CNT,  table.Columns(6).ColumnName)
      Assert.AreEqual("item4" & TIME, table.Columns(7).ColumnName)
      Assert.AreEqual("Note",         table.Columns(8).ColumnName)
      
      ' テーブルの行オブジェクトが１ヶ月の日数分生成されているか
      Assert.AreEqual(DateTime.DaysInMonth(2016, month), table.Rows.Count, "month " & month.ToString)
    Next
    
    '
    ' データ期間の始まりが月の途中でも、データテーブルの行は１ヶ月の日数分生成されることをテスト
    '
    Dim record2 As UserRecord = CreateUserRecord(New DateTime(2016, 10, 15), New DateTime(2016, 12, 20))
    For month As Integer = 10 To 12
      Dim table As DataTable = record.GetRecord(month)
      ' テーブルの行オブジェクトが１ヶ月の日数分生成されているか
      Assert.AreEqual(DateTime.DaysInMonth(2016, month), table.Rows.Count, "month " & month.ToString)
    Next
    
    '
    ' データ期間を１年以上にすると例外を投げる
    '
    Try
      CreateUserRecord(New DateTime(2016, 1, 1), New DateTime(2017, 1, 1))
      Assert.Fail()
    Catch ex As Exception
    End Try
  End Sub
  
  ''' <summary>
  ''' データテーブルに仮データをセットする。
  ''' </summary>
  Private Sub InsertData(record As UserRecord)
    For m As Integer = 10 To 12
      Dim t As DataTable = record.GetRecord(m)
      For d As Integer = 1 To DateTime.DaysInMonth(2016, m)
        If d Mod 4 = 0 Then
          t.Rows(d - 1)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME) = d + m
          
          Dim time As Double = 1.0
          If d Mod 8 = 0 Then time = 0
          t.Rows(d - 1)("item1" & WorkItemColumnsInfo.WORKTIME_COL_NAME)  = time
        End If
        
        If d Mod 6 = 0 Then
          t.Rows(d - 1)("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME) = d + m
          t.Rows(d - 1)("item2" & WorkItemColumnsInfo.WORKTIME_COL_NAME)  = 1
        End If
        
        If d Mod 8 = 0 Then
          t.Rows(d - 1)("item3" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME) = d * m
        End If
        
        If d Mod 10 = 0 Then
          t.Rows(d - 1)("item4" & WorkItemColumnsInfo.WORKTIME_COL_NAME)  = m
        End If
      Next
    Next
  End Sub
  
  <Test> _
  Public Sub TestGetDailyDataTable
    '
    ' 指定した期間の1日ごとのデータを格納したデータテーブルを取得できるかテスト
    '
    Dim record As UserRecord = CreateUserRecord(New DateTime(2016, 10, 15), New DateTime(2016, 12, 10))
    InsertData(record)
    
    Dim CNT  As String = WorkItemColumnsInfo.WORKCOUNT_COL_NAME
    Dim TIME As String = WorkItemColumnsInfo.WORKTIME_COL_NAME
    Dim PROD As String = WorkItemColumnsInfo.WORKPRODUCTIVITY_COL_NAME
    
    ' 10月のデータを取得
    ' データ期間は10月15日から始まる
    Dim table1 As DataTable = record.GetDailyDataTable(2016, 10)
    Assert.AreEqual(17, table1.Rows.Count)
    Assert.AreEqual(26, table1.Rows(1).GetOrDefault("item1" & CNT, 0))
    Assert.AreEqual(30, table1.Rows(5).GetOrDefault("item1" & CNT, 0))
    Assert.AreEqual(34, table1.Rows(9).GetOrDefault("item1" & CNT, 0))
    Assert.AreEqual(38, table1.Rows(13).GetOrDefault("item1" & CNT, 0))
    
    ' 11月のデータを取得
    Dim table2 As DataTable = record.GetDailyDataTable(2016, 11)
    Assert.AreEqual(30, table2.Rows.Count)
    Assert.AreEqual(17, table2.Rows(5).GetOrDefault("item2" & CNT, 0))
    Assert.AreEqual(23, table2.Rows(11).GetOrDefault("item2" & CNT, 0))
    Assert.AreEqual(29, table2.Rows(17).GetOrDefault("item2" & CNT, 0))
    Assert.AreEqual(35, table2.Rows(23).GetOrDefault("item2" & CNT, 0))
    Assert.AreEqual(41, table2.Rows(29).GetOrDefault("item2" & CNT, 0))
    
    ' 11/20から12/31のデータを取得
    ' 12月のデータ期間は10日で終わる
    Dim table3 As DataTable = record.GetDailyDataTable(New DateTerm(New DateTime(2016, 11, 20), New DateTime(2016, 12, 31)))
    Assert.AreEqual(21,   table3.Rows.Count)
    Assert.AreEqual(264,  table3.Rows(4).GetOrDefault("item3" & CNT, 0))
    Assert.AreEqual(96,   table3.Rows(18).GetOrDefault("item3" & CNT, 0))
    Assert.AreEqual(11.0, table3.Rows(0).GetOrDefault("item4" & TIME, 0.0))
    Assert.AreEqual(11.0, table3.Rows(10).GetOrDefault("item4" & TIME, 0.0))
    Assert.AreEqual(12.0, table3.Rows(20).GetOrDefault("item4" & TIME, 0.0))
  End Sub
  
  <Test> _
  Public Sub TestGetWeeklyDataTable
    '
    ' 指定した期間の1週間ごとのデータを格納したデータテーブルを取得できるかテスト
    '
    Dim record As UserRecord = CreateUserRecord(New DateTime(2016, 10, 15), New DateTime(2016, 12, 10))
    InsertData(record)
    
    Dim CNT  As String = WorkItemColumnsInfo.WORKCOUNT_COL_NAME
    Dim TIME As String = WorkItemColumnsInfo.WORKTIME_COL_NAME
    Dim PROD As String = WorkItemColumnsInfo.WORKPRODUCTIVITY_COL_NAME
    
    Dim term1  As New DateTerm(New DateTime(2016, 10, 1), New DateTime(2016, 10, 31))
    Dim table1 As DataTable = record.GetWeeklyDataTable(term1, False)
    Assert.AreEqual(4,  table1.Rows.Count)
    Assert.AreEqual(0,  table1.Rows(0).GetOrDefault("item1" & CNT, 0))
    Assert.AreEqual(56, table1.Rows(1).GetOrDefault("item1" & CNT, 0))
    Assert.AreEqual(72, table1.Rows(2).GetOrDefault("item1" & CNT, 0))
    Assert.AreEqual(0,  table1.Rows(3).GetOrDefault("item1" & CNT, 0))
    
    ' 作業時間がセットされていない行は合計に含めない
    Dim term2  As New DateTerm(New DateTime(2016, 10, 1), New DateTime(2016, 11, 5))
    Dim table2 As DataTable = record.GetWeeklyDataTable(term2, True)
    Assert.AreEqual(4,  table2.Rows.Count)
    Assert.AreEqual(0,  table2.Rows(0).GetOrDefault("item1" & CNT, 0))
    Assert.AreEqual(30, table2.Rows(1).GetOrDefault("item1" & CNT, 0))
    Assert.AreEqual(38, table2.Rows(2).GetOrDefault("item1" & CNT, 0))
    Assert.AreEqual(15, table2.Rows(3).GetOrDefault("item1" & CNT, 0))
  End Sub
  
  <Test> _
  Public Sub TestGetMonthlyDataTable
    '
    ' 指定した期間の1ヶ月ごとのデータを格納したデータテーブルを取得できるかテスト
    '
    Dim record As UserRecord = CreateUserRecord(New DateTime(2016, 10, 15), New DateTime(2016, 12, 10))
    InsertData(record)
    
    Dim CNT  As String = WorkItemColumnsInfo.WORKCOUNT_COL_NAME
    Dim TIME As String = WorkItemColumnsInfo.WORKTIME_COL_NAME
    Dim PROD As String = WorkItemColumnsInfo.WORKPRODUCTIVITY_COL_NAME
    
    Dim term1  As New DateTerm(New DateTime(2016, 10, 1), New DateTime(2016, 11, 25))
    Dim table1 As DataTable = record.GetMonthlyDataTable(term1, False)
    Assert.AreEqual(2,  table1.Rows.Count)
    Assert.AreEqual(128, table1.Rows(0).GetOrDefault("item1" & CNT, 0))
    Assert.AreEqual(104, table1.Rows(1).GetOrDefault("item2" & CNT, 0))
    
    ' 作業時間がセットされていない行は合計に含めない
    Dim term2  As New DateTerm(New DateTime(2016, 10, 1), New DateTime(2016, 11, 25))
    Dim table2 As DataTable = record.GetMonthlyDataTable(term2, True)
    Assert.AreEqual(2,  table2.Rows.Count)
    Assert.AreEqual(68, table2.Rows(0).GetOrDefault("item1" & CNT, 0))
    Assert.AreEqual(69, table2.Rows(1).GetOrDefault("item1" & CNT, 0))
  End Sub
End Class
