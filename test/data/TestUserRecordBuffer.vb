'
' 日付: 2017/01/02
'
Imports NUnit.Framework

Imports System.Data
Imports System.Linq
Imports Common.Account
Imports Common.Extensions

<TestFixture> _
Public Class TestUserRecordBuffer
  
  ' ユーザ情報オブジェクトを作成する
  Private Function CreateUserInfo(name As String, id As String) As UserInfo
    Return New UserInfo(name, id, "xxx")
  End Function
  
  <Test> _
  Public Sub TestCreateUserRecord
    Dim p As New ExcelProperties("test/ExcelForTest.properties")
    
    '
    ' ユーザレコードが作成されることをテスト
    ' ただし、作成されるだけでバッファオブジェクトに登録はされない
    '
    Dim buffer As New UserRecordBuffer(p)
    Dim user   As UserInfo   = CreateUserInfo("john", "001")
    Dim record As UserRecord = buffer.CreateUserRecord(user)
    
    AssertEqualUser(user, record)
    Assert.False(buffer.Stored(user))
  End Sub
  
  <Test> _
  Public Sub TestGetUserRecord
    Dim p As New ExcelProperties("test/ExcelForTest.properties")
    
    '
    ' ユーザレコードが作成されることをテスト
    ' 作成されたレコードはバッファに登録される
    '
    Dim buffer As New UserRecordBuffer(p)
    Dim user   As UserInfo   = CreateUserInfo("john", "001")
    Dim record As UserRecord = buffer.GetUserRecord(user)
    
    AssertEqualUser(user, record)
    Assert.True(buffer.Stored(user))    
  End Sub
  
  <Test> _
  Public Sub TestGetUserRecordAll
    Dim p As New ExcelProperties("test/ExcelForTest.properties")
    
    '
    ' 全ユーザのレコードを取得することをテスト
    '
    Dim buffer As New UserRecordBuffer(p)
    Dim users As UserInfo() = { 
      CreateUserInfo("john",   "001"),
      CreateUserInfo("paul",   "002"),
      CreateUserInfo("george", "003")
    }
    Array.ForEach(users, Sub(user) buffer.GetUserRecord(user))
    
    Dim l As List(Of UserRecord) = buffer.GetUserRecordAll().ToList
    Assert.AreEqual(3, l.Count)
    For idx As Integer = 0 To users.Count - 1
      AssertEqualUser(users(idx), l(idx))  
    Next
  End Sub
  
  ' ユーザ情報とユーザレコードが同一ユーザのものであるか判定する
  Private Sub AssertEqualUser(user As UserInfo, record As UserRecord)
    Assert.AreEqual(user.GetName,      record.GetName)
    Assert.AreEqual(user.GetSimpleId,  record.GetIdNumber)
  End Sub
  
  <Test> _
  Public Sub TestPlusToTotalRecord
    Dim p As New ExcelProperties("test/ExcelForTest.properties")
    
    '
    ' 集計レコードに値が加算されるかテスト
    '
    Dim buffer As New UserRecordBuffer(p)
    
    Dim user   As UserInfo   = CreateUserInfo("john", "001")
    Dim record As UserRecord = buffer.GetUserRecord(user)
    InsertData(record)
    
    Dim user2   As UserInfo   = CreateUserInfo("paul", "002")
    Dim record2 As UserRecord = buffer.GetUserRecord(user2)
    InsertData(record2)
    
    Dim user3   As UserInfo   = CreateUserInfo("george", "003")
    Dim record3 As UserRecord = buffer.GetUserRecord(user3)
    InsertData(record3)
    
    ' 指定したユーザの値を集計レコードに加算
    buffer.PlusToTotalRecord(user)
    Dim total As UserRecord = buffer.GetTotalRecord()
    Dim table As DataTable  = total.GetRecord(10)
    Assert.AreEqual(14, table.Rows(3).Field(Of Integer)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    
    buffer.PlusToTotalRecord(user2)
    Dim table2 As DataTable = total.GetRecord(11)
    Assert.AreEqual(46, table2.Rows(11).Field(Of Integer)("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    
    buffer.PlusToTotalRecord(user3)
    Dim table3 As DataTable = total.GetRecord(12)
    Assert.AreEqual(576, table3.Rows(15).Field(Of Integer)("item3" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    
    '
    ' 作業時間が０の行の値は集計値に加算されてないことをテスト
    '
    Dim totalE As UserRecord = buffer.GetTotalRecordExceptedUnfilled()
    Dim tableE As DataTable  = totalE.GetRecord(10)
    Assert.AreEqual(54, table.Rows(7).Field(Of Integer)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(0, tableE.Rows(7).GetOrDefault("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME, 0))
    
    '
    ' 1度加算されたユーザは２度加算されないことをテスト
    '
    buffer.PlusToTotalRecord(user)
    Assert.AreEqual(42, table.Rows(3).Field(Of Integer)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(69, table2.Rows(11).Field(Of Integer)("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(576, table3.Rows(15).Field(Of Integer)("item3" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
  End Sub
  
  <Test> _
  Public Sub TestMinusToTotalRecord
    Dim p As New ExcelProperties("test/ExcelForTest.properties")
    
    '
    ' 集計レコードから値が減算されるかテスト
    '
    Dim buffer As New UserRecordBuffer(p)
    
    Dim user   As UserInfo   = CreateUserInfo("john", "001")
    Dim record As UserRecord = buffer.GetUserRecord(user)
    InsertData(record)
    buffer.PlusToTotalRecord(user)
    
    Dim user2   As UserInfo   = CreateUserInfo("paul", "002")
    Dim record2 As UserRecord = buffer.GetUserRecord(user2)
    InsertData(record2)
    buffer.PlusToTotalRecord(user2)
    
    Dim user3   As UserInfo   = CreateUserInfo("george", "003")
    Dim record3 As UserRecord = buffer.GetUserRecord(user3)
    InsertData(record3)
    buffer.PlusToTotalRecord(user3)
    
    ' 指定したユーザの値を集計レコードに加算
    Dim total  As UserRecord = buffer.GetTotalRecord()
    Dim totalE As UserRecord = buffer.GetTotalRecordExceptedUnfilled()
    
    Dim table   As DataTable = total.GetRecord(10)
    Dim tableE  As DataTable = totalE.GetRecord(10)
    Dim table2  As DataTable = total.GetRecord(11)
    Dim tableE2 As DataTable = totalE.GetRecord(11)
     
    Assert.AreEqual(42,  table.Rows(3).Field(Of Integer)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(42,  tableE.Rows(3).Field(Of Integer)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(123, table2.Rows(29).Field(Of Integer)("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(123, tableE2.Rows(29).Field(Of Integer)("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    
    buffer.MinusToTotalRecord(user)
    Assert.AreEqual(28, table.Rows(3).Field(Of Integer)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(28, tableE.Rows(3).Field(Of Integer)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(82, table2.Rows(29).Field(Of Integer)("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(82, tableE2.Rows(29).Field(Of Integer)("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    
    buffer.MinusToTotalRecord(user2)
    Assert.AreEqual(14, table.Rows(3).Field(Of Integer)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(14, tableE.Rows(3).Field(Of Integer)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(41, table2.Rows(29).Field(Of Integer)("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(41, tableE2.Rows(29).Field(Of Integer)("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    
    buffer.MinusToTotalRecord(user3)
    Assert.AreEqual(0, table.Rows(3).Field(Of Integer)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(0, tableE.Rows(3).Field(Of Integer)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(0, table2.Rows(29).Field(Of Integer)("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(0, tableE2.Rows(29).Field(Of Integer)("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    
    '
    ' １度減算されたユーザは２度減算されないことをテスト
    '
    buffer.MinusToTotalRecord(user)
    Assert.AreEqual(0, table.Rows(3).Field(Of Integer)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(0, tableE.Rows(3).Field(Of Integer)("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(0, table2.Rows(29).Field(Of Integer)("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
    Assert.AreEqual(0, tableE2.Rows(29).Field(Of Integer)("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME))
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
End Class
