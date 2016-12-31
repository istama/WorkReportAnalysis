'
' 日付: 2016/12/30
'
Imports NUnit.Framework

Imports System.Data
Imports Common.COM

<TestFixture> _
Public Class TestColumnInfo
  
  <Test> _
  Public Sub TestCreate
    '
    ' 正しくコンストラクションされることをテスト
    '
    
    ' 通常のファクトリーメソッドのテスト
    Dim col As ColumnInfo = ColumnInfo.Create("name", "A", GetType(String), True)
    Assert.AreEqual("name",          col.name) ' 列名
    Assert.AreEqual("A",             col.col)  ' 列
    Assert.AreEqual(GetType(String), col.type) ' 型
    
    ' Excelの列を指定しないファクトリーメソッド
    Dim col2 As ColumnInfo = ColumnInfo.Create("age", GetType(Integer))
    Assert.AreEqual("age",            col2.name) ' 列名
    Assert.AreEqual(String.Empty,     col2.col)  ' 列
    Assert.AreEqual(GetType(Integer), col2.type) ' 型
    
    ' ダミー列生成用のファクトリーメソッド
    Dim col3 As ColumnInfo = ColumnInfo.Dummy
    Assert.AreEqual(String.Empty,    col3.name) ' 列名
    Assert.AreEqual(String.Empty,    col3.col)  ' 列
    Assert.AreEqual(GetType(String), col3.type) ' 型
    
    '
    ' 不正な列名なので例外を投げることをテスト
    '
    Try
      ' 列名が空
      Dim fail As ColumnInfo = ColumnInfo.Create("", "A", GetType(Integer), False)
      Assert.Fail("must throw exception")
    Catch ex As Exception
    End Try
    
    Try
      ' 列が空
      Dim fail As ColumnInfo = ColumnInfo.Create("age", "", GetType(Integer), False)
      Assert.Fail("must throw exception")
    Catch ex As Exception
    End Try
    
    Try
      ' 列の値が不正
      Dim fail As ColumnInfo = ColumnInfo.Create("age", "あ", GetType(Integer), False)
      Assert.Fail("must throw exception")
    Catch ex As Exception
    End Try
  End Sub
  
  <Test> _
  Public Sub TestCreateDataColumn
    '
    ' DataColumnが生成されることをテスト
    '
    
    ' 通常の列
    Dim dataCol As DataColumn = ColumnInfo.Create("name", "A", GetType(String), True).CreateDataColumn
    Assert.AreEqual("name",          dataCol.ColumnName)
    Assert.AreEqual(GetType(String), dataCol.DataType)
    
    ' Excelの列を持たない列
    Dim dataCol2 As DataColumn = ColumnInfo.Create("age", GetType(Integer), False).CreateDataColumn
    Assert.AreEqual("age",            dataCol2.ColumnName)
    Assert.AreEqual(GetType(Integer), dataCol2.DataType)
    
    ' ダミー列
    Dim dataCol3 As DataColumn = ColumnInfo.Dummy().CreateDataColumn
    Assert.AreEqual(Nothing, dataCol3)
  End Sub
  
  <Test>
  Public Sub TestCreateExcelColumnNode
    '
    ' Excelを読み込むためのノードが生成されることをテスト
    '
    
    ' 通常の列
    Dim node As Nullable(Of ExcelColumnNode) = 
      ColumnInfo.Create("name", "A", GetType(String), True).CreateExcelColumnNode
    Assert.AreEqual("name", node.Value.GetName)
    Assert.AreEqual("A",    node.Value.GetCol)
    
    ' Excel列を持たない列
    Dim node2 As Nullable(Of ExcelColumnNode) =
      ColumnInfo.Create("name", GetType(Integer), True).CreateExcelColumnNode
    Assert.AreEqual(Nothing, node2)
  End Sub
End Class

<TestFixture> _
Public Class TestWorkItemColumnsInfo
  Private params1 As ExcelProperties.WorkItemParams = CreateParams(1, "name",    "A", "B")
  Private params2 As ExcelProperties.WorkItemParams = CreateParams(2, "age",     "",  "C") 
  Private params3 As ExcelProperties.WorkItemParams = CreateParams(3, "weight",  "D", "")  
  Private params4 As ExcelProperties.WorkItemParams = CreateParams(4, "height",  "",  "")
  Private params5 As ExcelProperties.WorkItemParams = CreateParams(5, "",        "E", "F")
  
  Private Function CreateParams(id As Integer, name As String, cntCol As String, timeCol As String) As ExcelProperties.WorkItemParams
    Return New ExcelProperties.WorkItemParams With {
      .Id           = id,
      .name         = name,
      .WorkCountCol = cntCol,
      .WorkTimeCol  = timeCol
    }
  End Function
  
  <Test> _
  Public Sub TestCreate
    '
    ' 正しくコンストラクションされることをテスト
    '
    
    ' 通常の列設定の場合
    Dim col As WorkItemColumnsInfo = WorkItemColumnsInfo.Create(params1)
    
    Dim cnt As ColumnInfo = col.WorkCountColInfo
    Assert.AreEqual("name" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME, cnt.name)
    Assert.AreEqual("A",                                             cnt.col)
    Assert.AreEqual(GetType(Integer),                                cnt.type)
    
    Dim time As ColumnInfo = col.WorkTimeColInfo
    Assert.AreEqual("name" & WorkItemColumnsInfo.WORKTIME_COL_NAME, time.name)
    Assert.AreEqual("B",                                            time.col)
    Assert.AreEqual(GetType(Double),                                time.type)
    
    ' 件数列がない場合
    Dim col2 As WorkItemColumnsInfo = WorkItemColumnsInfo.Create(params2)
    
    Dim cnt2 As ColumnInfo = col2.WorkCountColInfo
    Assert.AreEqual(ColumnInfo.Dummy, cnt2) ' 列を設定しないとダミー列になる
    
    Dim time2 As ColumnInfo = col2.WorkTimeColInfo
    Assert.AreEqual("age" & WorkItemColumnsInfo.WORKTIME_COL_NAME, time2.name)
    Assert.AreEqual("C",                                           time2.col)
    Assert.AreEqual(GetType(Double),                               time2.type)
    
    ' 時間列がない場合
    Dim col3 As WorkItemColumnsInfo = WorkItemColumnsInfo.Create(params3)
    
    Dim cnt3 As ColumnInfo = col3.WorkCountColInfo
    Assert.AreEqual("weight" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME, cnt3.name)
    Assert.AreEqual("D",                                               cnt3.col)
    Assert.AreEqual(GetType(Integer),                                  cnt3.type)
    
    Dim time3 As ColumnInfo = col3.WorkTimeColInfo
    Assert.AreEqual(ColumnInfo.Dummy, time3) ' 列を設定しないとダミー列になる
    
    ' 列が１つも設定されていない場合
    Try
      ' 列情報のないパラメータは不正なので例外を投げる
      Dim col4 As WorkItemColumnsInfo = WorkItemColumnsInfo.Create(params4)
      Assert.Fail()
    Catch ex As Exception
    End Try
    
    ' 列名が設定されていない場合
    Try
      ' 列名のないパラメータは不正なので例外を投げる
      Dim col5 As WorkItemColumnsInfo = WorkItemColumnsInfo.Create(params5)
      Assert.Fail()
    Catch ex As Exception
    End Try
  End Sub
  
  <Test> _
  Public Sub TestCreateExcelColumnTree
    '
    ' Excelの読み込みノードが作成されることをテスト
    '
    
    ' 通常の列の場合
    Dim node As ExcelColumnNode = WorkItemColumnsInfo.Create(params1).CreateExcelColumnNodeTree
    Assert.AreEqual("A", node.GetCol)
    Assert.AreEqual(1,   node.GetChilds.Count)
    Assert.AreEqual("B", node.GetChilds(0).GetCol)
    
    ' 作業列のない場合
    Dim node2 As ExcelColumnNode = WorkItemColumnsInfo.Create(params2).CreateExcelColumnNodeTree
    Assert.AreEqual("C", node2.GetCol)
    Assert.AreEqual(0,   node2.GetChilds.Count)   
    
    ' 時間列のない場合
    Dim node3 As ExcelColumnNode = WorkItemColumnsInfo.Create(params3).CreateExcelColumnNodeTree
    Assert.AreEqual("D", node3.GetCol)
    Assert.AreEqual(0,   node3.GetChilds.Count) 
  End Sub
End Class

<TestFixture> _
Public Class TestUserRecordColumnsInfo
  Private properties As New ExcelProperties("test/ExcelForTest.properties")  
  
  <Test> _
  Public Sub TestConstructor
    Dim col As UserRecordColumnsInfo = UserRecordColumnsInfo.Create(properties)
    
    '
    ' 列の設定ファイルから列情報オブジェクトが正しくコンストラクションされることをテスト
    '
    Assert.AreEqual("Y", col.noteColInfo.col)
    Assert.AreEqual("Z", col.workDayColInfo.col)
    
    Dim l As New List(Of WorkItemColumnsInfo)(col.WorkItems)
    Assert.AreEqual("item1" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME,        l(0).WorkCountColInfo.name)
    Assert.AreEqual("item1" & WorkItemColumnsInfo.WORKTIME_COL_NAME,         l(0).WorkTimeColInfo.name)
    Assert.AreEqual("item1" & WorkItemColumnsInfo.WORKPRODUCTIVITY_COL_NAME, l(0).WorkProductivityColInfo.name)
    
    Assert.AreEqual("item2" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME,        l(1).WorkCountColInfo.name)
    Assert.AreEqual("item2" & WorkItemColumnsInfo.WORKTIME_COL_NAME,         l(1).WorkTimeColInfo.name)
    Assert.AreEqual("item2" & WorkItemColumnsInfo.WORKPRODUCTIVITY_COL_NAME, l(1).WorkProductivityColInfo.name)
    
    Assert.AreEqual("item3" & WorkItemColumnsInfo.WORKCOUNT_COL_NAME,        l(2).WorkCountColInfo.name)
    Assert.AreEqual(String.Empty,                                            l(2).WorkTimeColInfo.name)
    Assert.AreEqual(String.Empty,                                            l(2).WorkProductivityColInfo.name)
    
    Assert.AreEqual(String.Empty,                                            l(3).WorkCountColInfo.name)
    Assert.AreEqual("item4" & WorkItemColumnsInfo.WORKTIME_COL_NAME,         l(3).WorkTimeColInfo.name)
    Assert.AreEqual(String.Empty,                                            l(3).WorkProductivityColInfo.name)
  End Sub
  
  <Test> _
  Public Sub TestCreateDataTable
    Dim col As UserRecordColumnsInfo = UserRecordColumnsInfo.Create(properties)
    
    '
    ' 列の設定ファイルからDataTableが作成されることをテスト
    '
    Dim table As DataTable = col.CreateDataTable()
    Dim l     As New List(Of WorkItemColumnsInfo)(col.WorkItems)
    Assert.AreEqual(9, table.Columns.Count)
    Assert.AreEqual(l(0).WorkCountColInfo.name,        table.Columns(0).ColumnName)
    Assert.AreEqual(l(0).WorkTimeColInfo.name,         table.Columns(1).ColumnName)
    Assert.AreEqual(l(0).WorkProductivityColInfo.name, table.Columns(2).ColumnName) ' 件数と時間の列がある作業項目のみ生産性の列が生成される
    Assert.AreEqual(l(1).WorkCountColInfo.name,        table.Columns(3).ColumnName)
    Assert.AreEqual(l(1).WorkTimeColInfo.name,         table.Columns(4).ColumnName)
    Assert.AreEqual(l(1).WorkProductivityColInfo.name, table.Columns(5).ColumnName)
    Assert.AreEqual(l(2).WorkCountColInfo.name,        table.Columns(6).ColumnName)
    Assert.AreEqual(l(3).WorkTimeColInfo.name,         table.Columns(7).ColumnName)
    Assert.AreEqual(col.noteColInfo.name,              table.Columns(8).ColumnName)
    
    '
    ' DataTableの１列目に指定した名前の列を追加して生成する
    '
    Dim table2 As DataTable = col.CreateDataTable("NewColumn")
    Dim l2     As New List(Of WorkItemColumnsInfo)(col.WorkItems)
    Assert.AreEqual(10, table2.Columns.Count)
    Assert.AreEqual("NewColumn",                 table2.Columns(0).ColumnName)
    Assert.AreEqual(l2(0).WorkCountColInfo.name, table2.Columns(1).ColumnName)
    Assert.AreEqual(col.noteColInfo.name,        table2.Columns(9).ColumnName)
  End Sub
  
  <Test> _
  Public Sub TestCreateExcelColumnNodeTree
    Dim col As UserRecordColumnsInfo = UserRecordColumnsInfo.Create(properties)
    
    '
    ' 列の設定からExcelを読み込むノードが作成されることをテスト
    '
    Dim node As ExcelColumnNode = col.CreateExcelColumnNodeTree()
    AssertEqualNode(UserRecordColumnsInfo.WORKDAY_COL_NAME, "Z", node)
    
    Dim childs As List(Of ExcelColumnNode) = node.GetChilds()
    Dim CNT  As String = WorkItemColumnsInfo.WORKCOUNT_COL_NAME
    Dim TIME As String = WorkItemColumnsInfo.WORKTIME_COL_NAME
    AssertEqualNode("item1" & CNT,  "A", childs(0))
    AssertEqualNode("item1" & TIME, "B", childs(0).GetChilds(0))
    
    AssertEqualNode("item2" & CNT,  "C", childs(1))
    AssertEqualNode("item2" & TIME, "D", childs(1).GetChilds(0))
    
    AssertEqualNode("item3" & CNT,  "E", childs(2))
    AssertEqualNode("item4" & TIME, "F", childs(3))
    
    AssertEqualNode("Note",         "Y", childs(4))
  End Sub
  
  Private Sub AssertEqualNode(name As String, col As String, node As ExcelColumnNode)
    Assert.AreEqual(name, node.GetName)
    Assert.AreEqual(col,  node.GetCol)
  End Sub
End Class
