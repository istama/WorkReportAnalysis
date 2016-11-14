'
' 日付: 2016/07/13
'
Imports System.Data

Public Structure GridTable
'	Private userRecord As UserRecord
'	Private dateTerm As DateTerm
'	
'	Public Sub New(userRecord As UserRecord, dateTerms As DateTerm)
'		Me.userRecord = userRecord
'		Me.dateTerm = dateTerms
'	End Sub
'	
'	''' <summary>
'	''' 指定したシートの最終行に各列の合計値や平均値を追加して返す。
'	''' isDependingをtrueにすると、作業時間列が空の場合その行の件数列の値を合計値に含めないようにする。
'	''' </summary>
'	Public Function GetMonthlyTable(sheetName As String, isDepending As Boolean) As DataTable
'		If userRecord.HasTable(sheetName) = False Then Throw New ArgumentException("table is not found. / sheetName: " & sheetName)
'		
'		Dim table As DataTable = userRecord.GetTable(sheetName)
'		userRecord.AddTotalRowToTable(table, isDepending) ' 合計行を追加
'		
'		Return table
'	End Function
'	
'	''' <summary>
'	''' 週合計、月合計、総合計の行を持つテーブルを返す。
'	''' isDependingをtrueにすると、作業時間列が空の場合その行の件数列の値を合計値に含めないようにする。
'	''' </summary>
'	Public Function GetTotalTable(isDepending As Boolean) As DataTable
'		Dim table As DataTable = userRecord.CreateTable
'		' 各週の合計を求めた行をテーブルに追加する
'		For Each term In Me.dateTerm.WeeklyTermList
'			userRecord.AddTotalRowToTable(term, isDepending, table)
'		Next
'		table.Rows.Add(table.NewRow) ' 空行を追加
'		
'		' 各月の合計を求めた行をテーブルに追加する
'		For Each term In Me.dateTerm.MonthlyTermList
'			userRecord.AddTotalRowToTable(term, isDepending, table)
'		Next
'		table.Rows.Add(table.NewRow) ' 空行を追加		
'		
'		' 総合計を求めた行をテーブルに追加する
'		userRecord.AddTotalRowToTable(Me.dateTerm, isDepending, table)
'		
'		Return table
'	End Function
	
End Structure
