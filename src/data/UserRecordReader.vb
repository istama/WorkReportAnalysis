'
' 日付: 2016/10/18
'
Imports Common.COM
Imports Common.Util
Imports Common.Account
Imports System.Data
Imports Common.Extensions
Imports Common.IO

''' <summary>
''' ユーザのExcelデータを読み込むクラス。
''' </summary>
Public Class UserRecordReader
  Private ReadOnly properties As ExcelProperties  
  Private ReadOnly excel As ExcelReaderByColumnTree
  
  Private _cancel As Boolean = False
  
  Public Sub New(properties As ExcelProperties, excel As Excel4)
    If properties Is Nothing Then Throw New ArgumentNullException("properties is null")
    If excel      Is Nothing Then Throw New ArgumentNullException("excel is null")
    
    Me.properties     = properties
    Me.excel          = New ExcelReaderByColumnTree(excel)
  End Sub
  
  Public Sub Cancel
    Me._cancel = True
  End Sub
  
  ''' <summary>
  ''' 指定したユーザのレコードを読み込む。
  ''' </summary>
  Public Sub Read(userRecord As UserRecord)
    If userRecord Is Nothing Then Throw New ArgumentNullException("userRecord is null")
    
    If Me._cancel Then
      Return
    End If
          
    Dim filepath As String = String.Format(Me.properties.ExcelFilePath(), userRecord.GetIdNumber)
    Dim colTree As ExcelColumnNode = userRecord.GetCulumnNodeTree
    
    Try
      Me.excel.Open(filepath, True)
      
      ' 各月のレコードを読み込む
      userRecord.GetRecordDateTerm.MonthlyTerms.ForEach(
        Sub(m)
          Dim table As DataTable = userRecord.GetRecord(m.BeginDate.Month)
        
          If Me._cancel Then
            Return
          End If
          
          Dim sheetName As String = Me.properties.SheetName(m.BeginDate.Month)
          
          For day As Integer = m.BeginDate.Day To m.EndDate.Day
            ' 行を指定してデータを読み込む
            Dim row As Integer = day + Me.properties.FirstRow - 1
            Dim dict As IDictionary(Of String, String) = Me.excel.Read(row, filepath, sheetName, colTree)
            
            ' 読み込んだデータをDataTableの行にセットする
            Dim dataRow As DataRow = table.Rows(day - m.BeginDate.Day)
            dict.Keys.ForEach(
              Sub(k)
                If String.IsNullOrWhiteSpace(dict(k)) Then
                  Return
                End If
              
                'Log.out(userRecord.GetName & " " & m.BeginDate.Month.ToString & "/" & day.ToString & " " & k & ":" & dict(k)) 
                If table.Columns(k).DataType = GetType(Integer) Then
                  Dim v As Integer
                  If Integer.TryParse(dict(k), v) Then
                    dataRow(k) = v
                  End If
                ElseIf table.Columns(k).DataType = GetType(Double) Then
                  Dim v As Double
                  If Double.TryParse(dict(k), v) Then
                    dataRow(k) = v
                  End If
                Else
                  dataRow(k) = dict(k)
                End If
              End Sub)
          Next
        End Sub)
    Catch ex As Exception
      Throw ex
    Finally
      Me.excel.Close(filepath)
    End Try
  End Sub
End Class
