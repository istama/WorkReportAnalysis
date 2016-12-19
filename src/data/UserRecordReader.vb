'
' 日付: 2016/10/18
'
Imports System.Linq
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
  
  Public Sub New(properties As ExcelProperties, excel As IExcel)
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
    
    ' Excelファイルのパスを作成
    Dim filepath As String = String.Format(Me.properties.ExcelFilePath(), userRecord.GetIdNumber)
    ' 読み込む列
    Dim colTree As ExcelColumnNode = userRecord.GetCulumnNodeTree
    
    Try
      ' Excelファイルを開く
      Me.excel.Open(filepath, True)
      
      ' 各月のレコードを読み込む
      For Each terms As DateTerm In userRecord.GetRecordDateTerm.MonthlyTerms
        Dim table As DataTable = userRecord.GetRecord(m.BeginDate.Month)
        
        Dim sheetName As String = Me.properties.SheetName(m.BeginDate.Month)
        
        ' 日にちごとにデータを読み込む
        For day As Integer = m.BeginDate.Day To m.EndDate.Day
          ' 行を指定してデータを読み込む
          Dim row As Integer = day + Me.properties.FirstRow - 1
          Dim dict As IDictionary(Of String, String) = Me.excel.Read(row, filepath, sheetName, colTree)
            
          ' 読み込んだデータをDataTableの行にセットする
          Dim dataRow As DataRow = table.Rows(day - m.BeginDate.Day)
          For Each k In dict.Keys.Where(Function(k) Not String.IsNullOrWhiteSpace(dict(k)))
            ' 取得したデータをその列の値のデータ型に変換してセットする
            Dim t As Type = table.Columns(k).DataType
            If t = GetType(Integer) Then
              Dim v As Integer
              If Integer.TryParse(dict(k), v) Then
                dataRow(k) = v
              End If
            ElseIf t = GetType(Double) Then
              Dim v As Double
              If Double.TryParse(dict(k), v) Then
                dataRow(k) = v
              End If
            ElseIf t = GetType(String) Then
              dataRow(k) = dict(k)
            End If            
          Next
        Next
      Next
    Catch ex As Exception
      Throw ex
    Finally
      Me.excel.Close(filepath)
    End Try
  End Sub
End Class
