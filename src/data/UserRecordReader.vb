'
' 日付: 2016/10/18
'
Imports System.Data
Imports System.Linq

Imports Common.Account
Imports Common.COM
Imports Common.Util

''' <summary>
''' ユーザのExcelデータを読み込むクラス。
''' </summary>
Public Class UserRecordReader
  Private ReadOnly properties As ExcelProperties  
  Private ReadOnly excel As ExcelReaderByColumnTree
  
  Public Sub New(properties As ExcelProperties, excel As IExcel)
    If properties Is Nothing Then Throw New ArgumentNullException("properties is null")
    If excel      Is Nothing Then Throw New ArgumentNullException("excel is null")
    
    Me.properties     = properties
    Me.excel          = New ExcelReaderByColumnTree(excel)
  End Sub
  
  ''' <summary>
  ''' 指定したユーザのレコードを読み込む。
  ''' </summary>
  Public Sub Read(userRecord As UserRecord)
    If userRecord Is Nothing Then Throw New ArgumentNullException("userRecord is null")
    
    ' Excelファイルのパスを作成
    Dim filepath As String = String.Format(Me.properties.ExcelFilePath(), userRecord.GetIdNumber)
    ' 読み込む列
    Dim colTree As ExcelColumnNode = userRecord.GetExcelCulumnNodeTree
    
    Try
      ' Excelファイルを開く
      Me.excel.Open(filepath, True)
      
      ' 各月のレコードを読み込む
      For Each m As DateTerm In userRecord.GetRecordDateTerm.MonthlyTerms
        Dim table As DataTable = userRecord.GetRecord(m.BeginDate.Month)
        
        Dim rowIdx As Integer = m.BeginDate.Day - 1
        
        ' 1行ずつ読み込む
        For Each dict In ReadRecord(filepath, m, colTree)
          ' 読み込んだデータをDataTableの行にセットする
          Dim dataRow As DataRow = table.Rows(rowIdx)
          
          For Each k In dict.Keys.Where(Function(key) Not String.IsNullOrWhiteSpace(dict(key)))
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
          
          rowIdx += 1
        Next
      Next
    Catch ex As Exception
      Throw ex
    Finally
      Me.excel.Close(filepath)
    End Try
  End Sub
  
  ''' <summary>
  ''' 指定した月のレコードを１行ずつ読み込む。
  ''' </summary>
  Private Iterator Function ReadRecord(filepath As String, monthlyTerm As DateTerm, colTree As ExcelColumnNode) As IEnumerable(Of IDictionary(Of String, String))
    ' シート名を作成
    Dim sheetName As String = Me.properties.SheetName(monthlyTerm.BeginDate.Month)
    
    ' 日にちごとにデータを読み込む
    Dim first As Integer = monthlyTerm.BeginDate.Day + Me.properties.FirstRow - 1
    Dim _end  As Integer = monthlyTerm.EndDate.Day   + Me.properties.FirstRow - 1
    For row As Integer = first To _end
      Yield Me.excel.Read(row, filepath, sheetName, colTree)
    Next
  End Function
End Class
