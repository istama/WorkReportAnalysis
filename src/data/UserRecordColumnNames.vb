'
' 日付: 2016/11/14
'

''' <summary>
''' レコードの列名を持つ構造体。
''' </summary>
Public Structure UserRecordColumnNames
  Public Const WORKDAY_COL_NAME   As String = "出勤日"
  
  ''' 各作業項目ごとの列名を格納するリスト 
  Private ReadOnly workItems As List(Of WorkItemColumnNames)
  ''' 備考欄の列名 
  Private ReadOnly note As String
  ''' 出勤日の列名 
  Private ReadOnly workDay As String
  
  Public Sub New(properties As ExcelProperties)
    If properties Is Nothing Then Throw New NullReferenceException("properties is null")
    
    Me.workItems = New List(Of WorkItemColumnNames)
    
    Dim idx As Integer = 1
    While True 
      ' Excelのプロパティから新しい作業項目の設定が取得できたなら
      ' そこから列名を生成しリストに格納する。
      ' 取得できなかったらループを抜ける。
      Dim params As ExcelProperties.WorkItemParams = properties.GetWorkItemParams(idx)
      Dim colnames As WorkItemColumnNames? = WorkItemColumnNames.Create(params)
      If colnames.HasValue Then
        Me.workItems.Add(colnames.Value)
        idx += 1
      Else
        Exit While
      End If
    End While
  End Sub
End Structure

''' <summary>
''' １つの作業項目あたりの列名。
''' </summary>
Public Structure WorkItemColumnNames
  Public Const WORKCOUNT_COL_NAME        As String = "件数"
  Public Const WORKTIME_COL_NAME         As String = "作業時間"
  Public Const WORKPRODUCTIVITY_COL_NAME As String = "生産性"
  
  Public ReadOnly WorkCount As String
  Public ReadOnly WorkTime As String
  Public ReadOnly WorkProductivity As String
  
  Private Sub New(itemName As String, useWorkCount As Boolean, useWorkTime As Boolean)
    If useWorkCount Then
      Me.WorkCount = itemName & vbCrLf & WORKCOUNT_COL_NAME
    Else
      Me.WorkCount = String.Empty
    End If
    
    If useWorkTime Then
      Me.WorkTime = itemName & vbCrLf & WORKTIME_COL_NAME
    Else
      Me.WorkTime = String.Empty
    End If
    
    If useWorkCount AndAlso useWorkTime Then
      Me.WorkProductivity = itemName & vbCrLf & WORKPRODUCTIVITY_COL_NAME
    Else
      Me.WorkProductivity = String.Empty
    End If
  End Sub
  
  Public Shared Function Create(params As ExcelProperties.WorkItemParams) As WorkItemColumnNames?
    If params.Name Is Nothing OrElse params.Name = String.Empty Then
      Return Nothing
    Else
      Dim useWorkCount As Boolean = params.WorkCountCol <> String.Empty
      Dim useWorkTime  As Boolean = params.WorkTimeCol  <> String.Empty
      Return New WorkItemColumnNames(params.Name, useWorkCount, useWorkTime)
    End If
  End Function
End Structure