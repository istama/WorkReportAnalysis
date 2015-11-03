
Imports System.Security.Cryptography
Imports System.Runtime.Serialization
Imports System.Reflection
Imports MP.Details.IO
Imports MP.Details.Serialize
'Imports MP.Details.Security
Imports MP.Details.Sys
Imports AP = MP.Utils.Common.AppProperties

Namespace Utils

  Namespace Common
    Public Class MyLog
      Public Shared LogMode As Boolean = True
      Public Shared Log As Logging.Log = My.Application.Log

      Public Shared Sub Write(msg As String)
        If LogMode Then
          My.Application.Log.WriteEntry(msg)
        End If
      End Sub
    End Class

    Public Class MsgBox
      Public Shared Sub ShowWarn(msg As String)
        MessageBox.Show(msg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
      End Sub
      Public Shared Sub ShowWarn(ex As Exception)
        Show(ex, "Warning", MessageBoxIcon.Warning)
      End Sub

      Public Shared Sub ShowError(msg As String)
        MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
      End Sub
      Public Shared Sub ShowError(ex As Exception)
        Show(ex, "Error", MessageBoxIcon.Warning)
      End Sub
      Private Shared Sub Show(ex As Exception, title As String, icon As MessageBoxIcon)
        MessageBox.Show(ex.Message & vbCrLf & vbCrLf & ex.StackTrace, title, MessageBoxButtons.OK, icon)
      End Sub

    End Class

    Public Class FilePath
      Public Shared Function UserinfoFilePath() As String
        Return GetPath(AP.MANAGER, AP.KEY_USERINFO_FILE_DIR, AP.KEY_USERINFO_FILE_NAME)
      End Function

      Public Shared Function ReleaseVersionInfoFilePath() As String
        Return GetPath(AP.MANAGER, AP.KEY_LATEST_VERSIONINFO_FILE_DIR, AP.KEY_LATEST_VERSIONINFO_FILE_NAME)
      End Function

      Public Shared Function UpdateScriptPath() As String
        Return GetPath(AP.MANAGER, AP.KEY_UPDATE_SCRIPT_DIR, AP.KEY_UPDATE_SCRIPT_NAME)
      End Function

      Public Shared Function ExcludeFileForUpdatePath() As String
        Return GetPath(AP.MANAGER, AP.KEY_UPDATE_SCRIPT_DIR, AP.KEY_EXCLUDED_FILE_NAME_FROM_UPDATE)
      End Function

      Public Shared Function GetPath(entry As PropertyManager, dirKey As String, fileKey As String) As String
        Dim dir As String = entry.GetValue(dirKey)
        Dim file As String = entry.GetValue(fileKey)
        If dir <> "" AndAlso file <> "" Then
          Return dir & "\" & file
        Else
          Return ""
        End If
      End Function
    End Class

    Public Class AppProperties
      Private Shared SETTING_FILE_NAME = "setting.properties"

      Public Shared KEY_USERINFO_FILE_DIR = "UserInfoFileDir"
      Public Shared KEY_USERINFO_FILE_NAME = "UserInfoFileName"
      Public Shared KEY_WRITE_LOG = "WriteLogMode"
      Public Shared KEY_ENABLE_AUTO_UPDATE = "AutoUpdateMode"
      Public Shared KEY_LATEST_VERSIONINFO_FILE_DIR = "LatestVersionInfoFileDir"
      Public Shared KEY_LATEST_VERSIONINFO_FILE_NAME = "LatestVersionInfoFileName"
      Public Shared KEY_RELEASE_DIR_FOR_UPDATE = "ReleaseDirForUpdate"
      Public Shared KEY_UPDATE_SCRIPT_DIR = "UpdateScriptDir"
      Public Shared KEY_UPDATE_SCRIPT_NAME = "UpdateScriptName"
      Public Shared KEY_EXCLUDED_FILE_NAME_FROM_UPDATE = "ExcludedFileNameFromUpdate"

      Public Shared MANAGER = New PropertyManager(SETTING_FILE_NAME, DefaultSettingProperties(), True)

      Private Shared Function DefaultSettingProperties() As IDictionary(Of String, String)
        Dim dict As IDictionary(Of String, String) = New Dictionary(Of String, String)
        dict(KEY_USERINFO_FILE_DIR) = App.GetCurrentDirectory()
        dict(KEY_USERINFO_FILE_NAME) = "userinfo.txt"
        dict(KEY_WRITE_LOG) = "True"
        dict(KEY_ENABLE_AUTO_UPDATE) = "True"
        dict(KEY_LATEST_VERSIONINFO_FILE_DIR) = ""
        dict(KEY_LATEST_VERSIONINFO_FILE_NAME) = "version.txt"
        dict(KEY_RELEASE_DIR_FOR_UPDATE) = ""
        dict(KEY_UPDATE_SCRIPT_DIR) = App.GetCurrentDirectory()
        dict(KEY_UPDATE_SCRIPT_NAME) = "update.bat"
        dict(KEY_EXCLUDED_FILE_NAME_FROM_UPDATE) = "NotUpdatedFiles.txt"
        Return dict
      End Function

    End Class

    Public Class PropertyManager
      Private FilePath As String
      Private DefProperties As IDictionary(Of String, String)
      Private Properties As IDictionary(Of String, String) = New Dictionary(Of String, String)

      Private AllowDefPropertyKeysOnly As Boolean

      Private hasRead As Boolean = False

      Public Sub New(filePath As String, def As IDictionary(Of String, String), allowDefPropertyKeysOnly As Boolean)
        Me.FilePath = filePath
        Me.DefProperties = def
        Me.AllowDefPropertyKeysOnly = allowDefPropertyKeysOnly
      End Sub

      Public Function GetValue(key As String) As String
        Load()

        If Properties.ContainsKey(key) Then
          Return Properties(key)
        Else
          Reload(DefProperties)
          If Properties.ContainsKey(key) Then
            Return Properties(key)
          Else
            Return ""
          End If
        End If
      End Function

      Private Sub Load()
        If Not hasRead Then
          Try
            Properties = PropertyAccessor.GetProp(FilePath)
            If AllowDefPropertyKeysOnly Then
              Dim nProp As IDictionary(Of String, String) = RemoveKeysThatDoseNotContainsToDefProperties(Properties)
              If nProp.Count() < Properties.Count() Then
                PropertyAccessor.SetProp(FilePath, nProp)
                Properties = nProp
              End If
            End If
          Catch ex As System.IO.FileNotFoundException
            PropertyAccessor.SetProp(FilePath, DefProperties)
            Properties = DefProperties
          Catch ex As Exception
            MsgBox.ShowError(ex)
          End Try
          hasRead = True
        End If
      End Sub

      Private Sub Reload(addedProp As IDictionary(Of String, String))
        PropertyAccessor.AppendOnlyPropThatDoesNotExists(FilePath, addedProp)
        hasRead = False
        Load()
      End Sub

      Private Function RemoveKeysThatDoseNotContainsToDefProperties(prop As IDictionary(Of String, String)) As IDictionary(Of String, String)
        Dim nDict As New Dictionary(Of String, String)
        For Each k As String In prop.Keys
          If DefProperties.ContainsKey(k) Then
            nDict.Add(k, prop(k))
          End If
        Next
        Return nDict
      End Function

    End Class

    Public Class UpdateManager
      Private UpdateScriptFilePath As String
      Private VersionFilePath As String

      Sub New(updateScriptFilePath As String, versionFilePath As String)
        Me.UpdateScriptFilePath = updateScriptFilePath
        Me.VersionFilePath = versionFilePath
      End Sub

      Public Function hasUpdated() As Boolean
        MyLog.Write("最新バージョンか確認します。VersionFilePath: " & VersionFilePath)
        If VersionFilePath = "" Then
          MyLog.Write("最新バージョン情報へのパスがありません。")
          Return False
        ElseIf UpdateScriptFilePath = "" Then
          MyLog.Write("アップデートバッチへのパスがありません。")
          Return False
        Else
          Dim text As List(Of String) = FileAccessor.Read(VersionFilePath)

          If text.Count = 0 Then
            MyLog.Write("最新バージョン情報は確認できませんでした。")
            Return False
          ElseIf Version.IsApplicationOfLatestVersion(text(0)) Then
            MyLog.Write("現在最新バージョンです。 version: " & text(0))
            Return False
          Else
            MyLog.Write("最新バージョンがリリースされました。 version: " & text(0))
            Return True
          End If
        End If
      End Function

      Public Sub GenerateDefaultUpdateBatchIfEmpty(releaseVersionDir As String, excludeFilePath As String)
        MyLog.Write("アップデートバッチを生成します。 UpdateScriptFilePath: " & UpdateScriptFilePath)
        If UpdateScriptFilePath = "" Then
          MyLog.Write("アップデートバッチへのパスがありません。")
        Else
          Try
            Dim bat As List(Of String) = FileAccessor.Read(UpdateScriptFilePath)
            If bat.Count() = 0 OrElse bat(0) = "" Then
              Generate(UpdateScriptFilePath, releaseVersionDir, excludeFilePath)
            Else
              MyLog.Write("アップデートバッチは生成されています。")
            End If
          Catch ex As System.IO.FileNotFoundException
            Generate(UpdateScriptFilePath, releaseVersionDir, excludeFilePath)
          Catch ex As Exception
            MsgBox.ShowError(ex)
          End Try
        End If
      End Sub

      Private Sub Generate(updateFilePath As String, releaseVersionDir As String, excludeFilePath As String)
        Dim command As String = BatchCommands(releaseVersionDir, excludeFilePath)
        If command <> "" Then
          FileAccessor.Write(updateFilePath, command)
          MyLog.Write("アップデートバッチを生成しました。 command: " & command)
        Else
          MyLog.Write("アップデートバッチは生成されませんでした。")
        End If
      End Sub

      Private Function BatchCommands(releaseVersionDir As String, excludeFilePath As String) As String
        Dim comm As String = "xcopy /Y "
        Dim fromDir As String = releaseVersionDir
        Dim toDir As String = App.GetCurrentDirectory()
        Dim exclude As String = ""
        If excludeFilePath <> "" Then
          exclude = " /EXCLUDE:" & excludeFilePath
        End If

        If fromDir <> "" AndAlso toDir <> "" Then
          Return comm & fromDir & " " & toDir & exclude
        Else
          Return ""
        End If
      End Function
    End Class

    Public Class Version
      Public Shared Function IsApplicationOfLatestVersion(releaseVer As String) As Boolean
        Dim appVer As String = App.GetApplicationVersion()
        Dim appVerNums As String() = appVer.Split(".")
        Dim releaseVerNums As String() = releaseVer.Split(".")

        Dim isLatest = True

        Dim ToNum As Func(Of String, Integer) = Function(nStr) If(Char.IsDigit(nStr), Integer.Parse(nStr), 0)
        Dim verCouples As List(Of Tuple(Of Integer, Integer)) =
          appVerNums.Zip(releaseVerNums, Function(n1, n2) Tuple.Create(ToNum(n1), ToNum(n2)))

        For Each t As Tuple(Of Integer, Integer) In verCouples
          If t.Item1 > t.Item2 Then
            Exit For
          ElseIf t.Item1 < t.Item2
            isLatest = False
            Exit For
          End If
        Next


        'For i As Integer = 0 To (appVerNums.Length - 1)
        '  If i < releaseVerNums.Length AndAlso Char.IsDigit(appVerNums(i)) AndAlso Char.IsDigit(releaseVerNums(i)) Then
        '    Dim aVer As Integer = Integer.Parse(appVerNums(i))
        '    Dim rVer As Integer = Integer.Parse(releaseVerNums(i))
        '    If aVer > rVer Then
        '      Exit For
        '    ElseIf aVer < rVer Then
        '      isLatest = False
        '      Exit For
        '    End If
        '  Else
        '    Exit For
        '  End If
        'Next

        Return isLatest
      End Function
    End Class

    Public Class SerializedAccessor
      Private Serializer As ISerializer

      Sub New(serializer As ISerializer)
        Me.Serializer = serializer
      End Sub

      Public Function GetInfo(Of T)(fileName As String) As List(Of T)
        Dim list As New List(Of T)

        Try
          list =
            FileAccessor.Read(fileName).
              FindAll(Function(l) l <> "").
              ConvertAll(Function(l) Serializer.Deserialize(Of T)(l))

          'For Each line As String In FileAccessor.Read(fileName)
          '  If line <> "" Then
          '    list.Add(Serializer.Deserialize(Of T)(line))
          '  End If
          'Next
        Catch ex As System.IO.FileNotFoundException
          FileAccessor.Write(fileName, New List(Of String))
        End Try

        Return list
      End Function

      Public Sub AppendInfo(fileName As String, target As Object)
        Dim s As String = Serializer.Serialize(target)
        FileAccessor.Append(fileName, s)
      End Sub
    End Class

    Public Class MySerialize
      Private Shared FullNamespace As String = GetType(ISerializer).Namespace
      Private Shared ClassnameCSV As String = FullNamespace & ".MyCSV"
      Private Shared ClassnameJson As String = FullNamespace & ".MyJson"

      Public Shared Function GenerateAccessor() As SerializedAccessor
        Dim ver As Double = App.FrameworkVersionNumber()
        Dim type As Type = GenerateType(ver)
        Dim s As ISerializer = GenerateSerializer(type)
        Return New SerializedAccessor(s)
      End Function

      Private Shared Function GenerateType(ver As Double) As Type
        'JSONは使わない
        'If ver < 4.0 Then
        Return Type.GetType(ClassnameCSV)
        'Else
        'Return Type.GetType(ClassnameJson)
        'End If
      End Function

      Private Shared Function GenerateSerializer(type As Type) As ISerializer
        Dim target As Object = Activator.CreateInstance(type)
        Return TryCast(target, ISerializer)
      End Function

    End Class

    Class UserInfoManager
      Private UserInfoList As New List(Of Model.UserInfo)

      Sub New(userInfoList As List(Of Model.UserInfo))
        Me.UserInfoList = userInfoList
      End Sub

      Public Sub Add(userinfo As Model.UserInfo)
        UserInfoList.Add(userinfo)
      End Sub

      Public Function GetUserInfo(id As String, password As String) As Model.UserInfo
        Return UserInfoList.Find(Function(info) info.Id = id AndAlso info.Password = password)

        'Dim result As Model.UserInfo = Nothing
        'For Each info As Model.UserInfo In UserInfoList
        '  '暗号化は使わない
        '  'If info.Id = id AndAlso MyTripleDes.Decrypte(info.Password) = password Then
        '  'MessageBox.Show("id:" & info.Id & " pass:" & info.Password)
        '  If info.Id = id AndAlso info.Password = password Then
        '    result = info
        '    Exit For
        '  End If
        'Next

        'Return result
      End Function

      Public Function Certificate(id As String, password As String) As Boolean
        Return GetUserInfo(id, password) IsNot Nothing
      End Function

      Public Function Exists(id As String) As Boolean
        Return UserInfoList.Exists(Function(info) info.Id = id)
        'For Each info As Model.UserInfo In UserInfoList
        '  If info.Id = id Then
        '    Return True
        '  End If
        'Next
        'Return False
      End Function
    End Class

    'Public Module MyTripleDes
    '  Private des = New Simple3Des("dou*?,demo@.}ii===111")

    '  Public Function Encrypte(text As String) As String
    '    Return des.EncryptData(text)
    '  End Function

    '  Public Function Decrypte(text As String) As String
    '    Return des.DecryptData(text)
    '  End Function
    'End Module

  End Namespace

  Namespace MyFont
    Public Class LoadedFont
      Private MyFonts As System.Drawing.Text.PrivateFontCollection
      Private Cache As List(Of String)

      Sub New()
        MyFonts = New System.Drawing.Text.PrivateFontCollection()
        Cache = New List(Of String)
      End Sub

      Public Sub Add(fontFilePath As String)
        MyFonts.AddFontFile(fontFilePath)
      End Sub

      Public Function SearchFont(fontName As String, useFuzzySearch As Boolean) As String
        If Cache.Contains(fontName) Then
          Return fontName
        End If

        Dim check As Predicate(Of Font) = Function(f) f.Name = fontName
        If useFuzzySearch Then
          check = Function(f) f.Name.IndexOf(fontName) > 0
        End If

        Dim found As FontFamily = InstalledFontList().Find(check)
        If found IsNot Nothing Then
          Cache.Add(found.Name)
          Return found.Name
        Else
          Return ""
        End If

        'Dim fname As String = ""
        'For Each f As FontFamily In InstalledFontList()
        '  If useFuzzySearch Then
        '    If f.Name.IndexOf(fontName) >= 0 Then
        '      fName = f.Name
        '      Cache.Add(fName)
        '      Exit For
        '    End If
        '  Else
        '    If f.Name = fontName Then
        '      fName = f.Name
        '      Cache.Add(fName)
        '      Exit For
        '    End If
        '  End If
        'Next

        'Return fName
      End Function

      Public Function CreateFont(fontName As String, size As Integer) As Font
        Dim fontFamily As FontFamily = MyFonts.Families.ToList.Find(Function(ff) ff.Name = fontName)

        'Dim fontFamily As FontFamily = Nothing
        'For Each ff As System.Drawing.FontFamily In MyFonts.Families
        '  If ff.Name = fontName Then
        '    fontFamily = ff
        '    Exit For
        '  End If
        'Next

        If fontFamily IsNot Nothing Then
          'フォントオブジェクトの作成
          Return New Font(fontFamily, size, FontStyle.Regular)
        Else
          Throw New Exception(fontName & "が見つかりません。")
        End If
      End Function

      Public Function InstalledFontList() As List(Of FontFamily)
        Return New System.Drawing.Text.InstalledFontCollection().Families().ToList
      End Function

      Public Function MyFontList() As List(Of FontFamily)
        Return MyFonts.Families.ToList
      End Function
    End Class

  End Namespace

  Namespace Model
    Public Class UserInfo
      Private _Id As String
      Public Property Id() As String
        Get
          Return _Id
        End Get
        Set(value As String)
          _Id = value
        End Set
      End Property

      Private _Password As String
      Public Property Password() As String
        Get
          Return _Password
        End Get
        Set(value As String)
          _Password = value
        End Set
      End Property

      Private _Name As String
      Public Property Name() As String
        Get
          Return _Name
        End Get
        Set(value As String)
          _Name = value
        End Set
      End Property

      Sub New()
      End Sub

      Sub New(id As String, password As String, name As String)
        Me.Id = id
        Me.Password = password
        Me.Name = name
      End Sub

      Function GetIdNum() As String
        If Id.Length >= 3 Then
          Return Id.Substring(Id.Length - 3)
        Else
          Throw New Exception("IDが不正です。 / " + Id)
        End If
      End Function

    End Class
  End Namespace

End Namespace
