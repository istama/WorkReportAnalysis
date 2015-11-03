
Imports System.Security.Cryptography
Imports System.Runtime.Serialization
Imports System.Reflection
Imports System.IO

Namespace Details

  Namespace Serialize
    Public Interface ISerializer
      Function Serialize(target As Object) As String
      Function Deserialize(Of T)(str As String) As T
    End Interface

    Public Class MyCSV
      Implements ISerializer

      Public Function Serialize(target As Object) As String Implements ISerializer.Serialize
        Dim type As Type = target.GetType

        Dim result As String = ""
        For Each m As MemberInfo In GetGetter(type)
          If result.Length <> 0 Then
            result = result & ","
          End If
          result = result & CStr(type.InvokeMember(m.Name, BindingFlags.InvokeMethod, Nothing, target, Nothing))
        Next
        Return result
      End Function

      Public Function Deserialize(Of T)(csv As String) As T Implements ISerializer.Deserialize
        Dim type As Type = GetType(T)

        Dim members As MemberInfo() = GetSetter(type)
        Dim elems As String() = csv.Split(",")

        If members.Length = elems.Length Then
          Dim target As Object = type.InvokeMember(Nothing, BindingFlags.CreateInstance, Nothing, Nothing, New Object() {})

          For i As Integer = 0 To elems.Length - 1
            Dim m As MemberInfo = members(i)
            Dim e As String = elems(i)
            If m.MemberType = Reflection.MemberTypes.Method Then
              type.InvokeMember(m.Name, BindingFlags.InvokeMethod, Nothing, target, New Object() {e})
            End If
          Next

          Return target
        Else
          Throw New Exception("クラスフィールドの数とCVSの要素数が合いません。")
        End If
      End Function

      Private Function GetGetter(t As Type) As MemberInfo()
        Return GetProperties(t, "get_")
      End Function

      Private Function GetSetter(t As Type) As MemberInfo()
        Return GetProperties(t, "set_")
      End Function

      Private Function GetProperties(t As Type, prefix As String) As MemberInfo()
        Dim members As MemberInfo() = t.GetMembers()

        Dim ml As New List(Of MemberInfo)
        For Each m As MemberInfo In members
          If m.MemberType = Reflection.MemberTypes.Method AndAlso m.Name.IndexOf(prefix) = 0 Then
            ml.Add(m)
          End If
        Next

        Return ml.ToArray
      End Function

    End Class
  End Namespace

  Namespace IO
    Public Class PropertyAccessor
      Private Shared SEPARATOR As Char = "="

      Public Shared Function GetProp(fileName As String) As IDictionary(Of String, String)
        Dim pp As IDictionary(Of String, ArrayList) = GetPropDuplicatedKeys(fileName)

        Dim p As New Dictionary(Of String, String)
        For Each k In pp.Keys
          p.Add(k, pp(k)(0))
        Next

        Return p
      End Function

      Public Shared Function GetPropDuplicatedKeys(fileName As String) As IDictionary(Of String, ArrayList)
        Dim p As New Dictionary(Of String, ArrayList)
        Dim texts As List(Of String) = FileAccessor.Read(fileName)

        For Each t As String In texts
          Dim idx As Integer = t.IndexOf(SEPARATOR)
          If idx > 0 Then
            Dim key As String = t.Substring(0, idx)
            Dim value As String = t.Substring(idx + 1)
            If Not p.ContainsKey(key) Then
              p.Add(key, New ArrayList)
            End If
            p(key).Add(value)
          End If
        Next

        Return p
      End Function

      Public Shared Sub SetProp(fileName As String, p As IDictionary(Of String, String))
        FileAccessor.Write(fileName, DictToPropList(p, Nothing))
      End Sub

      Public Shared Sub AppendProp(fileName As String, key As String, value As String)
        Dim l As New List(Of String)(New String() {ToPropString(key, value)})
        FileAccessor.Append(fileName, l)
      End Sub

      Public Shared Sub AppendProp(fileName As String, p As IDictionary(Of String, String))
        FileAccessor.Append(fileName, DictToPropList(p, Nothing))
      End Sub

      Public Shared Sub AppendOnlyPropThatDoesNotExists(fileName As String, p As IDictionary(Of String, String))
        Dim keys As ICollection(Of String) = GetPropDuplicatedKeys(fileName).Keys
        'Dim k As IDictionary(Of String, ArrayList) = GetPropDuplicatedKeys(fileName).Keys.ToList
        FileAccessor.Append(fileName, DictToPropList(p, keys))
      End Sub

      Private Shared Function DictToPropList(dict As IDictionary(Of String, String), omittedKeys As ICollection(Of String)) As List(Of String)
        Dim l As New List(Of String)

        For Each k As String In dict.Keys
          If omittedKeys Is Nothing OrElse Not omittedKeys.Contains(k) Then
            l.Add(ToPropString(k, dict(k)))
          End If
        Next

        Return l
      End Function

      Private Shared Function ToPropString(key As String, value As String) As String
        Return key & SEPARATOR & value
      End Function
    End Class

    Public Class FileAccessor
      Public Shared Function Read(fileName As String) As List(Of String)
        Dim op As New Op()
        Dim f As Operate = AddressOf op.Input
        Return Access(fileName, OpenMode.Input, f)
      End Function

      Public Shared Sub Write(fileName As String, text As String)
        Dim l As New List(Of String)(New String() {text})
        Write(fileName, l, OpenMode.Output)
      End Sub

      Public Shared Sub Write(fileName As String, text As List(Of String))
        Write(fileName, text, OpenMode.Output)
      End Sub

      Public Shared Sub Append(fileName As String, text As String)
        Dim l As New List(Of String)(New String() {text})
        Write(fileName, l, OpenMode.Append)
      End Sub

      Public Shared Sub Append(fileName As String, text As List(Of String))
        Write(fileName, text, OpenMode.Append)
      End Sub

      Private Shared Sub Write(fileName As String, text As List(Of String), mode As Integer)
        Dim op As New Op(text)
        Dim f As Operate = AddressOf op.Output
        Access(fileName, mode, f)
      End Sub

      Private Shared Function Access(fileName As String, mode As Integer, op As Operate) As List(Of String)
        Dim text As List(Of String)
        Dim fh As Integer = FreeFile()
        Try
          FileOpen(fh, fileName, mode)
          text = op(fh)
        Finally
          FileClose(fh)
        End Try

        Return text
      End Function

      Private Class Op
        Private WrittingText As New List(Of String)

        Sub New()
        End Sub

        Sub New(text As List(Of String))
          Me.WrittingText = text
        End Sub

        Function Output(fh As Integer) As List(Of String)
          WrittingText.ForEach(Sub(t) PrintLine(fh, t))
          'For Each t As String In WrittingText
          '  PrintLine(fh, t)
          'Next

          Return Nothing
        End Function

        Function Input(fh As Integer) As List(Of String)
          Dim list As New List(Of String)

          Do While EOF(fh) = False
            Dim l As String = LineInput(fh)
            list.Add(l)
          Loop

          Return list
        End Function
      End Class

      Delegate Function Operate(fh As Integer) As List(Of String)
    End Class

  End Namespace

  Namespace Sys
    Public Class App
      Public Shared Function FrameworkVersionNumber() As Double
        Dim verStr As String = System.Reflection.Assembly.GetExecutingAssembly().ImageRuntimeVersion
        Return Double.Parse(verStr.Substring(1, 3))
      End Function

      Public Shared Function GetApplicationVersion() As String
        Dim ver As FileVersionInfo = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location)
        Return ver.FileVersion
      End Function

      Public Shared Function GetCurrentDirectory() As String
        Return System.IO.Directory.GetCurrentDirectory()
      End Function
    End Class
  End Namespace

  'Namespace Security
  '  NotInheritable Class Simple3Des
  '    Private TripleDes As New TripleDESCryptoServiceProvider

  '    Sub New(ByVal key As String)
  '      ' Initialize the crypto provider.
  '      TripleDes.Key = TruncateHash(key, TripleDes.KeySize \ 8)
  '      TripleDes.IV = TruncateHash("", TripleDes.BlockSize \ 8)
  '    End Sub

  '    Private Function TruncateHash(key As String, length As Integer) As Byte()
  '      Dim sha1 As New SHA1CryptoServiceProvider

  '      ' Hash the key. 
  '      Dim keyBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(key)
  '      Dim hash() As Byte = sha1.ComputeHash(keyBytes)

  '      ' Truncate or pad the hash. 
  '      ReDim Preserve hash(length - 1)
  '      Return hash
  '    End Function

  '    Public Function EncryptData(plaintext As String) As String
  '      ' Convert the plaintext string to a byte array. 
  '      Dim plaintextBytes() As Byte = System.Text.Encoding.Unicode.GetBytes(plaintext)

  '      ' Create the stream. 
  '      Dim ms As New System.IO.MemoryStream
  '      ' Create the encoder to write to the stream. 
  '      Dim encStream As New CryptoStream(ms, TripleDes.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

  '      ' Use the crypto stream to write the byte array to the stream.
  '      encStream.Write(plaintextBytes, 0, plaintextBytes.Length)
  '      encStream.FlushFinalBlock()

  '      ' Convert the encrypted stream to a printable string. 
  '      Return Convert.ToBase64String(ms.ToArray)
  '    End Function

  '    Public Function DecryptData(encryptedtext As String) As String
  '      ' Convert the encrypted text string to a byte array. 
  '      Dim encryptedBytes() As Byte = Convert.FromBase64String(encryptedtext)

  '      ' Create the stream. 
  '      Dim ms As New System.IO.MemoryStream
  '      ' Create the decoder to write to the stream. 
  '      Dim decStream As New CryptoStream(ms, TripleDes.CreateDecryptor(), System.Security.Cryptography.CryptoStreamMode.Write)

  '      ' Use the crypto stream to write the byte array to the stream.
  '      decStream.Write(encryptedBytes, 0, encryptedBytes.Length)
  '      decStream.FlushFinalBlock()

  '      ' Convert the plaintext stream to a string. 
  '      Return System.Text.Encoding.Unicode.GetString(ms.ToArray)
  '    End Function
  '  End Class
  'End Namespace
End Namespace