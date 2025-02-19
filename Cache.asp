<%
'clsApplicationCache Class @0-BAD95BE3
Class clsApplicationCache

' Cache uses Application variables to store 3-element arrays
' (key, number of hits, object or variable)
' PREFIX - Default prefix for cache elements
' PREFIX_LENGTH - Cached elements length

Public PREFIX
Public PREFIX_LENGTH

Private Sub Class_Initialize()
  PREFIX = "Cached:"
  PREFIX_LENGTH = 7
End Sub

' Release all cached elements from the cache
Sub Clear()
  Dim Key, Keys, KeyIndex
  Dim KeysString : KeysString = ""
  For Each Key in Application.Contents
    If Left(Key, PREFIX_LENGTH) = PREFIX Then KeysString = KeysString & "/" & Key
  Next
  Keys = Split(KeysString, "/")
  Application.Unlock 
  For KeyIndex = 1 To UBound(Keys)
    Application.Contents.Remove(Keys(KeyIndex))
  Next
  Application.Lock
End Sub

' Release elements with zero hits from the cache.
' (hits are stored in the second array element of the cache record).
Sub ClearUnused()
  Dim Key, Keys, KeyIndex
  Dim KeysString : KeysString = ""
  For Each Key in Application.Contents
    If Left(Key, PREFIX_LENGTH) = PREFIX Then 
      If Application(Key)(2) = 0 Then KeysString = KeysString & "/" & Key
    End If
  Next
  Keys = Split(KeysString, "/")
  Application.Unlock 
  For KeyIndex = 1 To UBound(Keys)
    Application.Contents.Remove(Keys(KeyIndex))
  Next
  Application.Lock
End Sub

' Reset all cache hit counters
Sub ResetCounters()
  Dim Key
  Dim Elements
  Application.Unlock
  For Each Key in Application.Contents
    If Left(Key, PREFIX_LENGTH) = PREFIX Then
      Elements = Application(Key)
      Elements(2) = 0
      Application(Key) = Elements
    End If
  Next
  Application.Lock
End Sub

' Retrieve the hit count for an object or variable
Function GetCounter(Key)
  Dim Elements
  Dim Result : Result = Empty
  Elements = Application(PREFIX & Key)
  If IsArray(Elements) Then Result = Elements(2)
  GetCounter = Result
End Function

' Check if a non-empty object exists within the cache
Function CacheItemExists(Key)
  Dim Elements
  Dim Result : Result = false
  Elements = Application(PREFIX & Key)
  If IsArray(Elements) Then
    Result = True
  End If
  CacheItemExists = Result
End Function

' Create a new Cache member
Sub Put(Key, CachingObject)
  Dim Elements(3)
  Elements(2) = 0
  If IsObject(CachingObject) Then 
    Set Elements(3) = CachingObject
  Else
    Elements(3) = CachingObject
  End If
  Application.Unlock
  Application(PREFIX & Key) = Elements
  Application.Lock
End Sub

' Retrieve cached object or variable
Function GetCachedElement(Key)
  Dim Elements
  Elements = Application(PREFIX & Key)
  If IsArray(Elements) Then
    If IsObject(Elements(3)) Then 
      Set GetCachedElement = Elements(3)
    Else
      GetCachedElement = Elements(3)
    End If
  Else
    GetCachedElement = Empty
  End If
End Function

End Class
'End clsApplicationCache Class

'FileSystem Object @0-AD64CEF7
Class clsCache_FileSystem
  Private FileSystem
  Private FSCache
  Private mEncoding

  Private Sub Class_Initialize()
    Set FileSystem = Server.CreateObject("Scripting.FileSystemObject")
    Set FSCache = Server.CreateObject("Scripting.Dictionary")
  End Sub

  Function ItemExists(name)
    ItemExists = FileSystem.FileExists(name)
  End Function

  Public Default Property Get Items(name)
    Dim res
    If FSCache.Exists(name) Then
      Set res = FSCache(name)
    Else
      Set res = New clsCacheItem_File
      res.Name = name
      res.Encoding = Encoding
      Set res.FSO = FileSystem
      Set FSCache(Name) = res
    End If
    Set Items = res
  End Property

  Private Sub Class_Terminate()
    Set FileSystem = Nothing
  End Sub

  Property Get CacheType()
    CacheType = ccsCacheHTML
  End Property

  Property Get Encoding()
    Encoding = mEncoding
  End Property

  Property Let Encoding(newEncoding)
    mEncoding = newEncoding
  End Property

End Class
'End FileSystem Object

'File Object @0-B7347821
Class clsCacheItem_File
  Private mContent
  Private IsOpen
  Private mFSO
  Private mName
  Private mEncoding

  Private Sub Class_Initialize()
    IsOpen = False
    Set mFSO = Nothing
  End Sub

  Private Sub Class_Terminate()
    Set mFSO = Nothing
  End Sub
  
  Public Default Property Get Content()
    If Not IsOpen Then
      mContent = GetFileContent(mFSO, mName)
      IsOpen = True
    End If
    Content = mContent
  End Property

  Property Get Name()
    Name = mName
  End Property

  Property Let Name(newName)
    If Not IsOpen Then _
      mName = newName
  End Property

  Property Get Encoding()
    Encoding = mEncoding
  End Property

  Property Let Encoding(newEncoding)
    mEncoding = newEncoding
  End Property

  Property Set FSO(newFSO)
    Set mFSO = newFSO
  End Property

  Sub Reset()
    IsOpen = False
  End Sub

  Private Function GetFileContent(FileSystem, FilePath)
    Dim Strm

    If Not (FileSystem Is Nothing Or FilePath = "") Then
      If FileSystem.FileExists(FilePath) Then
        If IsEmpty(mEncoding) Then 
          Set Strm = FileSystem.OpenTextFile(FilePath, ccsOpenForReading)
          If Strm.AtEndOfStream Then _
            GetFileContent = "" _
          Else _
            GetFileContent = Strm.ReadAll
          Set Strm = Nothing
        Else
          Set Strm = Server.CreateObject("ADODB.Stream")
          Strm.Open
          Strm.Charset = mEncoding
          Strm.LoadFromFile FilePath
          GetFileContent = Strm.ReadText(adReadAll)
          Strm.Close
          Set Strm = Nothing
        End If
        If Not CCSLocales.UseStaticTranslation Then
          GetFileContent = CCRegExpReplace(GetFileContent, "(<meta\s+http-equiv=['""]?content-type['""]?\s+content=['""]?text/html;\s*charset=)([-a-z0-9]+)(['""]?[^>]*>)", "$1" & CCSLocales.Locale.Charset & "$3", True) 
          GetFileContent = CCRegExpReplace(GetFileContent, "(<meta\s+http-equiv=['""]?content-type['""]?\s+content=['""]?text/html;\s*charset=)([-a-z0-9]+)(['""]?[^>]*>)", "$1" & CCSLocales.Locale.Charset & "$3", True) 
        End If 
        GetFileContent = Replace(GetFileContent,"{CCS_Style}", CCSStyle)
      Else
        Err.raise 1052, "File object", "File object: GetFileContent failed. Type missmatch."
      End If
    Else
      Err.raise 1051, "File object", "File object: GetFileContent failed. File " & FilePath & " is not found."
    End If
  End Function

End Class
'End File Object


%>
