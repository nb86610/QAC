<%


'Template class @0-88533A01

Const tpParse = False
Const tpParseSafe = True

Const ccsCacheHTML  = 1
Const ccsCacheArray = 2 ' Reserved
Const ccsCacheXML   = 3 ' Reserved

Const ccsOpenForReading = 1

Class clsTemplate
  
Public HEADContent ' String
Public Tree        ' Array
Public TreePaths   ' Array
Public MakeCleanup ' Boolean
Public InsertHEADContent'Boolean
Public Encoding
Private TreePath    ' String
Private TreeSize    ' Integer
Private mCache      ' Object

Private LONG_MAX_VALUE

Private BEGIN_OPEN
Private BEGIN_CLOSE
Private BEGIN_OPEN_LENGTH
Private BEGIN_CLOSE_LENGTH

Private END_OPEN
Private END_CLOSE
Private END_OPEN_LENGTH
Private END_CLOSE_LENGTH

Private BLOCK_PATH 
Private BLOCK_TYPE 
Private BLOCK_VALUE

Private TYPE_VARIABLE
Private TYPE_BEGIN_BLOCK
Private TYPE_END_BLOCK
Private TYPE_TEXT

Private Sub Class_Initialize()
  HEADContent = ""
  MakeCleanup = False
  InsertHEADContent = False
  Tree = Array(0)
  TreePath = ""
  Set mCache = Nothing

  LONG_MAX_VALUE = 2147483647

  BEGIN_OPEN = "<!-- BEGIN "
  BEGIN_CLOSE = " -->"
  BEGIN_OPEN_LENGTH = 11
  BEGIN_CLOSE_LENGTH = 4

  END_OPEN = "<!-- END "
  END_CLOSE = " -->"
  END_OPEN_LENGTH = 8
  END_CLOSE_LENGTH = 4

  BLOCK_PATH = 0 
  BLOCK_TYPE = 1 
  BLOCK_VALUE = 2

  TYPE_VARIABLE = 0
  TYPE_BEGIN_BLOCK = 1
  TYPE_END_BLOCK = 2
  TYPE_TEXT = 3

End Sub

Property Set Cache(newCache)
  Set mCache = newCache
  mCache.Encoding=Encoding
End Property

Sub LoadTemplate(FilePath)
  Dim TemplateArrays

  If Not (mCache Is Nothing) Then
    If mCache.ItemExists(FilePath) Then
      Select Case mCache.CacheType
        Case ccsCacheHTML
          TemplateArrays = BuildArraysFromHTML(mCache.Items(FilePath).Content)
      End Select
      Tree = TemplateArrays(0) : TreePaths = TemplateArrays(1)
    Else
      Err.raise 1050, "Template engine", "Template engine: LoadTemplate failed. File " & FilePath & " not found."
    End If
  Else
    Err.raise 1050, "Template engine", "Template engine: LoadTemplate failed. Template repository is not set."
  End If
End Sub

Sub LoadTemplateFromStr(Str)
  Dim TemplateArrays
  TemplateArrays = BuildArraysFromHTML(Str)
  Tree = TemplateArrays(0) 
  TreePaths = TemplateArrays(1)
End Sub

Function BuildArraysFromHTML(FileContent)
  Dim TreePaths(), Tree()
  Dim TplVarName
  Dim RootName : RootName = "main"

  Dim CurrentPosition : CurrentPosition = 1
' Brackets - two-dimensional array contains tag and block positions.
' 0 - Curly bracket {}
' 1 - Begin block <!-- BEGIN block -->
' 2 - End block <!-- END block -->
' Array elements:
' 0 - name, 1 - beginning tag bracket/block, 2 - ending tag bracket/block, 3 - number of template elements
  Const BLOCK_NAME = 0
  Const BLOCK_BEGIN = 1
  Const BLOCK_END = 2
  Const BLOCK_AMOUNT = 3

  Dim Brackets(3)
  Dim Counters(3), I
  Dim CurrentNode : CurrentNode = 0
  Dim Values(3)
  Dim BlockType


  Brackets(0) = GetRegExp("\{([A-z]+\d*\:*( [A-z]+\d*)*)+\}", FileContent)
  Brackets(1) = GetRegExp(BEGIN_OPEN & "[^-]*" & BEGIN_CLOSE, FileContent)
  Brackets(2) = GetRegExp(END_OPEN & "[^-]*" & END_CLOSE, FileContent)

  For I = 0 to 2 
    If Brackets(I)(BLOCK_AMOUNT) > 0 Then
      Values(I) = Brackets(I)(BLOCK_BEGIN)(0) 
    Else
      Values(I) = LONG_MAX_VALUE
    End If
    Counters(I) = 0
  Next

  Dim Paths(30), PathLength : PathLength = 0 : Paths(0) = RootName
  CurrentNode = CurrentNode + 1
  ReDim Preserve Tree(CurrentNode)
  Tree(CurrentNode - 1) = Array("/" & RootName & "/*", TYPE_BEGIN_BLOCK, "")

  BlockType = MinValue(Values)
  While Not (BlockType = LONG_MAX_VALUE)
    If (Brackets(BlockType)(BLOCK_BEGIN)(Counters(BlockType)) - CurrentPosition + 1) > 0 Then 
      CurrentNode = CurrentNode + 1
      ReDim Preserve Tree(CurrentNode)
      Tree(CurrentNode - 1) = Array(GetPath(Paths, "@text"), TYPE_TEXT, Mid(FileContent, CurrentPosition, Brackets(BlockType)(BLOCK_BEGIN)(Counters(BlockType)) - CurrentPosition + 1))
      CurrentPosition = Brackets(BlockType)(BLOCK_END)(Counters(BlockType)) + 1
    else
      CurrentPosition = Brackets(BlockType)(BLOCK_END)(Counters(BlockType)) + 1
    End If

    CurrentNode = CurrentNode + 1
    ReDim Preserve Tree(CurrentNode)
  
    Select Case BlockType
      Case TYPE_VARIABLE
        TplVarName = GetName(Brackets(BlockType)(BLOCK_NAME)(Counters(BlockType)), BlockType)
        Tree(CurrentNode - 1) = Array(GetPath(Paths, TplVarName) & "/*", BlockType, "")
        If Left(TplVarName,5) = "@res:" Then 
          Tree(CurrentNode - 1)(BLOCK_VALUE) = CCSLocales.GetText(Right(TplVarName, Len(TplVarName) - 5), "")
        End If
     Case TYPE_BEGIN_BLOCK
        PathLength = PathLength + 1
        Paths(PathLength) = GetName(Brackets(BlockType)(BLOCK_NAME)(Counters(BlockType)), BlockType)
        Tree(CurrentNode - 1) = Array(GetPath(Paths, Empty) & "*", BlockType, "")
      Case TYPE_END_BLOCK
        Tree(CurrentNode - 1) = Array(GetPath(Paths, Empty) & "*", BlockType, "")
        Paths(PathLength) = ""
        PathLength = PathLength - 1
    End Select

    Counters(BlockType) = Counters(BlockType) + 1
    If Counters(BlockType) = Brackets(BlockType)(BLOCK_AMOUNT) Then 
      Values(BlockType) = LONG_MAX_VALUE
    Else 
      Values(BlockType) = Brackets(BlockType)(BLOCK_BEGIN)(Counters(BlockType))
    end if

    BlockType = MinValue(Values)
  Wend

  If CurrentPosition < Len(FileContent) Then
    CurrentNode = CurrentNode + 1
    ReDim Preserve Tree(CurrentNode)
    Tree(CurrentNode - 1) = Array("/" & RootName & "/@text", TYPE_TEXT, Right(FileContent, Len(FileContent) - CurrentPosition + 1))
  End If

  CurrentNode = CurrentNode + 1
  ReDim Preserve Tree(CurrentNode)
  Tree(CurrentNode - 1) = Array("/" & RootName & "/*", TYPE_END_BLOCK, "")

  TreeSize = UBound(Tree) - 1
  ReDim TreePaths(TreeSize + 1)
  For I = 0 To TreeSize
    TreePaths(I) = Tree(I)(BLOCK_PATH) & ":" & I
  Next
  BuildArraysFromHTML = Array(Tree, TreePaths)
End Function

Sub PrintBlocks()
  Dim I
  Response.Write "<table border>"
  For I = 0 To TreeSize
    Response.Write "<tr>"
    Response.Write "<td>" & Tree(I)(BLOCK_PATH) & "</td>"
    Response.Write "<td>" & Tree(I)(BLOCK_TYPE) & "</td>"
    Response.Write "<td>" & Server.HTMLEncode(Tree(I)(BLOCK_VALUE)) & "</td>"
    Response.Write "<td>" & TreePaths(I) & "</td>"
    Response.Write "</tr>"
  Next
  Response.Write "</table>"
End Sub

Sub PrintArray()
  Dim I, J
  Response.Write "<table border=""1"">"
  For I = 0 To TreeSize
    Response.Write "<tr>"
    For J = 0 To UBound(Tree(I))
      Response.Write "<td>" & Server.HTMLEncode(Tree(I)(J)) & "</td>"
    Next
    Response.Write "</tr>"
  Next
  Response.Write "</table>"
End Sub

  Sub SetPath(NewPath)
    TreePath = NewPath
  End Sub

  Property Let Path(NewPath)
    TreePath = NewPath
  End Property

  Property Get Path()
    Path = TreePath
  End Property


Sub SetVar(VariableName, Value)
  Dim VariablePaths, I, VariablePosition, VariablePath
  VariablePaths = Filter(TreePaths, "/" & TreePath & VariableName & "/*")
  For I = 0 To UBound(VariablePaths)
    VariablePath = VariablePaths(i)
    VariablePosition = CLng(Mid(VariablePath, InStrRev(VariablePath, ":") + 1))
    If Tree(VariablePosition)(BLOCK_TYPE) = TYPE_VARIABLE or Tree(VariablePosition)(BLOCK_TYPE) = TYPE_BEGIN_BLOCK Then Tree(VariablePosition)(BLOCK_VALUE) = Value
  Next
End Sub

Function GetVar(VariableName)
  Dim VariablePaths, I, VariablePosition, VariablePath, Result, TotalVariables
  VariablePaths = Filter(TreePaths, "/" & TreePath & VariableName & "/*")
  TotalVariables = UBound(VariablePaths)
  If TotalVariables > 0 Then
    If (Tree(VariablePosition)(BLOCK_TYPE) = TYPE_VARIABLE And TotalVariables > 0) Or _
      (Tree(VariablePosition)(BLOCK_TYPE) = TYPE_BEGIN_BLOCK And TotalVariables > 1) Then
      ' ERROR
    Else
      VariablePath = VariablePaths(0)
      VariablePosition = CLng(Mid(VariablePath, InStrRev(VariablePath, ":") + 1))
      Result = Tree(VariablePosition)(BLOCK_VALUE)
    End If
  Else
    Result = ""
  End If
  GetVar = Result
End Function

Sub SetBlockVariable(BeginBlockIndex, EndBlockIndex, VariableName, Value)
   Dim I
   For I = BeginBlockIndex To EndBlockIndex
     If Tree(I)(BLOCK_TYPE) = TYPE_VARIABLE And InStr(Tree(I)(BLOCK_PATH),VariableName) > 0 Then Tree(I)(BLOCK_VALUE) = Value
   Next
End Sub

Function GetHTML(BlockName)
  Dim BlockPaths, I, BlockPath
  Dim BeginBlockIndex, EndBlockIndex, TargetIndex
  Dim BlockContent, TreeDeep
  TreeDeep = 0
  BlockPaths = Filter(TreePaths, "/" & TreePath & BlockName & "/*")

  If ((UBound(BlockPaths) - LBound(BlockPaths)) <> 1) Then
    Err.Raise 1050, "Template Engine.", "Parsing function: Block """ & BlockName & """ cannot be found or selected."
  Else
    BlockPath = BlockPaths(0)
    BeginBlockIndex = CLng(Mid(BlockPath, InStrRev(BlockPath, ":") + 1))
  End If

  GetHTML = Tree(BeginBlockIndex)(BLOCK_VALUE)
End Function


Sub ParseAndPrint(BlockName, Accumulate, Output, TargetBlock, SafeParse)
  Dim BlockPaths, I, BlockPath
  Dim BeginBlockIndex, EndBlockIndex, TargetIndex
  Dim BlockContent, TreeDeep
  TreeDeep = 0
  BlockPaths = Filter(TreePaths, "/" & TreePath & BlockName & "/*")

  If ((UBound(BlockPaths) - LBound(BlockPaths)) <> 1) Then
    If Not SafeParse Then _
      Err.Raise 1050, "Template Engine.", "Parsing function: Block """ & BlockName & """ cannot be found or selected."
  Else
    BlockPath = BlockPaths(0)
    BeginBlockIndex = CLng(Mid(BlockPath, InStrRev(BlockPath, ":") + 1))
    BlockPath = BlockPaths(1)
    EndBlockIndex = CLng(Mid(BlockPath, InStrRev(BlockPath, ":") + 1))

    For i = BeginBlockIndex + 1 To EndBlockIndex - 1
      If Tree(i)(BLOCK_TYPE) = TYPE_BEGIN_BLOCK Then 
        TreeDeep = TreeDeep + 1
        If TreeDeep = 1 Then BlockContent = BlockContent & Tree(i)(BLOCK_VALUE)
      ElseIf Tree(i)(BLOCK_TYPE) = TYPE_END_BLOCK Then 
        TreeDeep = TreeDeep - 1
      ElseIf TreeDeep = 0 Then
        BlockContent = BlockContent & Tree(i)(BLOCK_VALUE) 
      End If
    Next

    If IsEmpty(TargetBlock) Then
      TargetIndex = BeginBlockIndex
    Else
      BlockPaths = Filter(TreePaths, "/" & TreePath & TargetBlock & "/*")
      BlockPath = BlockPaths(0)
      TargetIndex = CLng(Mid(BlockPath, InStrRev(BlockPath, ":") + 1))    
    End If

    If Accumulate Then
      Tree(TargetIndex)(BLOCK_VALUE) = Tree(TargetIndex)(BLOCK_VALUE) & BlockContent
    Else
      Tree(TargetIndex)(BLOCK_VALUE) = BlockContent
    End If

    If Output Then Response.Write Tree(BeginBlockIndex)(BLOCK_VALUE)
  End If
End Sub

Sub ParseBlockByIndex(BeginBlockIndex, EndBlockIndex, Accumulate)
  Dim i, BlockContent, TreeDeep
  TreeDeep = 0
  For i = BeginBlockIndex + 1 To EndBlockIndex - 1
    If Tree(i)(BLOCK_TYPE) = TYPE_BEGIN_BLOCK Then 
      TreeDeep = TreeDeep + 1
      If TreeDeep = 1 Then BlockContent = BlockContent & Tree(i)(BLOCK_VALUE)
    ElseIf Tree(i)(BLOCK_TYPE) = TYPE_END_BLOCK Then 
      TreeDeep = TreeDeep - 1
    ElseIf TreeDeep = 0 Then
      BlockContent = BlockContent & Tree(i)(BLOCK_VALUE) 
    End If
  Next

  If Accumulate Then
    Tree(BeginBlockIndex)(BLOCK_VALUE) = Tree(BeginBlockIndex)(BLOCK_VALUE) & BlockContent
  Else
    Tree(BeginBlockIndex)(BLOCK_VALUE) = BlockContent
  End If
End Sub

Function BlockExists(BlockName, BlockType)
  Dim BlockPaths, Result : Result = False
  BlockPaths = Filter(TreePaths, "/" & TreePath & BlockName & "/*")

  If BlockType = "block" Then
    If (UBound(BlockPaths) - LBound(BlockPaths)) = 1 Then _
      Result = True
  ElseIf BlockType = "variable" Then
    If (UBound(BlockPaths) - LBound(BlockPaths)) = 0 Then _
      Result = True
  Else
    Err.Raise 1050, "Template library", "BlockExists function: Invalid BlockType parameter."
  End If
  BlockExists = Result
End Function

Sub HideBlock(BlockName)
  SetVar BlockName, ""
End Sub

Sub Parse(BlockName, Accumulate)
  ParseAndPrint BlockName, Accumulate, False, Empty, tpParse
End Sub

Sub PParse(BlockName, Accumulate)
  ParseAndPrint BlockName, Accumulate, True, Empty, tpParse
End Sub

Sub ParseTo(BlockName, Accumulate, TargetBlock)
  ParseAndPrint BlockName, Accumulate, False, TargetBlock, tpParse
End Sub

Sub ParseSafe(BlockName, Accumulate)
  ParseAndPrint BlockName, Accumulate, False, Empty, tpParseSafe
End Sub

Sub ParseSafeTo(BlockName, Accumulate, TargetBlock)
  ParseAndPrint BlockName, Accumulate, False, TargetBlock, tpParseSafe
End Sub

Function GetPath(Paths, PathAdding)
  Dim Path : Path = Join(Paths, "/")
  GetPath = "/" & Left(Path, InStr(Path, "//")) & PathAdding
End Function

Function GetFormattedTree()
  Dim Result : Result = ""
  Dim I
  For I = 0 To UBound(Tree)
    If IsArray(Tree(I)) Then 
      Result = Result & Tree(I)(BLOCK_PATH) & "&nbsp;"
    Else 
      Result = Result & Tree(I) & "&nbsp;"
    End If
  Next
End Function

Function GetName(BlockString, BlockType)
  Select Case BlockType
    Case TYPE_VARIABLE:
      GetName = "@" & mid(BlockString, 2, len(BlockString) - 2)
    Case TYPE_BEGIN_BLOCK:
      GetName = mid(BlockString, BEGIN_OPEN_LENGTH + 1, len(BlockString) - BEGIN_CLOSE_LENGTH - BEGIN_OPEN_LENGTH)
    Case TYPE_END_BLOCK:
      GetName = mid(BlockString, END_OPEN_LENGTH + 1, len(BlockString) - END_CLOSE_LENGTH - END_OPEN_LENGTH)
  End Select
End Function

Function SetBlockHTML(Block, HTML)
  Tree(Block.BeginOfBlock)(BLOCK_VALUE) = HTML
End Function

Function GetBlockHTML(Block)
  GetBlockHTML = Tree(Block.BeginOfBlock)(BLOCK_VALUE)
End Function

Function CleanHTML(HTML)
  Dim re, result, matches, match
  result = HTML
  If MakeCleanup Then
    Set re = New RegExp
    re.MultiLine = False
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "<meta[^>]*>"
    result = re.Replace(HTML, "")
    re.Pattern = "<!--\s*\[if[^\]]+\]>\s*<link[^>]+>\s*<!\[endif\]\s*-->|<!--\s*\[if[^\]]+\]><!\s*-->\s*<link[^>]+>\s*<!--\s*<!\[endif\]\s*-->|<!\[if[^\]]+\]>\s*<link[^>]+>\s*<!\[endif\]>"
    Set matches = re.Execute(HTML)
    For Each match In matches
      HEADContent = HEADContent & match.Value & vbLf
      result = replace(result, match.Value, "")
    Next
    re.Pattern = "<link[^>]+>"
    Set matches = re.Execute(HTML)
    For Each match In matches
      HEADContent = HEADContent & match.Value & vbLf
      result = replace(result, match.Value, "")
    Next
  End If
  If InsertHEADContent Then
    Set re = New RegExp
    re.MultiLine = False
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "(</\s*head\s*>)"
    result = re.Replace(HTML, HEADContent & "$1")
  End If
  CleanHTML = result
End Function

Function MinValue(Values)

  Dim MinimumType, MinimumValue, I 
  MinimumType = LONG_MAX_VALUE  
  MinimumValue = LONG_MAX_VALUE  

  For I = 0 To 2
    If Values(I) < MinimumValue Then
      MinimumValue = Values(I)
      MinimumType = I
    End If 
  Next

  MinValue = MinimumType
End Function

Function GetRegExp(RegExpPattern, FileContent)
  Dim MatchesValues(), MatchesBeginIndexes(), MatchesEndIndexes()
  Dim RegExpObject, Matches, TotalMatches, I

  Set RegExpObject = New RegExp
  RegExpObject.Pattern = RegExpPattern
  RegExpObject.IgnoreCase = True
  RegExpObject.Global = True
  Set Matches = RegExpObject.Execute(FileContent)
  Set RegExpObject = Nothing
  If Matches.Count > 0 Then
    TotalMatches = Matches.Count
    ReDim MatchesValues(TotalMatches)
    ReDim MatchesBeginIndexes(TotalMatches)
    ReDim MatchesEndIndexes(TotalMatches)
  Else
    TotalMatches = 0
  End If
  For I = 0 To TotalMatches - 1
    MatchesValues(I) = Matches.Item(I).Value
    MatchesBeginIndexes(I) = Matches.Item(I).FirstIndex
    MatchesEndIndexes(I) = Matches.Item(I).FirstIndex + Len(Matches.Item(I).Value)
  Next
  GetRegExp = Array(MatchesValues, MatchesBeginIndexes, MatchesEndIndexes, TotalMatches)
End Function

Function Block(Path)
  Dim BlockPaths, BeginBlockIndex, EndBlockIndex, BlockPath
  
  BlockPaths = Filter(TreePaths, "/" & Path & "/*")
  If ((UBound(BlockPaths) - LBound(BlockPaths)) <> 1) Then
    Set Block = Nothing
  Else
    BlockPath = BlockPaths(0)
    BeginBlockIndex = CLng(Mid(BlockPath, InStrRev(BlockPath, ":") + 1))
    BlockPath = BlockPaths(1)
    EndBlockIndex = CLng(Mid(BlockPath, InStrRev(BlockPath, ":") + 1))
    
    Dim TemplateBlock
    Set TemplateBlock = New clsTemplateBlock
    Set TemplateBlock.Template = Me
    With TemplateBlock
      .Path = Path
      .BeginOfBlock = BeginBlockIndex
      .EndOfBlock = EndBlockIndex
    End With
    Set Block = TemplateBlock
  End If
End Function

End Class

Class clsTemplateBlock
  Public Template
  Public Path
  Public BeginOfBlock
  Public EndOfBlock

  Sub SetVar(VariableName, Value)
    Template.SetVar Path & "/" & VariableName, Value
  End Sub

  Sub Clear()
    Template.SetVar Path, ""
  End Sub

  Property Let HTML(NewValue)
    Template.SetBlockHTML Me, NewValue
  End Property

  Property Get HTML()
    HTML = Template.GetBlockHTML(Me)
  End Property
  
  Function GetVar(VariableName, Value)
    GetVar = Template.GetVar(Path & "/" & VariableName)
  End Function

  
  Sub Show()
    Template.ParseBlockByIndex BeginOfBlock, EndOfBlock, False
  End Sub

  Sub Parse(Accumulate)
    Template.ParseBlockByIndex BeginOfBlock, EndOfBlock, Accumulate
  End Sub
  
  Sub PParse(Accumulate)
    Template.PParse Path, Accumulate
  End Sub

  Sub ParseTo(Accumulate, TargetBlock)
    Template.ParseTo Path, Accumulate, TargetBlock.Path
  End Sub

  Function BlockExists(BlockName, BlockType)
    BlockExists = Template.BlockExists(Path & "/" & BlockName, BlockType)
  End Function

  Function Block(NewPath)
    Set Block = Template.Block(Path & "/" & NewPath)
  End Function

  Property Get Variable(VarName)
    Variable = Template.GetVar(Path & "/@" & VarName)
  End Property

  Property Let Variable(VarName, NewValue)
    Template.SetVar Path & "/@" & VarName, NewValue
  End Property

  Property Let BlockVariable(VariableName, Value)
     Template.SetBlockVariable BeginOfBlock, EndOfBlock, "/@" & VariableName, Value
  End Property


  Property Let Visible(Value)
    If Value Then
      Template.Parse Path, False
    Else
      Template.SetVar Path, ""
    End If
  End Property

End Class

'End Template class


%>
