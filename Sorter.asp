<%

'Sorter Class @0-717EDB8A

Function CCCreateSorter(SorterName, Parent, FileName)
  Dim Sorter
  Set Sorter = New clsSorter
  With Sorter
    .ComponentName = SorterName
    .FileName = FileName
    Set .Parent = Parent
  End With
  Set CCCreateSorter = Sorter
End Function

Class clsSorter

  Public ComponentName, CCSEvents

  Dim OrderDirection
  Dim TargetName
  Dim FileName
  Dim Visible
  Dim Attributes
  Dim PageAttributes

  Private mParent
  Private CCSEventResult

  Private Sub Class_Initialize()
    Visible = True
    Set mParent = Nothing
    Set CCSEvents = CreateObject("Scripting.Dictionary")
    Set Attributes = New clsAttributes
    Set PageAttributes = New clsAttributes
    PageAttributes("pathToRoot") = PathToRoot
  End Sub

  Private Sub Class_Terminate()
    Set mParent = Nothing
    Set Attributes = Nothing
    Set PageAttributes = Nothing
  End Sub

  Property Set Parent(newParent)
    Set mParent = newParent
  End Property

  Private Function GetLink(Par)
    If CCSUseAmps then 
      GetLink = Replace(Par, "&", "&amp;")
    Else 
      GetLink = Par
    End If
  End Function

  Sub Show(Template)
    Dim IsOn, IsAsc
    Dim QueryString, SorterBlock
    Dim AscOnExist, AscOffExist, DescOnExist, DescOffExist
    Dim AscOn, AscOff, DescOn, DescOff

    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)

    If Not Visible Then _
      Exit Sub

    TargetName = mParent.ComponentName
    IsOn = (mParent.ActiveSorter = ComponentName)
    IsAsc = (isEmpty(mParent.SortingDirection) Or mParent.SortingDirection = "ASC")
    Set SorterBlock = Template.Block("Sorter " & ComponentName)
    Dim Prefix : Prefix = ComponentName & ":" 
    Attributes.Show SorterBlock, Prefix
    PageAttributes.Show SorterBlock, "page:"
    AscOnExist = SorterBlock.BlockExists("Asc_On", "block")
    If AscOnExist Then 
      Set AscOn = SorterBlock.Block("Asc_On")
      AscOn.Variable("CCS_PathToMasterPage") = PathToCurrentMasterPage
      Attributes.Show AscOn, Prefix
    End If
    AscOffExist = SorterBlock.BlockExists("Asc_Off", "block")
    If AscOffExist Then 
      Set AscOff = SorterBlock.Block("Asc_Off")
      Attributes.Show AscOff, Prefix
    End If

    DescOnExist = SorterBlock.BlockExists("Desc_On", "block")
    If DescOnExist Then 
      Set DescOn = SorterBlock.Block("Desc_On")
      DescOn.Variable("CCS_PathToMasterPage") = PathToCurrentMasterPage
      Attributes.Show DescOn, Prefix
    End If

    DescOffExist = SorterBlock.BlockExists("Desc_Off", "block")
    If DescOffExist Then 
      Set DescOff = SorterBlock.Block("Desc_Off")
      Attributes.Show DescOff, Prefix
    End If

    QueryString = CCGetQueryString("QueryString", Array(TargetName & "Page", "ccsForm"))
    QueryString = CCAddParam(QueryString, TargetName & "Order", ComponentName)

    If IsOn then
      If IsAsc then 
        OrderDirection = "DESC"
        If AscOnExist Then AscOn.Visible = True
        If AscOffExist Then AscOff.Visible = False
        If DescOnExist Then DescOn.Visible = False
        If DescOffExist Then
          DescOff.Variable("CCS_PathToMasterPage") = PathToCurrentMasterPage
          DescOff.Variable("Desc_URL") = FileName & "?" & GetLink(CCAddParam(QueryString, TargetName & "Dir", OrderDirection))
          DescOff.Visible = True
        End If
      Else 
        OrderDirection = "ASC"
        If AscOnExist Then AscOn.Visible = False
        If AscOffExist Then 
          AscOff.Variable("CCS_PathToMasterPage") = PathToCurrentMasterPage
          AscOff.Variable("Asc_URL") = FileName & "?" & GetLink(CCAddParam(QueryString, TargetName & "Dir", OrderDirection))
          AscOff.Visible = True
        End If
        If DescOnExist Then DescOn.Visible = True
        If DescOffExist Then DescOff.Visible = False
      End if
    Else
      OrderDirection = "ASC"
      If AscOnExist Then AscOn.Visible = False
      If AscOffExist Then 
      	AscOff.Variable("CCS_PathToMasterPage") = PathToCurrentMasterPage
        AscOff.Variable("Asc_URL") = FileName & "?" & GetLink(CCAddParam(QueryString, TargetName & "Dir", "ASC"))
        AscOff.Visible = True
      End If
      If DescOnExist Then DescOn.Visible = False
      If DescOffExist Then 
      	DescOff.Variable("CCS_PathToMasterPage") = PathToCurrentMasterPage
        DescOff.Variable("Desc_URL") = FileName & "?" & GetLink(CCAddParam(QueryString, TargetName & "Dir", "DESC"))
        DescOff.Visible = True
      End If
    End If

    QueryString = CCAddParam(QueryString, TargetName & "Dir", OrderDirection)
    SorterBlock.Variable("CCS_PathToMasterPage") = PathToCurrentMasterPage
    SorterBlock.Variable("Sort_URL") = FileName & "?" & GetLink(QueryString)
    SorterBlock.Visible = True
    Attributes.Show  SorterBlock, ComponentName & ":"

  End Sub

End Class
'End Sorter Class


%>
