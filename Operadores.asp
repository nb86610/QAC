<%@ CodePage=65001 %>
<%
'Include Common Files @1-70509225
%>
<!-- #INCLUDE VIRTUAL="/Qac/Common.asp"-->
<!-- #INCLUDE VIRTUAL="/Qac/Cache.asp" -->
<!-- #INCLUDE VIRTUAL="/Qac/Template.asp" -->
<!-- #INCLUDE VIRTUAL="/Qac/Sorter.asp" -->
<!-- #INCLUDE VIRTUAL="/Qac/Navigator.asp" -->
<%
'End Include Common Files

'Initialize Page @1-F2687EF9
' Variables
Dim PathToRoot, PathToRootOpt, ScriptPath, TemplateFilePath
Dim FileName
Dim TemplateSource
Dim GlobalHEADContent
Dim Scripts
Dim Redirect
Dim Tpl, HTMLTemplate
Dim TemplateFileName
Dim ComponentName
Dim PathToCurrentPage
Dim Attributes
Dim PathToCurrentMasterPage

' Events
Dim CCSEvents
Dim CCSEventResult

' Connections
Dim DBConnection1

' Page controls
Dim Operadores1
Dim ChildControls

Response.ContentType = CCSContentType
Redirect = ""
TemplateFileName = "Operadores.html"
Set CCSEvents = CreateObject("Scripting.Dictionary")
PathToCurrentPage = "./"
GlobalHEADContent = ""
Scripts = "|js/jquery/jquery.js|js/jquery/event-manager.js|js/jquery/selectors.js|"
FileName = "Operadores.asp"
PathToRoot = "./"
PathToRootOpt = ""
ScriptPath = Left(Request.ServerVariables("PATH_TRANSLATED"), Len(Request.ServerVariables("PATH_TRANSLATED")) - Len(FileName))
TemplateFilePath = ScriptPath
'End Initialize Page

'Initialize Objects @1-B9759793
CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeInitialize", Nothing)

Set DBConnection1 = New clsDBConnection1
DBConnection1.Open
Set Attributes = New clsAttributes
Attributes("pathToRoot") = PathToRoot
Attributes("pathToCurrentPage") = PathToRoot

' Controls
Set Operadores1 = New clsEditableGridOperadores1
Operadores1.Initialize DBConnection1
Dim SList, Script, ScriptIncludes
SList = Split(Scripts, "|")
For Each Script In SList
    If Script <> "" Then ScriptIncludes = ScriptIncludes & "<script src=""" & PathToRootOpt & Script & """ type=""text/javascript""></script>" & vbCrLf 
Next
Attributes("scriptIncludes") = ScriptIncludes

CCSEventResult = CCRaiseEvent(CCSEvents, "AfterInitialize", Nothing)
'End Initialize Objects

'Execute Components @1-50738595
Operadores1.ProcessOperations
'End Execute Components

'Go to destination page @1-C772AD3B
If Not ( Redirect = "" ) Then
    UnloadPage
    Response.Redirect Redirect
End If
'End Go to destination page

'Initialize HTML Template @1-FB5C1132
CCSEventResult = CCRaiseEvent(CCSEvents, "OnInitializeView", Nothing)
Set HTMLTemplate = New clsTemplate
HTMLTemplate.HEADContent = GlobalHEADContent
HTMLTemplate.Encoding = "utf-8"
Set HTMLTemplate.Cache = TemplatesRepository
If IsEmpty(TemplateSource) Then
    HTMLTemplate.LoadTemplate TemplateFilePath & TemplateFileName
Else
    HTMLTemplate.LoadTemplateFromStr TemplateSource
End If
HTMLTemplate.SetVar "@CCS_PathToRoot", PathToRoot
HTMLTemplate.SetVar "@CCS_PathToCurrentPage", PathToRoot
HTMLTemplate.SetVar "@CCS_PathToMasterPage", PathToCurrentMasterPage
Set Tpl = HTMLTemplate.Block("main")
CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Nothing)
'End Initialize HTML Template

'Show Page @1-BEECF3B7
Dim MainHTML
Attributes.Show HTMLTemplate, "page:"
Set ChildControls = CCCreateCollection(Tpl, Null, ccsParseOverwrite, _
    Array(Operadores1))
ChildControls.Show
HTMLTemplate.Parse "main", False
If IsEmpty(MainHTML) Then MainHTML = HTMLTemplate.GetHTML("main")
CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeOutput", Nothing)
If CCSEventResult Then Response.Write CleanHTML(MainHTML)
'End Show Page

'Unload Page @1-CB210C62
UnloadPage
Set Tpl = Nothing
Set HTMLTemplate = Nothing
'End Unload Page

'UnloadPage Sub @1-C5A17111
Sub UnloadPage()
    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeUnload", Nothing)
    If DBConnection1.State = adStateOpen Then _
        DBConnection1.Close
    Set DBConnection1 = Nothing
    Set CCSEvents = Nothing
    Set Attributes = Nothing
    Set Operadores1 = Nothing
End Sub
'End UnloadPage Sub

Class clsEditableGridOperadores1 'Operadores1 Class @2-42440C33

'Operadores1 Variables @2-1938655B

    ' Private variables
    Private VarPageSize
    ' Public variables
    Public ComponentName
    Public HTMLFormAction
    Public HTMLFormMethod
    Public PressedButton
    Public Errors
    Public IsFormSubmitted
    Public EditMode
    Public Visible
    Public Recordset
    Public TemplateBlock
    Public PageNumber
    Public IsDSEmpty
    Public RowNumber
    Public CachedColumns
    Public CachedColumnsNames
    Public CachedColumnsNumber
    Public SubmittedRows
    Public EmptyRows
    Public ErrorMessages
    Public Attributes
    Public PageAttributes

    Public CCSEvents
    Private CCSEventResult

    Public ActiveSorter, SortingDirection
    Public InsertAllowed
    Public UpdateAllowed
    Public DeleteAllowed
    Public ReadAllowed
    Public DataSource
    Public Command
    Public ValidatingControls
    Public Controls
    Public NoRecordsControls
    Private MaxCachedValues
    Private CachedValuesNumber
    Private NewEmptyRows
    Private ErrorControls

    ' Class variables
    Dim Sorter_MN
    Dim Sorter_Funcion
    Dim Sorter_Privilegios
    Dim Sorter_Planta
    Dim Sorter_Email
    Dim Sorter_Activo
    Dim MN
    Dim Funcion
    Dim Privilegios
    Dim Planta
    Dim Email
    Dim Activo
    Dim CheckBox_Delete_Panel
    Dim CheckBox_Delete
    Dim Navigator
    Dim Button_Submit
    Public Row
'End Operadores1 Variables

'Operadores1 Class_Initialize Event @2-ED0D7A53
    Private Sub Class_Initialize()

        Visible = True
        Set Errors = New clsErrors
        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set ErrorControls = CreateObject("Scripting.Dictionary")
        Set Attributes = New clsAttributes
        Set PageAttributes = New clsAttributes
        PageAttributes("pathToRoot") = PathToRoot
        Set DataSource = New clsOperadores1DataSource
        Set Command = New clsCommand
        InsertAllowed = True
        UpdateAllowed = True
        DeleteAllowed = True
        ReadAllowed = True
        Dim Method
        Dim OperationMode
        ComponentName = "Operadores1"

        ActiveSorter = CCGetParam("Operadores1Order", Empty)
        SortingDirection = CCGetParam("Operadores1Dir", Empty)
        If Not(SortingDirection = "ASC" Or SortingDirection = "DESC") Then _
            SortingDirection = Empty

        PageSize = CCGetParam(ComponentName & "PageSize", Empty)
        If IsNumeric(PageSize) And Len(PageSize) > 0 Then
            If PageSize <= 0 Then Errors.AddError(CCSLocales.GetText("CCS_GridPageSizeError", Empty))
            If PageSize > 100 Then PageSize = 100
        End If
        If Not IsNumeric(PageSize) Or IsEmpty(PageSize) Then _
            PageSize = 20 _
        Else _
            PageSize = CInt(PageSize)
        PageNumber = CInt(CCGetParam(ComponentName & "Page", 1))

        If CCGetFromGet("ccsForm", Empty) = ComponentName Then
            IsFormSubmitted = True
            EditMode = True
        Else
            IsFormSubmitted = False
            EditMode = False
        End If
        Method = IIf(IsFormSubmitted, ccsPost, ccsGet)
        Set Sorter_MN = CCCreateSorter("Sorter_MN", Me, FileName)
        Set Sorter_Funcion = CCCreateSorter("Sorter_Funcion", Me, FileName)
        Set Sorter_Privilegios = CCCreateSorter("Sorter_Privilegios", Me, FileName)
        Set Sorter_Planta = CCCreateSorter("Sorter_Planta", Me, FileName)
        Set Sorter_Email = CCCreateSorter("Sorter_Email", Me, FileName)
        Set Sorter_Activo = CCCreateSorter("Sorter_Activo", Me, FileName)
        Set MN = CCCreateControl(ccsLabel, "MN", Empty, ccsText, Empty, CCGetRequestParam("MN", Method))
        Set Funcion = CCCreateControl(ccsTextBox, "Funcion", "Funcion", ccsText, Empty, CCGetRequestParam("Funcion", Method))
        Set Privilegios = CCCreateControl(ccsTextBox, "Privilegios", "Privilegios", ccsInteger, Empty, CCGetRequestParam("Privilegios", Method))
        Set Planta = CCCreateControl(ccsTextBox, "Planta", "Planta", ccsText, Empty, CCGetRequestParam("Planta", Method))
        Set Email = CCCreateControl(ccsTextBox, "Email", "Email", ccsText, Empty, CCGetRequestParam("Email", Method))
        Set Activo = CCCreateControl(ccsCheckBox, "Activo", Empty, ccsBoolean, DefaultBooleanFormat, CCGetRequestParam("Activo", Method))
        Set CheckBox_Delete_Panel = CCCreatePanel("CheckBox_Delete_Panel")
        Set CheckBox_Delete = CCCreateControl(ccsCheckBox, "CheckBox_Delete", Empty, ccsBoolean, DefaultBooleanFormat, CCGetRequestParam("CheckBox_Delete", Method))
        CheckBox_Delete.CheckedValue = true
        CheckBox_Delete.UncheckedValue = false
        Set Navigator = CCCreateNavigator(ComponentName, "Navigator", FileName, 10, tpSimple)
        Navigator.PageSizes = Array("1", "5", "10", "25", "50")
        Set Button_Submit = CCCreateButton("Button_Submit", Method)
        Set ValidatingControls = New clsControls
        ValidatingControls.addControls Array(Funcion, Privilegios, Planta, Email, Activo)
        If Not IsFormSubmitted Then
            If IsEmpty(CheckBox_Delete.State) Then _
                CheckBox_Delete.State = ccsUnchecked
        End If

        SubmittedRows = 0
        NewEmptyRows = 0
        EmptyRows = 1
        CheckBox_Delete_Panel.AddComponent(CheckBox_Delete)

        InitCachedColumns()

        IsDSEmpty = True
    End Sub
'End Operadores1 Class_Initialize Event

'Operadores1 InitCachedColumns Method @2-0D42A35F
    Sub InitCachedColumns()
        Dim RetrievedNumber, i
        CachedColumnsNumber = 1
        ReDim CachedColumnsNames(CachedColumnsNumber)
        CachedColumnsNames(0) = "MN"

        RetrievedNumber = 0
        CachedColumns = GetCachedColumns()

        If IsArray(CachedColumns) Then
            RetrievedNumber = UBound(CachedColumns)
            If RetrievedNumber > 0 Then
                MaxCachedValues = CInt(RetrievedNumber / CachedColumnsNumber)
                If (RetrievedNumber Mod CachedColumnsNumber) > 0 Then
                    MaxCachedValues = MaxCachedValues + 1
                End If
                CachedValuesNumber = MaxCachedValues
            End If
        End If

        If RetrievedNumber = 0 Then
            MaxCachedValues = 50
            ReDim CachedColumns(MaxCachedValues * CachedColumnsNumber)
            CachedValuesNumber = 0
        End If 

        If SubmittedRows > 0 Or NewEmptyRows > 0 Then
            EmptyRows = NewEmptyRows
        End If

        DataSource.CachedColumns = CachedColumns
        DataSource.CachedColumnsNumber = CachedColumnsNumber
        ReDim ErrorMessages(SubmittedRows + EmptyRows)
    End Sub
'End Operadores1 InitCachedColumns Method

'Operadores1 Initialize Method @2-6EE7F8D7
    Sub Initialize(objConnection)

        If Not Visible Then Exit Sub

        Set DataSource.Connection = objConnection
        With DataSource

            Set .Connection = objConnection
            .PageSize = PageSize
            .SetOrder ActiveSorter, SortingDirection
            .AbsolutePage = PageNumber
        End With
    End Sub
'End Operadores1 Initialize Method

'Operadores1 Class_Terminate Event @2-3ADA93ED
    Private Sub Class_Terminate()
        Set CCSEvents = Nothing
        Set Errors = Nothing
        Set Attributes = Nothing
        Set PageAttributes = Nothing
    End Sub
'End Operadores1 Class_Terminate Event

'Operadores1 Validate Method @2-F7DC8E56
    Function Validate()
        Dim Validation
        Dim i, InsertedRows, Method, IsDeleted, IsEmptyRow, IsNewRow,EGErrors
        Method = IIf(IsFormSubmitted, ccsPost, ccsGet)
        Validation = True

        If SubmittedRows > 0 Then
            Set EGErrors = New clsErrors
            EGErrors.AddErrors(Errors)
            Errors.Clear
            For i = 1 To SubmittedRows
                IsDeleted = (Len(CCGetRequestParam("CheckBox_Delete_" & CStr(i), Method)) > 0)
                IsEmptyRow = (Len(CCGetRequestParam("Funcion_" & CStr(i), Method)) = 0)
                IsEmptyRow = IsEmptyRow And (Len(CCGetRequestParam("Privilegios_" & CStr(i), Method)) = 0)
                IsEmptyRow = IsEmptyRow And (Len(CCGetRequestParam("Planta_" & CStr(i), Method)) = 0)
                IsEmptyRow = IsEmptyRow And (Len(CCGetRequestParam("Email_" & CStr(i), Method)) = 0)
                IsEmptyRow = IsEmptyRow And (Len(CCGetRequestParam("Activo_" & CStr(i), Method)) = 0)

                If (Not IsDeleted) And (Not IsEmptyRow Or (i < SubmittedRows - EmptyRows + 1)) Then
                    Funcion.Errors.Clear
                    Funcion.Text = CCGetRequestParam("Funcion_" & CStr(i), Method)
                    Privilegios.Errors.Clear
                    Privilegios.Text = CCGetRequestParam("Privilegios_" & CStr(i), Method)
                    Planta.Errors.Clear
                    Planta.Text = CCGetRequestParam("Planta_" & CStr(i), Method)
                    Email.Errors.Clear
                    Email.Text = CCGetRequestParam("Email_" & CStr(i), Method)
                    Activo.Errors.Clear
                    Activo.Text = CCGetRequestParam("Activo_" & CStr(i), Method)
                    ValidatingControls.Validate
                    CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidateRow", Me)
                    If Not ValidatingControls.IsValid() or Errors.Count >0 Then
                        Errors.AddErrors Funcion.Errors
                        Errors.AddErrors Privilegios.Errors
                        Errors.AddErrors Planta.Errors
                        Errors.AddErrors Email.Errors
                        Errors.AddErrors Activo.Errors
                        Errors.AddErrors CheckBox_Delete.Errors
                        ErrorMessages(i) = Errors.ToString()
                        Validation = False
                        Errors.Clear
                    End If
                End If
            Next
            Errors.AddErrors(EGErrors)
            Set EGErrors = Nothing
        End If

        CCSEventResult = CCRaiseEvent(CCSEvents, "OnValidate", Me)
        Validate = Validation And (Errors.Count = 0)
    End Function
'End Operadores1 Validate Method

'Operadores1 ProcessOperations Method @2-FEB0BD33
    Sub ProcessOperations()
        Dim TmpWhere: TmpWhere = Datasource.Where

        If Not ( Visible And IsFormSubmitted ) Then Exit Sub

        If IsFormSubmitted Then
            PressedButton = IIf(EditMode, "Button_Submit", "Button_Submit")
            If Button_Submit.Pressed Then
                PressedButton = "Button_Submit"
            End If
        End If
        Redirect = FileName & "?" & CCGetQueryString("QueryString",Array("ccsForm", "Button_Submit.x", "Button_Submit"))

        If Validate() Then
            If PressedButton = "Button_Submit" Then
                CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSubmit", Me)
                If Not Button_Submit.OnClick() Or (Not InsertRows() And InsertAllowed) Or (Not UpdateRows() And UpdateAllowed) Or (Not DeleteRows() And DeleteAllowed) Then
                    Redirect = ""
                End If
                CCSEventResult = CCRaiseEvent(CCSEvents, "AfterSubmit", Me)
            End If
        Else
            Redirect = ""
        End If

        Datasource.Where = TmpWhere
    End Sub
'End Operadores1 ProcessOperations Method

'Operadores1 InsertRows Method @2-9A50BC90
    Function InsertRows()
        If Not InsertAllowed Then InsertRows = False : Exit Function

        Dim i, InsertedRows, Method, IsDeleted, IsEmptyRow, HasErrors

        Method = IIf(IsFormSubmitted, ccsPost, ccsGet)

        If SubmittedRows > 0 Then
            i = SubmittedRows - EmptyRows
            For i = (SubmittedRows - EmptyRows + 1) To SubmittedRows
                IsDeleted = (Len(CCGetRequestParam("CheckBox_Delete_" & CStr(i), Method)) > 0)
                IsEmptyRow = True
                IsEmptyRow = IsEmptyRow And (Len(CCGetRequestParam("Funcion_" & CStr(i), Method)) = 0)
                IsEmptyRow = IsEmptyRow And (Len(CCGetRequestParam("Privilegios_" & CStr(i), Method)) = 0)
                IsEmptyRow = IsEmptyRow And (Len(CCGetRequestParam("Planta_" & CStr(i), Method)) = 0)
                IsEmptyRow = IsEmptyRow And (Len(CCGetRequestParam("Email_" & CStr(i), Method)) = 0)
                IsEmptyRow = IsEmptyRow And (Len(CCGetRequestParam("Activo_" & CStr(i), Method)) = 0)

                If (Not IsDeleted) And (Not IsEmptyRow) Then
                    Funcion.Text = CCGetRequestParam("Funcion_" & CStr(i), Method)
                    DataSource.Funcion.Value = Funcion.Value
                    Privilegios.Text = CCGetRequestParam("Privilegios_" & CStr(i), Method)
                    DataSource.Privilegios.Value = Privilegios.Value
                    Planta.Text = CCGetRequestParam("Planta_" & CStr(i), Method)
                    DataSource.Planta.Value = Planta.Value
                    Email.Text = CCGetRequestParam("Email_" & CStr(i), Method)
                    DataSource.Email.Value = Email.Value
                    DataSource.Activo.Value = IIf(IsEmpty(CCGetRequestParam("Activo_" & CStr(i), Method)),Activo.UncheckedValue,Activo.CheckedValue)
                    DataSource.CurrentRow = i
                    DataSource.Insert(Command)


                    If DataSource.Errors.Count > 0 Then
                        HasErrors = True
                        ErrorMessages(i) = DataSource.Errors.ToString()
                        DataSource.Errors.Clear
                    End If
                End If
            Next
        End If

        InsertRows = Not(HasErrors)
    End Function
'End Operadores1 InsertRows Method

'Operadores1 UpdateRows Method @2-7B0282D9
    Function UpdateRows()
        If Not UpdateAllowed Then UpdateRows = False : Exit Function

        Dim i, InsertedRows, Method, IsDeleted, HasErrors
        Method = IIf(IsFormSubmitted, ccsPost, ccsGet)
        If SubmittedRows > 0 Then
            For i = 1 To SubmittedRows - EmptyRows
                IsDeleted = (Len(CCGetRequestParam("CheckBox_Delete_" & CStr(i), Method)) > 0)

                If Not IsDeleted Then
                    Funcion.Text = CCGetRequestParam("Funcion_" & CStr(i), Method)
                    DataSource.Funcion.Value = Funcion.Value
                    Privilegios.Text = CCGetRequestParam("Privilegios_" & CStr(i), Method)
                    DataSource.Privilegios.Value = Privilegios.Value
                    Planta.Text = CCGetRequestParam("Planta_" & CStr(i), Method)
                    DataSource.Planta.Value = Planta.Value
                    Email.Text = CCGetRequestParam("Email_" & CStr(i), Method)
                    DataSource.Email.Value = Email.Value
                    DataSource.Activo.Value = IIf(IsEmpty(CCGetRequestParam("Activo_" & CStr(i), Method)),Activo.UncheckedValue,Activo.CheckedValue)
                    DataSource.CurrentRow = i
                    DataSource.Update(Command)


                    If DataSource.Errors.Count > 0 Then
                        HasErrors = True
                        ErrorMessages(i) = DataSource.Errors.ToString()
                        DataSource.Errors.Clear
                    End If
                End If
            Next
        End If
        UpdateRows = Not(HasErrors)
    End Function
'End Operadores1 UpdateRows Method

'Operadores1 DeleteRows Method @2-9938C450
    Function DeleteRows()
        Dim i, Method, HasErrors

        Method = IIf(IsFormSubmitted, ccsPost, ccsGet)
        If Not DeleteAllowed Then DeleteRows = False : Exit Function


        If SubmittedRows > 0 Then
            For i = 1 To SubmittedRows - EmptyRows
                If Len(CCGetRequestParam("CheckBox_Delete_" & CStr(i), Method)) > 0 Then
                    DataSource.CurrentRow = i
                    DataSource.Delete(Command)


                    If DataSource.Errors.Count > 0 Then
                        HasErrors = True
                        ErrorMessages(i) = DataSource.Errors.ToString()
                        DataSource.Errors.Clear
                    End If
                End If
            Next
        End If

        DeleteRows = Not(HasErrors)
    End Function
'End Operadores1 DeleteRows Method

'GetFormScript Method @2-EDC89B42
    Function GetFormScript(TotalRows)
        Dim script,i: script = ""
        GetFormScript = script
    End Function
'End GetFormScript Method

'GetFormState Method @2-8BEF9A95
    Function GetFormState
        Dim FormState, i, LastValueIndex, NewRows

        FormState = ""
        LastValueIndex = CachedValuesNumber * CachedColumnsNumber - 1

        If EditMode And LastValueIndex >= 0 Then

            For i = 1 To CachedColumnsNumber
                FormState = FormState & CachedColumnsNames(i-1)
                If i < CachedColumnsNumber Or LastValueIndex >= 0 Then FormState = FormState & ";"
            Next

            For i = 0 To LastValueIndex
                If IsNull(CachedColumns(i)) Then  CachedColumns(i) = ""
                FormState = FormState & CCToHTML(CCEscapeLOV(CachedColumns(i)))
                If i < LastValueIndex Then FormState = FormState & ";"
            Next
        End If

        NewRows = IIf(InsertAllowed, EmptyRows, 0)
        GetFormState = CStr(SubmittedRows - NewRows) & ";" & CStr(NewRows)
        If Len(FormState) > 0 Then GetFormState = GetFormState & ";" & FormState

    End Function
'End GetFormState Method

'GetCachedColumns Method @2-AA749F06
    Function GetCachedColumns
        Dim FormState, i, TotalValuesNumber, TempColumns, NewCachedColumns, TempValuesNumber
        Dim NewSubmittedRows : NewSubmittedRows = 0

        NewCachedColumns = Empty
        FormState = CCGetRequestParam("FormState", ccsPost)

        If CCGetFromGet("ccsForm", Empty) = ComponentName Then
            If Not IsNull(FormState) Then
                If InStr(FormState,"\;") > 0 Then _
                    FormState = Replace(FormState, "\;", "<!--semicolon-->")
                If InStr(FormState,";") > 0 Then 
                    TempColumns = Split(FormState,";")
                    If IsArray(TempColumns) Then 
                        TempValuesNumber = UBound(TempColumns) - 1
                        If TempValuesNumber >= 0 Then
                            NewSubmittedRows = TempColumns(0)
                            NewEmptyRows     = TempColumns(1)
                        End If
                        SubmittedRows = CLng(NewSubmittedRows) + CLng(NewEmptyRows)

                        If TempValuesNumber > 1 And TempValuesNumber >= CachedColumnsNumber Then
                            ReDim NewCachedColumns(TempValuesNumber - CachedColumnsNumber + 1)
                            For i = 0 To TempValuesNumber - CachedColumnsNumber - 1
                                NewCachedColumns(i) = Replace(CCUnEscapeLOV(TempColumns(i + CachedColumnsNumber + 2)),"<!--semicolon-->",";")
                            Next
                        End If
                    End If
                Else
                    SubmittedRows = FormState
                End If
            End If
        End If

        GetCachedColumns = NewCachedColumns
    End Function
'End GetCachedColumns Method

'Operadores1 Show Method @2-0CE6270B
    Sub Show(Tpl)
        Dim StaticControls, RowControls, NoRecordsBlock

        If Not Visible Then Exit Sub

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeSelect", Me)
        Set Recordset = DataSource.Open(Command)
        If Recordset.State = adStateOpen Then 
            EditMode = Not Recordset.EOF 
        Else
            EditMode = False
        End If
        IsDSEmpty = Not EditMode

        HTMLFormAction = FileName & "?" & CCAddParam(Request.ServerVariables("QUERY_STRING"), "ccsForm", "Operadores1")
        Set TemplateBlock = Tpl.Block("EditableGrid " & ComponentName)
        TemplateBlock.Template.SetVar "@HTMLFormName", ComponentName
        TemplateBlock.Template.SetVar "@Action", IIF(CCSUseAmps, Replace(HTMLFormAction, "&", CCSAmps), HTMLFormAction)
        TemplateBlock.Template.SetVar "@HTMLFormProperties", "action=""" & HTMLFormAction & """ method=""post"" " & "name=""" & ComponentName & """"
        TemplateBlock.Template.SetVar "@HTMLFormEnctype", "application/x-www-form-urlencoded"
        TemplateBlock.Variable("CCS_PathToMasterPage") = PathToCurrentMasterPage

        Set NoRecordsBlock = TemplateBlock.Block("NoRecords")
        If DataSource.Errors.Count > 0 Then
            Errors.AddErrors(DataSource.Errors)
            DataSource.Errors.Clear
            With TemplateBlock.Block("Error")
                .Variable("Error") = Errors.ToString
                .Parse False
            End With
        End If

        Set NoRecordsControls = CCCreateCollection(NoRecordsBlock, Null, ccsParseOverwrite, _
            Array())
        Set StaticControls = CCCreateCollection(TemplateBlock, Null, ccsParseOverwrite, _
            Array(Sorter_MN, Sorter_Funcion, Sorter_Privilegios, Sorter_Planta, Sorter_Email, Sorter_Activo, Navigator, Button_Submit))
        Set RowControls = CCCreateCollection(TemplateBlock.Block("Row"), Null, ccsParseAccumulate, _
            Array(MN, Funcion, Privilegios, Planta, Email, Activo, CheckBox_Delete_Panel))

        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShow", Me)
        If Not Visible Then Exit Sub

        RowControls.PreserveControlsVisible
        If Not DeleteAllowed Then CheckBox_Delete.Visible = False

        Dim i, j
        i = 1 : j = 0
        If EditMode And ReadAllowed Then
            If Recordset.Errors.Count > 0 Then
                With TemplateBlock.Block("Error")
                    .Variable("Error") = Recordset.Errors.ToString
                    .Parse False
                End With
            ElseIf Not Recordset.EOF Then
                While Not Recordset.EoF And (i-1) < pageSize
                    RowNumber = i
                    Attributes("rowNumber") = i
                    MN.Value = Recordset.Fields("MN")

                    If Not IsFormSubmitted Then
                        Funcion.Value = Recordset.Fields("Funcion")
                    Else
                        Funcion.Text = CCGetRequestParam("Funcion_" & CStr(i), ccsPost)
                    End If
                    If Not IsFormSubmitted Then
                        Privilegios.Value = Recordset.Fields("Privilegios")
                    Else
                        Privilegios.Text = CCGetRequestParam("Privilegios_" & CStr(i), ccsPost)
                    End If
                    If Not IsFormSubmitted Then
                        Planta.Value = Recordset.Fields("Planta")
                    Else
                        Planta.Text = CCGetRequestParam("Planta_" & CStr(i), ccsPost)
                    End If
                    If Not IsFormSubmitted Then
                        Email.Value = Recordset.Fields("Email")
                    Else
                        Email.Text = CCGetRequestParam("Email_" & CStr(i), ccsPost)
                    End If
                    If Not IsFormSubmitted Then
                        Activo.Value = Recordset.Fields("Activo")
                    Else
                        Activo.Value = CCGetRequestParam("Activo_" & CStr(i), ccsPost)
                    End If
                    If IsFormSubmitted Then 
                        CheckBox_Delete.Value = CCGetRequestParam("CheckBox_Delete_" & CStr(i), ccsPost)
                    End If
                    MN.ExternalName = "MN_" & CStr(i)
                    Funcion.ExternalName = "Funcion_" & CStr(i)
                    Privilegios.ExternalName = "Privilegios_" & CStr(i)
                    Planta.ExternalName = "Planta_" & CStr(i)
                    Email.ExternalName = "Email_" & CStr(i)
                    Activo.ExternalName = "Activo_" & CStr(i)
                    CheckBox_Delete.ExternalName = "CheckBox_Delete_" & CStr(i)

                    If j >= MaxCachedValues Then
                            MaxCachedValues = MaxCachedValues + 50
                            ReDim Preserve CachedColumns(MaxCachedValues*CachedColumnsNumber)
                    End If
                    CachedColumns(j * CachedColumnsNumber) = Recordset.Recordset.Fields("MN")
                    CachedValuesNumber = i

                    If IsFormSubmitted Then
                        If Len(ErrorMessages(i)) > 0 Then
                            With TemplateBlock.Block("Row").Block("RowError")
                                .Variable("Error") = ErrorMessages(i)
                                .Parse False
                            End With
                        Else
                            TemplateBlock.Block("Row").Block("RowError").Visible = False
                        End If
                    End If

                    CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShowRow", Me)
                    Attributes.Show TemplateBlock.Block("Row"), "Operadores1:"
                    RowControls.Show
                    i = i + 1: j = j + 1

                    Recordset.MoveNext
                    If Not Recordset.EoF And (i-1) < PageSize Then _
                        TemplateBlock.Block("Separator").ParseTo ccsParseAccumulate, TemplateBlock.Block("Row")
                Wend
            End If
            Attributes.Show TemplateBlock, "Operadores1:"
            MN.Value = Recordset.Fields("MN")
        ElseIf Not EditMode And (Not InsertAllowed Or EmptyRows=0)Then
            NoRecordsControls.Show
            NoRecordsBlock.Parse ccsParseOverwrite
        End If

        If Not InsertAllowed And Not UpdateAllowed And Not DeleteAllowed Then
            Button_Submit.Visible = False
        End If

        CheckBox_Delete.Visible = False
        CheckBox_Delete_Panel.Visible = CheckBox_Delete.Visible
        MN.Value = ""
        Funcion.Value = ""
        Privilegios.Value = ""
        Planta.Value = ""
        Email.Value = ""
        Activo.Value = ""

        Dim NewRows
        NewRows = IIf(InsertAllowed, EmptyRows, 0)
        For i = i To i + NewRows - 1
            Attributes("rowNumber") = i
            TemplateBlock.Block("Separator").ParseTo ccsParseAccumulate, TemplateBlock.Block("Row")
            MN.ExternalName = "MN_" & CStr(i)
            Funcion.ExternalName = "Funcion_" & CStr(i)
            Privilegios.ExternalName = "Privilegios_" & CStr(i)
            Planta.ExternalName = "Planta_" & CStr(i)
            Email.ExternalName = "Email_" & CStr(i)
            Activo.ExternalName = "Activo_" & CStr(i)

            If IsFormSubmitted Then 
                CheckBox_Delete.Value = CCGetRequestParam("CheckBox_Delete_" & CStr(i), ccsPost)
            End If

            If IsFormSubmitted Then
                MN.Text = CCGetRequestParam("MN_" & CStr(i), ccsPost)
                Funcion.Text = CCGetRequestParam("Funcion_" & CStr(i), ccsPost)
                Privilegios.Text = CCGetRequestParam("Privilegios_" & CStr(i), ccsPost)
                Planta.Text = CCGetRequestParam("Planta_" & CStr(i), ccsPost)
                Email.Text = CCGetRequestParam("Email_" & CStr(i), ccsPost)
                Activo.Text = CCGetRequestParam("Activo_" & CStr(i), ccsPost)

                If Len(ErrorMessages(i)) > 0 Then
                    With TemplateBlock.Block("Row").Block("RowError")
                        .Variable("Error") = ErrorMessages(i)
                        .Parse False
                    End With
                Else
                    TemplateBlock.Block("Row").Block("RowError").Visible = False
                End If
            End If

            CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeShowRow", Me)
            Attributes.Show TemplateBlock.Block("Row"), "Operadores1:"
            PageAttributes.Show TemplateBlock, "page:"
            RowControls.Show
        Next

        SubmittedRows = i - 1
        TemplateBlock.Template.SetVar "@FormScript", GetFormScript(i - 1)
        TemplateBlock.Template.SetVar "@FormState", GetFormState()

        If IsFormSubmitted Then
            If Errors.Count > 0 Or DataSource.Errors.Count > 0 Then
                Errors.addErrors DataSource.Errors
                With TemplateBlock.Block("Error")
                    .Variable("Error") = Errors.ToString
                    .Parse False
                End With
            End If
        End If

        Navigator.PageSize = PageSize
        Navigator.SetDataSource Recordset
        StaticControls.Show

    End Sub
'End Operadores1 Show Method

'Operadores1 PageSize Property Let @2-54E46DD6
    Public Property Let PageSize(NewValue)
        VarPageSize = NewValue
        DataSource.PageSize = NewValue
    End Property
'End Operadores1 PageSize Property Let

'Operadores1 PageSize Property Get @2-9AA1D1E9
    Public Property Get PageSize()
        PageSize = VarPageSize
    End Property
'End Operadores1 PageSize Property Get

End Class 'End Operadores1 Class @2-A61BA892

Class clsOperadores1DataSource 'Operadores1DataSource Class @2-AB72E5EE

'DataSource Variables @2-14DC2172
    Public Errors, Connection, Parameters, CCSEvents

    Public Recordset
    Public SQL, CountSQL, Order, Where, Orders, StaticOrder
    Public PageSize
    Public PageCount
    Public AbsolutePage
    Public Fields
    Dim WhereParameters
    Public AllParamsSet
    Public CachedColumns
    Public CachedColumnsNumber
    Public CurrentRow
    Public CmdExecution
    Public InsertOmitIfEmpty
    Public UpdateOmitIfEmpty

    Private CurrentOperation
    Private CCSEventResult

    ' Datasource fields
    Public MN
    Public Funcion
    Public Privilegios
    Public Planta
    Public Email
    Public Activo
'End DataSource Variables

'DataSource Class_Initialize Event @2-0C00C133
    Private Sub Class_Initialize()

        Set CCSEvents = CreateObject("Scripting.Dictionary")
        Set Fields = New clsFields
        Set Recordset = New clsDataSource
        Set Recordset.DataSource = Me
        Set Errors = New clsErrors
        Set Connection = Nothing
        AllParamsSet = True
        Set MN = CCCreateField("MN", "MN", ccsText, Empty, Recordset)
        Set Funcion = CCCreateField("Funcion", "Funcion", ccsText, Empty, Recordset)
        Set Privilegios = CCCreateField("Privilegios", "Privilegios", ccsInteger, Empty, Recordset)
        Set Planta = CCCreateField("Planta", "Planta", ccsText, Empty, Recordset)
        Set Email = CCCreateField("Email", "Email", ccsText, Empty, Recordset)
        Set Activo = CCCreateField("Activo", "Activo", ccsBoolean, Array("true", "false", Empty), Recordset)
        Fields.AddFields Array(MN, Funcion, Privilegios, Planta, Email, Activo)
        Set InsertOmitIfEmpty = CreateObject("Scripting.Dictionary")
        InsertOmitIfEmpty.Add "Funcion", True
        InsertOmitIfEmpty.Add "Privilegios", True
        InsertOmitIfEmpty.Add "Planta", True
        InsertOmitIfEmpty.Add "Email", True
        InsertOmitIfEmpty.Add "Activo", False
        Set UpdateOmitIfEmpty = CreateObject("Scripting.Dictionary")
        UpdateOmitIfEmpty.Add "Funcion", True
        UpdateOmitIfEmpty.Add "Privilegios", True
        UpdateOmitIfEmpty.Add "Planta", True
        UpdateOmitIfEmpty.Add "Email", True
        UpdateOmitIfEmpty.Add "Activo", False
        Orders = Array( _ 
            Array("Sorter_MN", "MN", ""), _
            Array("Sorter_Funcion", "Funcion", ""), _
            Array("Sorter_Privilegios", "Privilegios", ""), _
            Array("Sorter_Planta", "Planta", ""), _
            Array("Sorter_Email", "Email", ""), _
            Array("Sorter_Activo", "Activo", ""))

        SQL = "SELECT TOP {SqlParam_endRecord} MN, [Apellido y Nombre] AS [Apellido_y Nombre], Funcion, Privilegios, Planta, [Fecha de Alta] AS [Fecha_de Alta], Email, Activo  " & vbLf & _
        "FROM Operadores {SQL_Where} {SQL_OrderBy}"
        CountSQL = "SELECT COUNT(*) " & vbLf & _
        "FROM Operadores"
        Where = ""
        Order = ""
        StaticOrder = ""
    End Sub
'End DataSource Class_Initialize Event

'SetOrder Method @2-68FC9576
    Sub SetOrder(Column, Direction)
        Order = Recordset.GetOrder(Order, Column, Direction, Orders)
    End Sub
'End SetOrder Method

'BuildTableWhere Method @2-7A10879E
    Public Sub BuildTableWhere()
        If CurrentRow > 0 Then
            Where = "MN=" & Connection.ToSQL(CachedColumns((CurrentRow - 1) * CachedColumnsNumber), ccsText)
        End If
    End Sub
'End BuildTableWhere Method

'Open Method @2-20B00191
    Function Open(Cmd)
        Errors.Clear
        CurrentRow = 0
        If Connection Is Nothing Then
            Set Open = New clsEmptyDataSource
            Exit Function
        End If
        Set Cmd.Connection = Connection
        Cmd.CommandOperation = cmdOpen
        Cmd.PageSize = PageSize
        Cmd.ActivePage = AbsolutePage
        Cmd.CommandType = dsTable
        Set WhereParameters = Nothing
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildSelect", Me)
        Cmd.SQL = SQL
        Cmd.CountSQL = CountSQL
        Cmd.Where = Where
        Cmd.OrderBy = Order
        If(Len(StaticOrder)>0) Then
            If Len(Order)>0 Then Cmd.OrderBy = ", "+Cmd.OrderBy
            Cmd.OrderBy = StaticOrder + Cmd.OrderBy
        End If
        Cmd.Options("TOP") = True
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteSelect", Me)
        If Errors.Count = 0 And CCSEventResult Then _
            Set Recordset = Cmd.Exec(Errors)
        CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteSelect", Me)
        Set Recordset.FieldsCollection = Fields
        Set Open = Recordset
    End Function
'End Open Method

'DataSource Class_Terminate Event @2-41B4B08D
    Private Sub Class_Terminate()
        If Recordset.State = adStateOpen Then _
            Recordset.Close
        Set Recordset = Nothing
        Set Parameters = Nothing
        Set Errors = Nothing
    End Sub
'End DataSource Class_Terminate Event

'Delete Method @2-ADD851B7
    Sub Delete(Cmd)
        CmdExecution = True
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildDelete", Me)
        Set Cmd.Connection = Connection
        Cmd.CommandOperation = cmdExec
        Cmd.CommandType = dsTable
        Cmd.CommandParameters = Empty
        BuildTableWhere
        If Not AllParamsSet Then
            Errors.AddError(CCSLocales.GetText("CCS_CustomOperationError_MissingParameters", Empty))
        End If
        Cmd.SQL = "DELETE FROM Operadores" & IIf(Len(Where) > 0, " WHERE " & Where, "")
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteDelete", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteDelete", Me)
        End If
    End Sub
'End Delete Method

'Update Method @2-7F68E002
    Sub Update(Cmd)
        CmdExecution = True
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildUpdate", Me)
        Set Cmd.Connection = Connection
        Cmd.CommandOperation = cmdExec
        Cmd.CommandType = dsTable
        Cmd.CommandParameters = Empty
        BuildTableWhere
        If Not AllParamsSet Then
            Errors.AddError(CCSLocales.GetText("CCS_CustomOperationError_MissingParameters", Empty))
        End If
        Dim IsDef_Funcion : IsDef_Funcion = CCIsDefined("Funcion_" & CurrentRow, "Form")
        Dim IsDef_Privilegios : IsDef_Privilegios = CCIsDefined("Privilegios_" & CurrentRow, "Form")
        Dim IsDef_Planta : IsDef_Planta = CCIsDefined("Planta_" & CurrentRow, "Form")
        Dim IsDef_Email : IsDef_Email = CCIsDefined("Email_" & CurrentRow, "Form")
        Dim IsDef_Activo : IsDef_Activo = CCIsDefined("Activo_" & CurrentRow, "Form")
        If Not UpdateOmitIfEmpty("Funcion") Or IsDef_Funcion Then Cmd.AddSQLStrings "Funcion=" & Connection.ToSQL(Funcion, Funcion.DataType), Empty
        If Not UpdateOmitIfEmpty("Privilegios") Or IsDef_Privilegios Then Cmd.AddSQLStrings "Privilegios=" & Connection.ToSQL(Privilegios, Privilegios.DataType), Empty
        If Not UpdateOmitIfEmpty("Planta") Or IsDef_Planta Then Cmd.AddSQLStrings "Planta=" & Connection.ToSQL(Planta, Planta.DataType), Empty
        If Not UpdateOmitIfEmpty("Email") Or IsDef_Email Then Cmd.AddSQLStrings "Email=" & Connection.ToSQL(Email, Email.DataType), Empty
        If Not UpdateOmitIfEmpty("Activo") Or IsDef_Activo Then Cmd.AddSQLStrings "Activo=" & Connection.ToSQL(Activo, Activo.DataType), Empty
        CmdExecution = Cmd.PrepareSQL("Update", "Operadores", Where)
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteUpdate", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteUpdate", Me)
        End If
    End Sub
'End Update Method

'Insert Method @2-90479E9C
    Sub Insert(Cmd)
        CmdExecution = True
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeBuildInsert", Me)
        Set Cmd.Connection = Connection
        Cmd.CommandOperation = cmdExec
        Cmd.CommandType = dsTable
        Cmd.CommandParameters = Empty
        Dim IsDef_Funcion : IsDef_Funcion = CCIsDefined("Funcion_" & CurrentRow, "Form")
        Dim IsDef_Privilegios : IsDef_Privilegios = CCIsDefined("Privilegios_" & CurrentRow, "Form")
        Dim IsDef_Planta : IsDef_Planta = CCIsDefined("Planta_" & CurrentRow, "Form")
        Dim IsDef_Email : IsDef_Email = CCIsDefined("Email_" & CurrentRow, "Form")
        Dim IsDef_Activo : IsDef_Activo = CCIsDefined("Activo_" & CurrentRow, "Form")
        If Not InsertOmitIfEmpty("Funcion") Or IsDef_Funcion Then Cmd.AddSQLStrings "Funcion", Connection.ToSQL(Funcion, Funcion.DataType)
        If Not InsertOmitIfEmpty("Privilegios") Or IsDef_Privilegios Then Cmd.AddSQLStrings "Privilegios", Connection.ToSQL(Privilegios, Privilegios.DataType)
        If Not InsertOmitIfEmpty("Planta") Or IsDef_Planta Then Cmd.AddSQLStrings "Planta", Connection.ToSQL(Planta, Planta.DataType)
        If Not InsertOmitIfEmpty("Email") Or IsDef_Email Then Cmd.AddSQLStrings "Email", Connection.ToSQL(Email, Email.DataType)
        If Not InsertOmitIfEmpty("Activo") Or IsDef_Activo Then Cmd.AddSQLStrings "Activo", Connection.ToSQL(Activo, Activo.DataType)
        CmdExecution = Cmd.PrepareSQL("Insert", "Operadores", Empty)
        CCSEventResult = CCRaiseEvent(CCSEvents, "BeforeExecuteInsert", Me)
        If Errors.Count = 0  And CmdExecution Then
            Cmd.Exec(Errors)
            CCSEventResult = CCRaiseEvent(CCSEvents, "AfterExecuteInsert", Me)
        End If
    End Sub
'End Insert Method

End Class 'End Operadores1DataSource Class @2-A61BA892


%>
