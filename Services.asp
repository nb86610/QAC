<%
'Services Body @1-6A862803
Function CCCreateService(ServiceType, ServiceID, DataSource, Fields,Headers)
    Dim I
	Dim Service : Set Service = New clsService
    With Service
        .ServiceType = ServiceType
		.ServiceId = ServiceId
		.DataSource = DataSource
		.Fields=Fields
		If Not Headers Is Nothing Then 
		    For I = 0 To Ubound(Headers) Step 2
    	  		.Header.Add Headers(I), Headers(I+1)
    		Next
		End If
	 End With
	Set CCCreateService = Service
End Function

Class clsService
  Public SQL
  Public HTTPHeaders
  Public Errors
  	
  Private mDataSource
  Private mFields
  Private mConnection
  Private mServiceId
  Private mOutputFormatter
  Private ArNames
  Private mOutputAssoc
  Private mServiceType
  Private mResultArray

  Private mInputType
  Private mInputData
  Private mInputParser

  Private DS
  Private ShownRecords

  Private Sub Class_Initialize()
	mInputType =  ccsPost
	Set mInputParser = Nothing
	Set mOutputFormatter = Nothing
        Set HTTPHeaders = Server.CreateObject("Scripting.Dictionary")
	Set Errors = New clsErrors
	Set mFields = Nothing
	Set mConnection = Nothing
	Set mDataSource = Nothing
	Redim mResultArray(0)
	mOutputAssoc = True
	Set DS = New clsDataSource
  End Sub

  Private Sub Class_Terminate()
    Dim I
    Set Errors = Nothing
    Set HTTPHeaders = Nothing
    Set mFields = Nothing
	Set mDataSource = Nothing
	For I=0 To UBound(mResultArray)
      Set mResultArray(I) = Nothing
	Next
	Set mResultArray = Nothing
  End Sub


  Function AddDatabaseField(Field)
    If mFields is Nothing Then _ 
		Set mFields=New clsFields
	If 	TypeName(Field) = "clsField" Then 
		mFields.AddFields(Field)
	Else 
	   mFields.AddFields(CCCreateField(Field, Field, ccsText, Empty, DS)) 
	End If
  End Function

  Property Get InputParser
    Set InputParser = mInputParser
  End Property
  Property Let InputParser(NewValue)
    If Not mInputParser is Nothing Then _
	Set mInputParser = Nothiing
    Set mInputParser = NewValue
  End Property


  Property Get OutputFormatter
    Set OutputFormatter = mOutputFormatter
  End Property
  Property Let OutputFormatter(NewValue)
    If Not mOutputFormatter is Nothing Then _
	Set mOutputFormatter = Nothiing
    Set mOutputFormatter = NewValue
  End Property
	
  Property Get ServiceId
     ServiceId = mServiceId
  End Property
  Property Let ServiceId(NewValue)
     mServiceId = ServiceId
  End Property

  Property Get OutputAssoc
     OutputAssoc = mOutputAssoc
  End Property
  Property Let OutputAssoc(NewValue)
     mOutputAssoc = NewValue
  End Property

  Property Get Connection
    Set Connection = mConnection
  End Property
  Property Let Connection(NewValue)
    If Not mConnection Is Nothing Then _
		Set mConnection = Nothing
    Set mConnection = NewValue
	mConnection.Open
  End Property


  Property Get DataSource
    Set DataSource = mDataSource
  End Property
  Property Let DataSource(NewValue)
    If Not mDataSource Is Nothing Then _
		Set mDataSource = Nothing
    Set mDataSource = NewValue
  End Property

  Property Get Fields
    Set Fields = mFields
  End Property
  Property Let Fields(NewValue)
    Set mFields = NewValue
  End Property
  
  Property Get ServiceType
    ServiceType = mServiceType
  End Property
  Property Let ServiceType(NewValue)
    ServiceType = NewValue
  End Property
  

  Public Function Execute
     GetDBData
     If Errors.Count = 0 Then
        If Not OutputFormatter Is Nothing Then 
          Execute =  OutputFormatter.Format(mResultArray)
        Else
       	  Execute = ""
        End If
     Else
        Execute =""  
     End  If
  End Function
		
	

  Public Sub GetDBData
    Dim cmdErrors,Ar,Count,RowCount,fld, Dic, N
    mDataSource.Connection.Open
    Set cmdErrors = new clsErrors
    Set DS = mDataSource.Exec(cmdErrors)
    If cmdErrors.Count > 0 Then 
      Errors.AddErrors cmdErrors.Errors
    Else
      RowCount = 0
      ShownRecords = 0
      While HasNextRow()
        Redim Preserve mResultArray(RowCount)
        If OutputAssoc Then 
          Set Dic = CreateObject("Scripting.Dictionary")
        Else 
          ReDim Ar(Fields.Count-1)
          N = LBound(Ar)
        End If
        If Fields is Nothing Then 
          'use all Recordset
          For Each Fld in DS.Recordset.Fields
            If OutputAssoc Then 
              Dic.Add  Fld.Name, DS.Recordset(Fld.Name).Value
            Else 
              Ar(N) = DS.Recordset(Fld.Name)
              N = N + 1
            End If
          Next
        Else 
          Fields.InitEnum
          While Not Fields.EndOfEnum
            Set Fld = Fields.NextItem
            If OutputAssoc Then 
              If Fld.DBFieldName <> "" Then 
                Dic.Add  Fld.Name, DS.Recordset(Fld.DBFieldName).Value
              Else
                Dic.Add  Fld.Name, Empty
              End If
            Else 
	      If Fld.DBFieldName <> "" Then 
	        Ar(N) = DS.Recordset(Fld.DBFieldName).Value
              Else
                Ar(N) = Empty
              End If
 	      N = N + 1
	    End If
          Wend	
        End If
        If OutputAssoc Then 
          Set mResultArray(RowCount) = Dic
        Else 
          mResultArray(RowCount) = Ar
        End If 
        RowCount = RowCount + 1
        ShownRecords = RowCount
        DS.Recordset.MoveNext
      Wend
    End If

  End Sub
 
  Private Function HasNextRow()
    If Not IsEmpty(mDataSource.PageSize) And mDataSource.PageSize > 0 Then
       HasNextRow = Not DS.Recordset.EOF  And ShownRecords < mDataSource.PageSize
    Else 
       HasNextRow = Not DS.Recordset.EOF
    End If 
  End Function


  Private Sub DisplayHeaders
   Dim HeaderName
   For Each HeaderName In Headers
     Response.AddHeader HeaderName, Headers(HeaderName)
   Next
  End Sub
	
	

 Public Sub Operation(Command, Oper)
   If Oper = "Delete" Then 
     DataSource.Delete(Command)
   ElseIf Oper = "Insert" Then
     DataSource.Insert(Command)
   ElseIf Oper = "Update" Then
     DataSource.Update(Command)
   End If	
 End Sub


 Public Sub SetInputData(InputType)
    Dim i
    mInputType = InputType
    mInputData = ""
    If InputType =  ccsPost Then 
	  If IsMutipartEncoding Then
	    mInputData = BinToText(Request.BinaryRead(Request.TotalBytes))
      Else
	    mInputData = Request.Form.Item
	  End If
    End If

    If Not InputParser is Nothing Then _
      InputParser.Parse(mInputData)
 End Sub

  Public Function GetInputData(Field)
    If mInputParser is Nothing Then 
	GetInputData = Empty
    Else 	
	GetInputData = mInputParser.GetValue(Field)
    End If
  End Function

  Private Function BinToText(DataBin) 
    Dim objRS :
    Set objRS = CreateObject("ADODB.Recordset")
    objRS.Fields.Append "txt",201,Request.TotalBytes
    objRS.Open
    objRS.AddNew
    objRS.Fields("txt").AppendChunk(DataBin)
    objRS.Update
    Dim DataTextRes : DataTextRes= objRS("txt").Value
    objRS.Close
    Set objRS = Nothing
    BinToText = DataTextRes
End Function

End Class




Class clsInputJSONParser
 
   Private JPar

   Private Sub Class_Initialize()
	Set JPar = new CCBJsonParser
   End Sub
	
   Public Function Parse(Text) 
	JPar.JSON = Text
   End Function

   Public Function GetValue(Name) 
	GetValue = JPar.valueFor(Name)
   End Function

   Private Sub Class_Terminate()
	Set JPar = Nothing
   End Sub

End Class



Class clsJSONFormatter
   Private Sub Class_Initialize()
   End Sub

   Private Sub Class_Terminate()
   End Sub

   Public Function Format(Data) 
      Format =  JSONEncode(Data)
   End Function

   Private Function JSONEncode(Arg)
     Dim Keys, j, t
     Dim i : i=0
     Dim o : o=""
     Dim u : u=""
     Dim v : v=""
     Dim z : z=""
     Dim r : r=""
     JSONEncode = ""
     If IsArray(Arg) Then
       o = ""
       For i=LBound(Arg) To UBound(Arg)
         v = JSONEncode(Arg(i))
         If o <> "" Then 	o = o & ","
         o = o & v
       Next
       JSONEncode = "["& o &"]"
     ElseIf TypeName(Arg) = "Dictionary" Then
       o = ""
       If Arg.Count = 0 Then 
  	 JSONEncode = "{}"
       Else
	 Keys  = Arg.Keys   
         For j = 0 To Arg.Count -1 
   	   v = JSONEncode(Arg(Keys(j)))
 	   If o <> "" Then o = o & ","
 	   o = o & """" & Keys(j) & """:" & v
	   JSONEncode = "{" & o & "}"
	 Next
       End If	
     Else 
      t = VarType(Arg)
      Select  Case t
        Case 2,3,4,5,6
          JSONEncode = Replace(CStr(Arg), ",", ".")
        Case 0,7,8
          JSONEncode = """" & Arg & """"
        Case Else 
          JSONEncode = CStr(Arg)
        End Select
      End If
  End Function
End Class

Function CCCreateTreeFormatter(CategoryIdField, SubCategoryIdField, CategoryNameField) 
		Dim F
		F =  New clsTreeFormatter
		With F 
		  .CategoryIdField = CategoryIdField 
		  .SubCategoryIdField = SubCategoryIdField
		  .CategoryNameField = CategoryNameField
        End With
	CreateTreeFormatter = F
End Function

Class clsTreeFormatter
   Private CategotyIdField
   Private CategoryNameField
   Private SubCategoryIdField
   Private JSONFormatter

   Private Sub Class_Initialize()
   	Set JSONFormatter = New clsJSONFormatter
   End Sub

   Private Sub Class_Terminate()
   	Set JSONFormatter = Nothing
   End Sub

   Public Function Format(Data)
      Dim I, N, OutputArray, Row, IsFolder, RowAr
   	  Dim LastCategoryId : LastCategoryId = Empty   
	  ReDim OutputArray(UBound(Data)-LBound(Data)+1) 	

      For Each Row in Data
	    If LastCategoryId <> Row.Item(CategoryIdField) Then
			LastCategoryId = Row.Item(CategoryIdField)
		    IsFolder = IIF(Row.Item(SubCategoryIdField)="",True,False)
			Set RowAr = CreateObject("Scripting.Dictionary")
			RowAr.Add "objectId", LastCategoryId
			RowAr.Add "title", Row.Item(CategoryNameField)
			RowAr.Add "isFolder", IsFolder
			OutputArray(N) = RowAr
			N=N+1
		End If
	  Next
	  ReDim Preserve OutputArray(N) 	
      Format = JSONEncode(OutputArray)
   End Function
End Class   


Class clsOutputFormatter 
   Private Sub Class_Initialize()
   End Sub
	
   Public Function Format(Data) 
	Format = ""
   End Function
End Class

Function CCEscapeDouble(Str)
	CCEscapeDouble= Replace(Str, """","""""")
End Function


Class clsListFormatter 
   Public Function Format(Data) 
     Dim Res, Keys, Line, i, j, l
     Res = "<ul>"
     For i=LBound(Data) To UBound(Data)
       l = "<li>"
       If TypeName(Data(i)) = "Dictionary" Then
         Keys  = Data(i).Keys   
         For j = 0 To Data(i).Count -1 
           l = l & Data(i)(Keys(j))
	 Next
       ElseIf IsArray(Data(i)) Then 
         Line = Data(i)
         For j = LBound(Line) To UBound(Line) 
           l = l & Line(j)
         Next
       Else 
         l =  Data(i)
       End If 
       Res = Res & l & "</li>" 
     Next 
     Format = Res & "</ul>"
   End Function
End Class   


Function CCBuildSnippet(ParamsArr) 
    Dim Output, Param, ControlName, ParamsStr, FeatureName
    CCBuildSnippet = ""
    If ParamsArr.Count > 0 Then 
      Output = "<script language=""JavaScript"" type=""text/javascript"">"
      For Each FeatureName in ParamsArr
	ControlName = ParamsArr(FeatureName)(0)
	ParamsStr = ParamsArr(FeatureName)(1)
	Output = Output & "document.getElementById(""" & ControlName  & """).ccs" &  FeatureName  & "Data = " & ParamsStr & ";"
      Next	
      CCBuildSnippet = Output &	"</script>"
    End If 
End Function


Function CCBuildSnippetTest(ParamsArr) 
    Dim Output, Param, ControlName, ParamsStr, FeatureName, I
    CCBuildSnippet = ""
    If UBound(ParamsArr) > 1 Then 
      FeatureName = ParamsArr(0)
      Output = "<script language=""JavaScript"" type=""text/javascript"">"
      For I =1 To UBound(ParamsArr) Step 2
	ControlName = ParamsArr(I)
	ParamsStr = ParamsArr(I+1)
	Output = Output & "document.getElementById(""" & ControlName  & """).ccs" &  FeatureName  & "Data = " & ParamsStr & ";"
      Next	
      CCBuildSnippet = Output &	"</script>"
    End If 
End Function

Function CCCreateTemplateFormatter(TemplatePath)
    Dim F, Strm
    Set F = New clsTemplateFormatter
    Set Strm = Server.CreateObject("ADODB.Stream")
    Strm.Open
    Strm.Charset = "utf-8"
    Strm.LoadFromFile TemplatePath
    F.Template = Strm.ReadText(adReadAll)
    Strm.Close
    Set Strm = Nothing
    Set CCCreateTemplateFormatter = F
End Function

Class clsTemplateFormatter
   Private mTemplate

   Private Sub Class_Initialize()
   End Sub
	
  Private Sub Class_Terminate()
  End Sub

  Property Let Template(NewValue)
     mTemplate = NewValue 
  End Property

	         
  Public Function Format(Rows) 
    Dim I, J, Row, Tpl,Item, RowBlock, EscItem
    Set HTMLTemplate = new clsTemplate
    Set HTMLTemplate.Cache = TemplatesRepository
    HTMLTemplate.LoadTemplateFromStr(mTemplate)
    Set Tpl=HTMLTemplate.Block("main")
    'HTMLTemplate.PrintBlocks
    Set RowBlock = Tpl.Block("Row")
    For I=LBound(Rows) To UBound(Rows)
      If Not IsEmpty(Rows(I)) Then
        Set Row = Rows(I)
        For Each Item In Row
          EscItem = Replace(Item, "&", "&amp;")
          If IsNull(Row(Item)) Then
            RowBlock.Variable(EscItem) = ""
          Else
            RowBlock.Variable(EscItem) = Replace(Row(Item), "&", "&amp;")
          End If
        Next
        RowBlock.Parse ccsParseAccumulate
      End If
    Next
    Tpl.Parse ccsParseOverwrite
    Format = HTMLTemplate.GetHTML("main")
    Set RowBlock = Nothing
    Set Tpl = Nothing
    Set HTMLTemplate = Nothing
  End Function

End Class


'End Services Body

'Services CCBJsonParser parser @1-24A672B7
'=================================================================================================================
'	CCBJsonParser version 0.1 & CCBJsonParser 0.1
'	Copyright Â© 2007 Cliff Pruitt
'
'	Created  by cpruitt on Thu 06/07/2007 at 16:56:41 EDT
'	http://www.crayoncowboy.com/software
'=================================================================================================================

' 	CCBJsonParser (including the CCBJSONTranslator JScript object) is a simple JSON parser for ASP scripts
' 	The current development version includes only basic functionality and very little error handling
' 	
' 	* CONTENTS *
' 	+	CCBJsonParser						- Parses a JSON formatted Object Representation
' 	+	CCBJSONTranslator (JScript object)	- Helper Object used to manuiplate load JSON representation.
' 											  This object will probably never be used directly & is not documented.
' 	+	newCCBJSONTranslator(<JSON string>)	- Utility function to create a new instance of CCBJSONTranslator in VBScript
' 	+	CCBundefined (constant) 			- A unique constant value used to represent a
' 											  JScript undefined value (used for comparisons & "if" statements)
' 	
' 	* PURPOSE *
' 	CCBJsonParser is intended only to translate a JSON formatted data representation into something usable in
' 	an ASP script.  Currently the entire representation is returned as one or more VBScript Dictionary objects.
' 	This means that, although JSON representations of arrays are accurately parsed into memory, they are returned
' 	as VBScript Safe Arrays.  Instead they are returned as dictionaries with keys 0 - (array.length). This will
' 	probablybe fixed in the future.
' 	
' 	
' 	Using the following few properties and methods you should be able to access any data indicated by a JSON representation.
' 	
' 	
' 	* USAGE *
' 	+ 	After creating an instance of the Object, the JSON representation (string) is assigned to the objects .JSON property.
' 		(This populates the contents of the CCBJSONTranslator object but does not actually create any VBScript objects)
' 		
' 	+	After the JSON property is populated, VBScript can access the data in one of two ways.
' 		1.	Call the .parse method on the object.  This will load the complete contents of the JSON
' 			representation into memory and assign it to the objects .dictionary property.
' 		2.	Use the .valueFor() method to fetch a sub-section of the JSON representation.
' 		
' 	* OBJECT PROPERTIES & METHODS *
' 	
' 	-- PROPERTIES --
' 
' 	+	version		: (string) (r/o) returns the version of the class
' 	+	className	: (string) (r/o) returns the name of the class (CCBJsonParser)
' 	+	JSON		: (string) sets / returns the JSON representation of the data
' 	+	dictionary	: (dictionary) (r/o) returns a dictionary representation of the entire JSON data structure
' 	+	description	: (string) (r/o) returns an HTML description of the object instace comparing the JSON string to the obj.dictionary value.
' 	
' 	-- METHODS --
' 	+	parse					: parses the JSON representation of the data, builds a collection of nested dictionaries
' 					  			  and sets the value of the .dictionary property
' 	+	valueFor(<scopeString>)	: returns a the value of the object indicated by the "scopeString".  If the object is a
' 								  collection, a dictionary is returned, otherwise a primitive (string, integer, boolean, etc...)
' 		- scopeString	: The scopeString argument should be a "path" to the data sub-item in JavaScript dot/array notation format.
' 						  For example, consider the following statement:
' 						
' 							set fName = Obj.valueFor("Sports.Soccer.Grade10.Games[4].Score.Winner")
' 						
' 						  This would Look into the Sports > Soccer > Grade10 > Games array
' 						  From there it would examine the 4th element and get the element's Score object.
' 						  It would then get the "Winner" proerty of the Score object which is what would
' 						  be returned by Obj.valueFor.
' 						
' 	
' 	==== SEE THE EXAMPLE FILE FOR MORE USAGE EXAMPLES ====
' 

'=================================================================================================================



Const CCBundefined = "{7LV65C98O-UOB5-0SDF-MGDN-INLLDDJQ94CD}"
Class CCBJsonParser
	
	
	
	'/-----------------------------  DECLARE VARIABLES  ----------------------------/
	
	Private p_lastError, P_Data, p_json, P_JST
	Private p_className, p_version

	'/-------------------------  INITIALIZE AND TERMINATE  -------------------------/
	
	Private Sub Class_Initialize()
		p_className	= "CCBJsonParser"
		p_version	= "0.1"
		call init()
	End Sub
	
	Private Sub Class_Terminate()
		Set P_Data = Nothing
	End Sub
	
	'	/***  Designated Initializer  ***/
	Private Function init()
		Dim v_out : v_out = true
		p_json = "{}"
		
		If isObject(P_Data) Then Set P_Data = Nothing
		Set P_Data = Server.CreateObject("Scripting.Dictionary")
		P_Data.CompareMode = VBTextCompare '// or = 1
		
		If isObject(P_JST) Then Set P_JST = Nothing
		Set P_JST = newCCBJSONTranslator(p_json)
		
		init = v_out
	End Function
	
	'/----------------------------- PROPERTY ACCESSORS  ----------------------------/
	
	'--	GET version from p_version
	Public Property Get version()
		version = p_version
	End Property
	
	'--	GET JSON from p_JSON
	Public Property Get JSON()
		JSON = p_JSON
	End Property

	'--	LET p_JSON = JSON
	Public Property Let JSON(newValue)
		If trim(newValue) & "" = "" Then newValue = "{}"
		P_JST.setData(newValue)
		p_JSON = newValue
	End Property
	
	'--	GET dictionary from p_dictionary
	Public Property Get dictionary()
		Set dictionary = P_Data
	End Property
	
	'/------------------------------  OBJECT METHODS  ------------------------------/
	
	Public Function parse()
		Dim v_out : v_out = true
		If isObject(P_Data) Then Set P_Data = Nothing
		Set P_Data = JSObjectToVBDictionary(P_JST)
		parse = v_out
	End Function
	
	Public Function description()
		Dim v_out : v_out = null
		v_out = v_out & "<h1>" & p_className & " object: Version " & p_version & "</h1>"
		
		v_out = v_out & li("JSON: " & code(Me.JSON))
		v_out = v_out & li("DICTIONARY REPRESENTATION:" & describeDict(P_Data))
		
		v_out = ul(v_out)
		
		description = v_out
	End Function
	
	Public Function valueFor(scopeString)
		Dim v_out, typ
		
		scopeString = scopeString & ""
		If inStr(scopeString, ".") > 0 Then
			scopeString = replace(scopeString, ".", "['", 1,1)
			scopeString = replace(scopeString, ".", "']['")
			scopeString = scopeString & "']"
		End If
		
		typ = P_JST.typeForKey(scopeString)
		If typ = "object" Then
			Set	valueFor = JSObjecttoVBDictionaryByScope(P_JST,scopeString)
		Else
			v_out = P_JST.valForKey(scopeString)
			valueFor = v_out
		End If
	End Function
	
	'/-----------------------------  PRIVATE FUNCTIONS  ----------------------------/
	Private Function JSObjectToVBDictionary(JSTranslationObject)
		Set JSObjectToVBDictionary = JSObjecttoVBDictionaryByScope(JSTranslationObject,null)
	End Function
	
	Private Function JSObjecttoVBDictionaryByScope(JT,path)
		Dim v_out, jkeys, k, jkey, typ, O
		Dim a()
		Set v_out = Server.CreateObject("Scripting.Dictionary")
		v_out.CompareMode = VBTextCompare '// or = 1

		jkeys = split(JT.keysFor(path), ",")

		for each k in jkeys
			If trim(k) <> "" Then
				If trim(path & "") <> "" Then
					jkey = path & "['" & k & "']"
				Else
					jkey = k
				End If
				
				typ = JT.typeForKey(jkey)

				Select Case typ
					Case "object", "array"
						Set O = JSObjecttoVBDictionaryByScope(JT,jkey)
						v_out.add k, O
						Set O = Nothing
					Case "arrayx"
						length = JT.lengthForKey(jkey)
						reDim a(length)
						for i = 0 to length - 1
							subKey = jkey & "['" & i & "']"
							innerType = JT.typeForKey(subKey)
							
							Select Case innerType & ""
								Case "object"
									Set a(i) = JSObjecttoVBDictionaryByScope(JT,subKey)
								Case "array"
									Set a(i) = JSObjecttoVBDictionaryByScope(JT,subKey)
								Case Else
									a(i) = typeName(JT.valForKey(subKey))
							End Select
							
							 'typeName(JSObjecttoVBDictionaryByScope(JT,subKey))
						Next
						v_out.add k, a
					Case Else
						v_out.add k, JT.valForKey(jkey)
				End Select
			End If
		Next

		Set JSObjecttoVBDictionaryByScope = v_out
		Set v_out = Nothing
	End Function
	
	Private Function ul(str)
		Dim v_out : v_out = "<ul>" & str & "</ul>"
		ul = v_out
	End Function

	Private Function li(str)
		Dim v_out : v_out = "<li>" & str & "</li>"
		li = v_out
	End Function
	
	Private Function code(str)
		Dim v_out : v_out = "<div style=""white-space:pre; font-family:monospace; margin:0px; padding:4px; font-size:11px;"">" & str & "</div>"
		code = v_out
	End Function
	
	Private Function describeDict(dictObj)
		Dim v_out : v_out = null
		Dim val   : val = ""
		Dim k
		
		
		If dictObj.count > 0 Then
			for each k in dictObj.keys
				val = ""
				If isObject(dictObj(k)) Then
					If lCase(typeName(dictObj(k))) = "dictionary" Then
						val = DescribeDict(dictObj(k))
					Else
						val = "[OBJECT]"
					End If
				ElseIf isArray(dictObj(k)) Then
					val = "Array:"
					for i = 0 to uBound(dictObj(k))
						If isObject(dictObj(k)(i)) or isArray(dictObj(k)(i)) Then
							val = DescribeDict(dictObj(k)(i))
						Else
							val = val & li(dictObj(k)(i))
						End If
						val = ul(val)
					Next
					
				Else
					val = dictObj(k)
				End If
				v_out = v_out & li(k & ": " & val)
			Next
		Else
			v_out = v_out & li("This JSON representation has no properties. JSON = {}")
		End If

		DescribeDict = ul(v_out)
	End Function
	
	'	/***  UNUSED FOR NOW  ***/
	'	//	Get last object Error
	Public Function lastError()
		Dim v_out : v_out = null		
		v_out = p_lastError
		lastError = v_out
	End Function
	
	'	//	Clear Last Error
	Public Function clearError()
		p_lastError = null
	End Function
End Class

'End Services CCBJsonParser parser

'Services CCBJSONTranslator @1-AC6FFB83
%>

<script language="jscript" runat="server">
	function CCBJSONTranslator(inputData){
		var CCBundefined = "{7LV65C98O-UOB5-0SDF-MGDN-INLLDDJQ94CD}"
		var _this = this;
		
		this.init = function(inputData){
			this.setData(inputData);
		}
		
		this.setData = function(strJSON){
			eval('this.data =' + strJSON);
		}
		
		this.data = {}
		
		this.valForKey = function(keyString){
			var v_out;
			if(!keyString){
				val = this.data
			} else {
				try{
					val = eval("this.data." + keyString);
				} catch(e) {
					val = CCBundefined;
				}
			}

			if(_this.isArray(val) && false){
				val = eval("this.data." + keyString)
				var O = {}
				for(i=0; i<val.length; i++){
					O['' + i + ''] = val[i];
				}
				return O;
			} else {
				return val;
			}
		}
		
		this.typeForKey = function(keyString){
			if(_this.valForKey(keyString) == CCBundefined){
				return typeof(undefined);
			} else if(_this.isArray(_this.valForKey(keyString))){
				return 'array'
			} else {
				return typeof(_this.valForKey(keyString));
			}
			
			
		}
		
		this.lengthForKey = function(keyString){
			var val = _this.valForKey(keyString);
			if(_this.isArray(val)){
				return val.length;
			} else {
				val = eval("this.data." + keyString);
				return 0;
			}
			
		}

		this.keysFor = function(dataTarget){
			if(!dataTarget) {
				dataTarget = this.data;
			} else {
				dataTarget = eval("this.data." + dataTarget);
			}
			
			var v_out = '';
			for(itm in dataTarget){
				v_out += itm + ',';
			}
			return v_out;
		}
		
		this.isArray = function(obj){
            return !(obj != null && obj.constructor && obj.constructor.toString().indexOf("Array") == -1);
		}
		
		this.init(inputData);
	}

	function newCCBJSONTranslator(data){
		return new CCBJSONTranslator(data);
	}
</script>
<% 
'End Services CCBJSONTranslator

%>
