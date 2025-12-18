VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' WebRequest v4.1.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' `WebRequest` is used to create detailed requests
' (including formatting, querystrings, headers, cookies, and much more).
'
' Usage:
' ```VB.net
' Dim Request As New WebRequest
' Request.Resource = "users/{Id}"
'
' Request.Method = WebMethod.HttpPut
' Request.RequestFormat = WebFormat.UrlEncoded
' Request.ResponseFormat = WebFormat.Json
'
' Dim Body As New Dictionary
' Body.Add "name", "Tim"
' Body.Add "project", "VBA-Web"
' Set Request.Body = Body
'
' Request.AddUrlSegment "Id", 123
' Request.AddQuerystringParam "api_key", "abcd"
' Request.AddHeader "Authorization", "Token ..."
'
' ' -> PUT (Client.BaseUrl)users/123?api_key=abcd
' '    Authorization: Token ...
' '
' '    name=Tim&project=VBA-Web
' ```
'
' Errors:
' 11020 / 80042b0c / -2147210484 - Cannot add body parameter to non-Dictionary
' 11021 / 80042b0d / -2147210483 - Cannot add body node to non-XML body
'
' @class WebRequest
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' VBA-Git Annotations
' https://github.com/VBA-Tools-v2/VBA-Git | https://radiuscore.co.nz
'
' @excludeobfuscation
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' RubberDuck Annotations
' https://rubberduckvba.com/ | https://github.com/rubberduck-vba/Rubberduck/
'
'@folder VBA-Web
'@ignoremodule
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Type TWebRequest
    Resource As String
    Method As VbWebMethod
    Headers As Collection
    QuerystringParams As Collection
    UrlSegments As Dictionary
    Cookies As Collection
    UserAgent As String
    RequestFormat As VbWebFormat
    ResponseFormat As VbWebFormat
    CustomRequestFormat As String
    CustomResponseFormat As String
    Body As Variant
    ConvertedBody As Variant
    ContentType As String
    Accept As String
    AcceptEncoding As String
    ContentLength As Long
    Id As String
End Type

Private This As TWebRequest

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

''
' Set the request's portion of the url to be appended to the client's BaseUrl.
' Can include Url Segments for dynamic values
' and Querystring parameters are smart enough to be appended to existing querystring
' (or added to resource if there isn't an existing querystring).
'
' @example
' ```VB.net
' Dim Client As New WebClient
' Client.BaseUrl = "https://api.example.com/"
'
' Dim Request As New WebRequest
' Request.Resource = "messages"
'
' ' -> Url: https://api.example.com/messages
'
' Request.Resource = "messages/{id}?a=1"
' Request.AddUrlSegment "id", 123
' Request.AddQuerystringParam "b", 2
'
' ' -> Url: https://api.example.com/messages/123?a=1&b=2
' ```
'
' @property Resource
' @type String
''
Public Property Get Resource() As String
    Resource = This.Resource
End Property
Public Property Let Resource(Value As String)
    This.Resource = Value
End Property

''
' Set the HTTP method to be used for the request:
' GET, POST, PUT, PATCH, DELETE
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Method = VbWebMethod.vbWebHttpGet
' Request.Method = VbWebMethod.vbWebHttpPost
' ' or HttpPut / HttpPatch / HttpDelete
' ```
'
' @property Method
' @type VbWebMethod
''
Public Property Get Method() As VbWebMethod
    Method = This.Method
End Property
Public Property Let Method(Value As VbWebMethod)
    This.Method = Value
End Property

''
' _Note_ To add headers, use [`AddHeader`](#/WebRequest/AddHeader).
'
' `Collection` of Headers to include with request,
' stored as `KeyValue` (`Dictionary: {Key: "...", Value: "..."}`).
'
' @property Headers
' @type Collection
''
Public Property Get Headers() As Collection
    Set Headers = This.Headers
End Property
Public Property Set Headers(Value As Collection)
    Set This.Headers = Value
End Property

''
' _Note_ To add querystring parameters, use [`AddQuerystringParam`](#/WebRequest/AddQuerystringParam).
'
' `Collection` of querystring parameters to include with request,
' stored as `KeyValue` (`Dictionary: {Key: "...", Value: "..."}`).
'
' @property QuerystringParams
' @type Collection
''
Public Property Get QuerystringParams() As Collection
    Set QuerystringParams = This.QuerystringParams
End Property
Public Property Set QuerystringParams(Value As Collection)
    Set This.QuerystringParams = Value
End Property

''
' _Note_ To add Url Segments, use [`AddUrlSegment`](#/WebRequest/AddUrlSegment)
'
' Url Segments are used to easily add dynamic values to `Resource`.
' Create a Url Segement in `Resource` with curly brackets and then
' replace with dynamic value with [`AddUrlSegment`](#AddUrlSegment).
'
' @example
' ```VB.net
' Dim Request As New WebRequest
'
' Dim User As String
' Dim Id As Long
' User = "Tim"
' Id = 123
'
' ' OK: Use string concatenation for dynamic values
' Request.Resource = User & "/messages/" & Id
'
' ' BETTER: Use Url Segments for dynamic values
' Request.Resource = "{User}/messages/{Id}"
' Request.AddUrlSegment "User", User
' Request.AddUrlSegment "Id", Id
'
' Request.FormattedResource ' = "Tim/messages/123"
' ```
'
' @property UrlSegments
' @type Dictionary
''
Public Property Get UrlSegments() As Dictionary
    Set UrlSegments = This.UrlSegments
End Property
Public Property Set UrlSegments(Value As Dictionary)
    Set This.UrlSegments = Value
End Property

''
' _Note_ To add cookies, use [`AddCookie`](#/WebRequest/AddCookie).
'
' `Collection` of cookies to include with request,
' stored as `KeyValue` (`Dictionary: {Key: "...", Value: "..."}`).
'
' @property Cookies
' @type Collection
''
Public Property Get Cookies() As Collection
    Set Cookies = This.Cookies
End Property
Public Property Set Cookies(Value As Collection)
    Set This.Cookies = Value
End Property

''
' User agent to use with request
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.UserAgent = "Mozilla/5.0"
'
' ' -> (Header) User-Agent: Mozilla/5.0
' ```
'
' @property UserAgent
' @type String
' @default "VBA-Web v#.#.# (https://github.com/VBA-tools/VBA-Web)"
''
Public Property Get UserAgent() As String
    UserAgent = This.UserAgent
End Property
Public Property Let UserAgent(Value As String)
    This.UserAgent = Value
End Property

''
' Set `RequestFormat`, `ResponseFormat`, and `Content-Type` and `Accept`
' headers for the `WebRequest`
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Format = VbWebFormat.vbWebJson
' ' -> Request.RequestFormat = VbWebFormat.vbWebJson
' '    Request.ResponseFormat = VbWebFormat.vbWebJson
' '    (Header) Content-Type: application/json
' '    (Header) Accept: application/json
' ```
'
' @property Format
' @type WebFormat
''
Public Property Get Format() As VbWebFormat
    Format = Me.RequestFormat
End Property
Public Property Let Format(Value As VbWebFormat)
    Me.RequestFormat = Value
    Me.ResponseFormat = Value
End Property

''
' Set the format to use for converting the response `Body` to string and for the `Content-Type` header
'
' _Note_ If `WebFormat.Custom` is used, the [`CustomRequestFormat`](#/WebRequest/CustomRequestFormat) must be set.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Body = Array("A", "B", "C")
'
' Request.RequestFormat = VbWebFormat.vbWebJson
'
' ' -> (Header) Content-Type: application/json
' ' -> Convert Body to JSON string
' Request.Body ' = "["A","B","C"]"
' ```
'
' @property RequestFormat
' @type WebFormat
' @default WebFormat.Json
''
Public Property Get RequestFormat() As VbWebFormat
    RequestFormat = This.RequestFormat
End Property
Public Property Let RequestFormat(Value As VbWebFormat)
    If Value <> This.RequestFormat Then
        This.RequestFormat = Value
        
        ' Clear cached converted body
        This.ConvertedBody = Empty
    End If
End Property

''
' Set the format to use for converting the response `Content` to `Data` and for the `Accept` header
'
' _Note_ If `WebFormat.Custom` is used, the [`CustomResponseFormat`](#/WebRequest/CustomResponseFormat) must be set.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.ResponseFormat = VbWebFormat.vbWebJson
'
' ' -> (Header) Accept: application/json
'
' Dim Response As WebResponse
' ' ... from Execute
' Response.Content = "{""message"":""Howdy!""}"
'
' ' -> Parse Content to JSON Dictionary
' Debug.Print Response.Data("message") ' -> "Howdy!"
' ```
'
' @property ResponseFormat
' @type WebFormat
' @default WebFormat.Json
''
Public Property Get ResponseFormat() As VbWebFormat
    ResponseFormat = This.ResponseFormat
End Property
Public Property Let ResponseFormat(Value As VbWebFormat)
    If Value <> This.ResponseFormat Then
        This.ResponseFormat = Value

        ' Clear cached converted body
        This.ConvertedBody = Empty
    End If
End Property

''
' Use converter registered with [`WebHelpers.RegisterConverter`](#/WebHelpers/RegisterConverter)
' to convert `Body` to string and set `Content-Type` header.
'
' (Automatically sets `RequestFormat` to `WebFormat.Custom`)
'
' @example
' ```VB.net
' WebHelpers.RegisterConverter "csv", "text/csv", "Module.ConvertToCsv", "Module.ParseCsv"
'
' Dim Request As New WebRequest
' Request.CustomRequestFormat = "csv"
'
' ' -> (Header) Content-Type: text/csv
' ' -> Body converted to string with Module.ConvertToCsv
' ```
'
' @property CustomRequestFormat
' @type String
''
Public Property Get CustomRequestFormat() As String
    CustomRequestFormat = This.CustomRequestFormat
End Property
Public Property Let CustomRequestFormat(Value As String)
    If Not Value = This.CustomRequestFormat Then
        This.CustomRequestFormat = Value

        ' Clear cached converted body
        This.ConvertedBody = Empty

        If Not Value = vbNullString Then
            This.RequestFormat = VbWebFormat.vbWebFormatCustom
        End If
    End If
End Property

''
' Use converter registered with [`WebHelpers.RegisterConverter`](#/WebHelpers/RegisterConverter)
' to convert the response `Content` to `Data` and set `Accept` header.
'
' (Automatically sets `ResponseFormat` to `WebFormat.Custom`)
'
' @example
' ```VB.net
' WebHelpers.RegisterConverter "csv", "text/csv", "Module.ConvertToCsv", "Module.ParseCsv"
'
' Dim Request As New WebRequest
' Request.CustomResponseFormat = "csv"
'
' ' -> (Header) Accept: text/csv
' ' -> WebResponse Content converted Data with Module.ParseCsv
' ```
'
' @property CustomResponseFormat
' @type String
''
Public Property Get CustomResponseFormat() As String
    CustomResponseFormat = This.CustomResponseFormat
End Property
Public Property Let CustomResponseFormat(Value As String)
    If Not Value = This.CustomResponseFormat Then
        This.CustomResponseFormat = Value

        ' Clear cached converted body
        This.ConvertedBody = Empty

        If Not Value = vbNullString Then
            ResponseFormat = VbWebFormat.vbWebFormatCustom
        End If
    End If
End Property

''
' Set automatically from `RequestFormat` or `CustomRequestFormat`,
' but can be overriden to set `Content-Type` header for request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.ContentType = "text/csv"
'
' ' -> (Header) Content-Type: text/csv
' ```
'
' @property ContentType
' @type String
' @default Media-type of request format
''
Public Property Get ContentType() As String
    If Not This.ContentType = vbNullString Then
        ContentType = This.ContentType
    Else
        ContentType = WebHelpers.FormatToMediaType(Me.RequestFormat, Me.CustomRequestFormat)
    End If
End Property
Public Property Let ContentType(Value As String)
    This.ContentType = Value
End Property

''
' Set automatically from `ResponseFormat` or `CustomResponseFormat`,
' but can be overriden to set `Accept` header for request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Accept = "text/csv"
'
' ' -> (Header) Accept: text/csv
' ```
'
' @property Accept
' @type String
' @default Media-type of response format
''
Public Property Get Accept() As String
    If Not This.Accept = vbNullString Then
        Accept = This.Accept
    Else
        Accept = WebHelpers.FormatToMediaType(Me.ResponseFormat, Me.CustomResponseFormat)
    End If
End Property
Public Property Let Accept(Value As String)
    This.Accept = Value
End Property

''
' WinHTTP does not support decompression at this time (Jan. 2019).
' If not Accept-Encoding is passed to the server, [RFC 7231](https://tools.ietf.org/html/rfc7231#section-5.3.4)
' states that "any content-coding is considered acceptable by the user agent"
' -> Explicitly set Accept-Encoding
'
' cURL supports --compressed, which automatically decompresses gzip and other compressed responses
' -> If AcceptEncoding != "identity", enable --compressed flag
''
Public Property Get AcceptEncoding() As String
    If Not This.AcceptEncoding = vbNullString Then
        AcceptEncoding = This.AcceptEncoding
    Else
        AcceptEncoding = "identity"
    End If
End Property
Public Property Let AcceptEncoding(Value As String)
    This.AcceptEncoding = Value
End Property

''
' Set automatically by length of `Body`,
' but can be overriden to set `Content-Length` header for request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.ContentLength = 200
'
' ' -> (Header) Content-Length: 200
' ```
'
' @property ContentLength
' @type Long
' @default Length of `Body`
''
Public Property Get ContentLength() As Long
    If This.ContentLength >= 0 Then
        ContentLength = This.ContentLength
    Else
        ContentLength = VBA.Len(Me.Body)
    End If
End Property
Public Property Let ContentLength(Value As Long)
    This.ContentLength = Value
End Property

''
' - Get: Body value converted to string using `RequestFormat` or `CustomRequestFormat`
' - Let: Use `String` or `Array` for Body
' - Set: Use `Collection`, `Dictionary`, or `Object` for Body
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.RequestFormat = WebFormat.Json
'
' ' Let: String|Array
' Request.Body = "text"
' Debug.Print Request.Body ' -> "text"
'
' Request.Body = Array("A", "B", "C")
' Debug.Print Request.Body ' -> "["A","B","C"]"
'
' ' Set: Collection|Dictionary|Object
' Dim Body As Object
' Set Body = New Collection
' Body.Add "Howdy!"
' Set Request.Body = Body
' Debug.Print Request.Body ' -> "["Howdy!"]"
'
' Set Body = New Dictionary
' Body.Add "a", 123
' Body.Add "b", 456
' Set Request.Body = Body
' Debug.Print Request.Body ' -> "{"a":123,"b":456}"
' ```
'
' @property Body
' @type String|Array|Collection|Dictionary|Variant
''
Public Property Get Body() As Variant
    If Not VBA.IsEmpty(This.Body) Then
        If VBA.VarType(This.Body) = vbString And Not Me.RequestFormat = vbWebFormatBinary Then
            Body = This.Body
        ElseIf VBA.IsEmpty(This.ConvertedBody) Then
            ' Convert body and cache
            This.ConvertedBody = WebHelpers.ConvertToFormat(This.Body, Me.RequestFormat, Me.CustomRequestFormat)
            Body = This.ConvertedBody
        Else
            Body = This.ConvertedBody
        End If
    End If
End Property
Public Property Let Body(Value As Variant)
    This.ConvertedBody = Empty
    This.Body = Value
End Property
Public Property Set Body(Value As Variant)
    This.ConvertedBody = Empty
    Set This.Body = Value
End Property

''
' Get `Resource` with Url Segments replaced and Querystring added.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Resource = "examples/{Id}"
' Request.AddUrlSegment "Id", 123
' Request.AddQuerystringParam "message", "Hello"
'
' Debug.Print Request.FormattedResource ' -> "examples/123?message=Hello"
' ```
'
' @property FormattedResource
' @type String
''
Public Property Get FormattedResource() As String
    Dim web_Segment As Variant
    Dim web_Encoding As VbUrlEncodingMode

    FormattedResource = Me.Resource

    ' Replace url segments
    For Each web_Segment In Me.UrlSegments.Keys
        FormattedResource = VBA.Replace(FormattedResource, "{" & web_Segment & "}", WebHelpers.UrlEncode(Me.UrlSegments.Item(web_Segment)))
    Next web_Segment

    ' Add querystring
    If Me.QuerystringParams.Count > 0 Then
        If VBA.InStr(FormattedResource, "?") <= 0 Then
            FormattedResource = FormattedResource & "?"
        Else
            FormattedResource = FormattedResource & "&"
        End If

        ' For querystrings, W3C defines form-urlencoded as the required encoding,
        ' but the treatment of space -> "+" (rather than "%20") can cause issues
        '
        ' If the request format is explicitly form-urlencoded, use FormUrlEncoding (space -> "+")
        ' otherwise, use subset of RFC 3986 and form-urlencoded that should work for both cases (space -> "%20")
        If Me.RequestFormat = VbWebFormat.vbWebFormatFormUrlEncoded Then
            web_Encoding = VbUrlEncodingMode.vbUrlEncodingForm
        Else
            web_Encoding = VbUrlEncodingMode.vbUrlEncodingQuery
        End If
        FormattedResource = FormattedResource & WebHelpers.ConvertToUrlEncoded(Me.QuerystringParams, EncodingMode:=web_Encoding)
    End If
End Property

''
' @internal
' @property Id
' @type String
''
Public Property Get Id() As String
    If This.Id = vbNullString Then: This.Id = WebHelpers.CreateNonce
    Id = This.Id
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Add header to be sent with request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.AddHeader "Authentication", "Bearer ..."
'
' ' -> (Header) Authorization: Bearer ...
' ```
'
' @method AddHeader
' @param {String} Key
' @param {Variant} Value
''
Public Sub AddHeader(Key As String, Value As Variant)
    Me.Headers.Add WebHelpers.CreateKeyValue(Key, Value)
End Sub

''
' Add/replace header to be sent with request.
' `SetHeader` should be used for headers that can only be included once with a request
' (e.g. Authorization, Content-Type, etc.).
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.AddHeader "Authorization", "A..."
' Request.AddHeader "Authorization", "B..."
'
' ' -> Headers:
' '    Authorization: A...
' '    Authorization: B...
'
' Request.SetHeader "Authorization", "C..."
'
' ' -> Headers:
' '    Authorization: C...
' ```
'
' @method SetHeader
' @param {String} Key
' @param {Variant} Value
''
Public Sub SetHeader(Key As String, Value As Variant)
    WebHelpers.AddOrReplaceInKeyValues Me.Headers, Key, Value
End Sub

''
' Url Segments are used to easily add dynamic values to `Resource`.
' Create a Url Segement in `Resource` with curly brackets and then
' replace with dynamic value with `AddUrlSegment`.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Dim User As String
' Dim Id As Long
'
' User = "Tim"
' Id = 123
'
' ' OK: Use string concatenation for dynamic values
' Request.Resource = User & "/messages/" & Id
'
' ' BETTER: Use Url Segments for dynamic values
' Request.Resource = "{User}/messages/{Id}"
' Request.AddUrlSegment "User", User
' Request.AddUrlSegment "Id", Id
'
' Debug.Print Request.FormattedResource ' > "Tim/messages/123"
' ```
'
' @method AddUrlSegment
' @param {String} Key
' @param {String} Value
''
Public Sub AddUrlSegment(Segment As String, Value As Variant)
    Me.UrlSegments.Item(Segment) = Value
End Sub

''
' Add querysting parameter to be used in `FormattedResource` for request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Resource = "messages"
' Request.AddQuerystringParam "from", "Tim"
'
' Request.FormattedResource ' = "messages?from=Tim"
' ```
'
' @method AddQuerystringParam
' @param {String} Key
' @param {Variant} Value
''
Public Sub AddQuerystringParam(Key As String, Value As Variant)
    Me.QuerystringParams.Add WebHelpers.CreateKeyValue(Key, Value)
End Sub

''
' Add cookie to be sent with request.
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.AddCookie "a", "abc"
' Request.AddCookie "b", 123
'
' ' -> (Header) Cookie: a=abc; b=123;
' ```
'
' @method AddCookie
' @param {String} Key
' @param {Variant} Value
''
Public Sub AddCookie(Key As String, Value As Variant)
    Me.Cookies.Add WebHelpers.CreateKeyValue( _
        web_EncodeCookieName(Key), _
        WebHelpers.UrlEncode(Value, EncodingMode:=VbUrlEncodingMode.vbUrlEncodingCookie) _
    )
End Sub

''
' Add XML Node to `Body`.
'
' `Body` must be a `Dictionary` (if it's an `Array` or `Collection` an error is thrown).
' `WebFormat` must be `XML`.
'
' @example
' ```VB.net
' Dim Request as New WebRequest
' Request.Format = WebFormat.XML
'
' Request.AddBodyParameter "a", 123, "alphabet"
' Request.AddBodyParameter "b", 456, "alphabet"
'
' Debug.Print Request.Body ' -> "<alphabet><a>123</a><b>456</b></alphabet>"
' ```
'
' @method AddBodyNode
' @param {String} Name
' @param {Variant} Value
' @param {ParentNodeName} String
' @throws 11020 / 80042b0c / -2147210484 - Cannot add body parameter to non-Dictionary
'         11021 / 80042b0d / -2147210483 - Cannot add body node to non-XML body
''
Public Sub AddBodyNode(ByVal Name As String, Optional ByVal Value As Variant = Null, _
                       Optional ByVal ParentNode As String, _
                       Optional ByVal AttributeName As String, Optional ByVal AttributeValue As String)
    Dim web_Attributes As Collection
    Dim web_ErrorDescription As String
    
    If Not Me.Format = VbWebFormat.vbWebFormatXml Then
        web_ErrorDescription = "Can only add body node to an XML body. Set Format = `WebFormat.XML` and try again."
        
        WebHelpers.LogError web_ErrorDescription, "WebRequest.AddBodyNode", 11021 + vbObjectError
        Err.Raise 11021 + vbObjectError, "WebRequest.AddBodyNode", web_ErrorDescription
    Else
        If VBA.IsEmpty(This.Body) Then
            Set This.Body = CreateNode(vbNullString)
        ElseIf Not TypeOf This.Body Is Dictionary Then
            web_ErrorDescription = "Cannot add body parameter to non-Dictionary Body (existing Body must be of type Dictionary)"
    
            WebHelpers.LogError web_ErrorDescription, "WebRequest.AddBodyNode", 11020 + vbObjectError
            Err.Raise 11020 + vbObjectError, "WebRequest.AddBodyNode", web_ErrorDescription
        End If
        
        ' Validate required dictionary keys exist (they won't be present if the body was created before Format = `WebFormat.XML`).
        If Not This.Body.Exists("nodeName") Or Not This.Body.Exists("childNodes") Then
            web_ErrorDescription = "Can only add body node to an XML body. Set Format = `WebFormat.XML`, clear the existing body (Body = Empty) and try again."
            
            WebHelpers.LogError web_ErrorDescription, "WebRequest.AddBodyNode", 11021 + vbObjectError
            Err.Raise 11021 + vbObjectError, "WebRequest.AddBodyNode", web_ErrorDescription
        End If
        
        ' Set Parent Node name if given.
        If Not ParentNode = vbNullString Then This.Body.Item("nodeName") = ParentNode
        
        ' Add node to 'childNodes' collection.
        If Not AttributeName = vbNullString Then
            Set web_Attributes = New Collection
            web_Attributes.Add CreateAttribute(AttributeName, AttributeValue)
        End If
        This.Body.Item("childNodes").Add CreateNode(Name, Value, , web_Attributes)
        
        ' Clear cached converted body
        This.ConvertedBody = Empty
    End If
End Sub

''
' Add `Key-Value` to `Body`.
' `Body` must be a `Dictionary` (if it's an `Array` or `Collection` an error is thrown)
'
' @example
' ```VB.net
' Dim Request As New WebRequest
' Request.Format = WebFormat.Json
'
' Request.AddBodyParameter "a", 123
' Debug.Print Request.Body ' -> "{"a":123}"
'
' ' Can add parameters to existing Dictionary
' Dim Body As New Dictionary
' Body.Add "a", 123
'
' Set Request.Body = Body
' Request.AddBodyParameter "b", 456
'
' Debug.Print Request.Body ' -> "{"a":123,"b":456}"
' ```
'
' @method AddBodyParameter
' @param {Variant} Key
' @param {Variant} Value
' @throws 11020 / 80042b0c / -2147210484 - Cannot add body parameter to non-Dictionary
''
Public Sub AddBodyParameter(Key As Variant, Value As Variant)
    If VBA.IsEmpty(This.Body) Then
        Set This.Body = New Dictionary
    ElseIf Not TypeOf This.Body Is Dictionary Then
        Dim web_ErrorDescription As String
        web_ErrorDescription = "Cannot add body parameter to non-Dictionary Body (existing Body must be of type Dictionary)"

        WebHelpers.LogError web_ErrorDescription, "WebRequest.AddBodyParameter", 11020 + vbObjectError
        Err.Raise 11020 + vbObjectError, "WebRequest.AddBodyParameter", web_ErrorDescription
    End If

    If VBA.IsObject(Value) Then
        Set This.Body.Item(Key) = Value
    Else
        This.Body.Item(Key) = Value
    End If

    ' Clear cached converted body
    This.ConvertedBody = Empty
End Sub

''
' Prepare request for execution
'
' @internal
' @method Prepare
''
Public Sub Prepare()
    ' Add/replace general headers for request
    Me.SetHeader "User-Agent", Me.UserAgent
    Me.SetHeader "Accept", Me.Accept
    Me.SetHeader "Accept-Encoding", Me.AcceptEncoding
    If Me.Method <> VbWebMethod.vbWebHttpGet Or Me.ContentLength > 0 Then
        Me.SetHeader "Content-Type", Me.ContentType
        Me.SetHeader "Content-Length", VBA.CStr(Me.ContentLength)
    End If
End Sub

''
' Clone request
'
' @internal
' @method Clone
' @return {WebRequest}
''
Public Function Clone() As WebRequest
    Set Clone = New WebRequest

    ' Note: Clone underlying for properties with default values
    Clone.Resource = Me.Resource
    Clone.Method = Me.Method
    Clone.UserAgent = Me.UserAgent
    Clone.Accept = This.Accept
    Clone.ContentType = This.ContentType
    Clone.ContentLength = This.ContentLength
    Clone.RequestFormat = Me.RequestFormat
    Clone.ResponseFormat = Me.ResponseFormat
    Clone.CustomRequestFormat = Me.CustomRequestFormat
    Clone.CustomResponseFormat = Me.CustomResponseFormat

    Set Clone.Headers = WebHelpers.CloneCollection(Me.Headers)
    Set Clone.QuerystringParams = WebHelpers.CloneCollection(Me.QuerystringParams)
    Set Clone.UrlSegments = WebHelpers.CloneDictionary(Me.UrlSegments)
    Set Clone.Cookies = WebHelpers.CloneCollection(Me.Cookies)

    If VBA.IsObject(This.Body) Then
        Set Clone.Body = This.Body
    Else
        Clone.Body = This.Body
    End If
End Function

''
' Create WebRequest from options
'
' @method CreateFromOptions
' @param {Dictionary} Options
' @param {Collection} [Options.Headers] Collection of `KeyValue`
' @param {Collection} [Options.Cookies] Collection of `KeyValue`
' @param {Collection} [Options.QuerystringParams] Collection of `KeyValue`
' @param {Dictionary} [Options.UrlSegments]
''
Public Sub CreateFromOptions(Options As Dictionary)
    If Not Options Is Nothing Then
        If Options.Exists("Headers") Then
            Set Me.Headers = Options.Item("Headers")
        End If
        If Options.Exists("Cookies") Then
            Set Me.Cookies = Options.Item("Cookies")
        End If
        If Options.Exists("QuerystringParams") Then
            Set Me.QuerystringParams = Options.Item("QuerystringParams")
        End If
        If Options.Exists("UrlSegments") Then
            Set Me.UrlSegments = Options.Item("UrlSegments")
        End If
    End If
End Sub

' ============================================= '
' Private Functions
' ============================================= '

' Encode cookie name
'
' References:
' - RFC 6265 https://tools.ietf.org/html/rfc6265
Private Function web_EncodeCookieName(web_CookieName As Variant) As String
    Dim web_CookieVal As String
    Dim web_StringLen As Long

    web_CookieVal = VBA.CStr(web_CookieName)
    web_StringLen = VBA.Len(web_CookieVal)

    If web_StringLen > 0 Then
        Dim web_Result() As String
        Dim web_i As Long
        Dim web_CharCode As Integer
        Dim web_Char As String
        ReDim web_Result(web_StringLen)

        ' ALPHA / DIGIT / "!" / "#" / "$" / "&" / "'" / "*" / "+" / "-" / "." / "^" / "_" / "`" / "|" / "~"
        ' Note: "%" is allowed in spec, but is currently excluded due to parsing issues

        ' Loop through string characters
        For web_i = 1 To web_StringLen
            ' Get character and ascii code
            web_Char = VBA.Mid$(web_CookieVal, web_i, 1)
            web_CharCode = VBA.Asc(web_Char)

            Select Case web_CharCode
                Case 65 To 90, 97 To 122
                    ' ALPHA
                    web_Result(web_i) = web_Char
                Case 48 To 57
                    ' DIGIT
                    web_Result(web_i) = web_Char
                Case 33, 35, 36, 38, 39, 42, 43, 45, 46, 94, 95, 96, 124, 126
                    ' "!" / "#" / "$" / "&" / "'" / "*" / "+" / "-" / "." / "^" / "_" / "`" / "|" / "~"
                    web_Result(web_i) = web_Char

                Case 0 To 15
                    web_Result(web_i) = "%0" & VBA.Hex$(web_CharCode)
                Case Else
                    web_Result(web_i) = "%" & VBA.Hex$(web_CharCode)
            End Select
        Next web_i

        web_EncodeCookieName = VBA.Join$(web_Result, vbNullString)
    End If
End Function

Private Sub Class_Initialize()
    ' Set default values
    Me.RequestFormat = VbWebFormat.vbWebFormatJson
    Me.ResponseFormat = VbWebFormat.vbWebFormatJson
    Me.UserAgent = WebUserAgent

    Set Me.Headers = New Collection
    Set Me.QuerystringParams = New Collection
    Set Me.UrlSegments = New Dictionary
    Set Me.Cookies = New Collection
    Me.ContentLength = -1
End Sub

