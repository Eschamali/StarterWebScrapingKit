VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/**
' * JSON - JSON Reader and Writer
' * @description A high-performance zero-copy JSON reader and lightweight JSON writer for VBA (x86/x64 compatible).
' * Uses compact token trees, SAFEARRAY string aliasing, native InStr quote scanning, lazy node wrappers,
' * typed accessors, raw field access, token iteration for huge arrays, and Stringify support for parsed JSON,
' * primitive VBA values, arrays, Collections, Dictionaries, and nested JSON nodes.
' * Designed for fast parsing and traversal without Dictionary, Collection, or per-node object allocation during parse.
' * @author UesleiDev
' * @version 1.0.1
' */

Option Explicit

#If VBA7 Then
    '/** @description Fast memory block copy used to attach and clear SAFEARRAY descriptors. */
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As LongPtr)

    '/** @description Returns the internal address of a dynamic array descriptor. */
    Private Declare PtrSafe Function VarPtrArray Lib "vbe7" Alias "VarPtr" (ByRef Var() As Any) As LongPtr

    '/** @description Performs ordinal UTF-16 string comparison without allocating key substrings. */
    Private Declare PtrSafe Function CompareStringOrdinal Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal cchCount1 As Long, ByVal lpString2 As LongPtr, ByVal cchCount2 As Long, ByVal bIgnoreCase As Long) As Long

    '/** @description Native pointer size used when copying SAFEARRAY descriptor pointers. */
    Private Const PTR_SIZE As LongPtr = 8
#Else
    '/** @description Fast memory block copy used to attach and clear SAFEARRAY descriptors. */
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

    '/** @description Returns the internal address of a dynamic array descriptor. */
    Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (ByRef Var() As Any) As Long

    '/** @description Performs ordinal UTF-16 string comparison without allocating key substrings. */
    Private Declare Function CompareStringOrdinal Lib "kernel32" (ByVal lpString1 As Long, ByVal cchCount1 As Long, ByVal lpString2 As Long, ByVal cchCount2 As Long, ByVal bIgnoreCase As Long) As Long

    '/** @description Native pointer size used when copying SAFEARRAY descriptor pointers. */
    Private Const PTR_SIZE As Long = 4
#End If

'/** @description Default minimum token capacity used for parsed documents. */
Private Const JSON_MIN_TOKEN_CAPACITY As Long = 65536

'/** @description Maximum initial token capacity guessed from text length before growth begins. */
Private Const JSON_MAX_INITIAL_TOKEN_CAPACITY As Long = 3000000

'/** @description Token capacity threshold after which growth changes from 2x to 1.5x. */
Private Const JSON_LARGE_TOKEN_CAPACITY As Long = 1048576

'/** @description Token preallocation heuristic based on source JSON length. */
Private Const JSON_TOKEN_CAPACITY_DIVISOR As Long = 24

'/** @description Number of spaces used by default when pretty-printing JSON. */
Private Const JSON_DEFAULT_INDENT_SIZE As Long = 2

'/** @description Public JsonType return values. */
Private Const JSON_TYPE_OBJECT As String = "object"
Private Const JSON_TYPE_ARRAY As String = "array"
Private Const JSON_TYPE_STRING As String = "string"
Private Const JSON_TYPE_NUMBER As String = "number"
Private Const JSON_TYPE_BOOLEAN As String = "boolean"
Private Const JSON_TYPE_NULL As String = "null"

'/** @description JSON writer literals and separators. */
Private Const JSON_LITERAL_NULL As String = "null"
Private Const JSON_LITERAL_TRUE As String = "true"
Private Const JSON_LITERAL_FALSE As String = "false"
Private Const JSON_EMPTY_OBJECT As String = "{}"
Private Const JSON_EMPTY_ARRAY As String = "[]"
Private Const JSON_OBJECT_OPEN As String = "{"
Private Const JSON_OBJECT_CLOSE As String = "}"
Private Const JSON_ARRAY_OPEN As String = "["
Private Const JSON_ARRAY_CLOSE As String = "]"
Private Const JSON_COMMA As String = ","
Private Const JSON_COLON As String = ":"
Private Const JSON_COLON_PRETTY As String = ": "
Private Const JSON_DATE_FORMAT As String = "yyyy-mm-dd\THH:nn:ss"
Private Const JSON_DECIMAL_SEPARATOR As String = "."
Private Const VBA_DECIMAL_SEPARATOR As String = ","

'/** @description TypeName values used for supported object serialization. */
Private Const VBA_TYPE_JSON As String = "JSON"
Private Const VBA_TYPE_COLLECTION As String = "Collection"
Private Const VBA_TYPE_DICTIONARY As String = "Dictionary"
Private Const VBA_TYPE_SCRIPTING_DICTIONARY As String = "Scripting.Dictionary"

'/** @description UTF-16 character codes used by the zero-copy parser. */
Private Const CH_TAB As Integer = 9
Private Const CH_LF As Integer = 10
Private Const CH_BACKSPACE As Integer = 8
Private Const CH_FORM_FEED As Integer = 12
Private Const CH_CR As Integer = 13
Private Const CH_SPACE As Integer = 32
Private Const CH_QUOTE As Integer = 34
Private Const CH_PLUS As Integer = 43
Private Const CH_COMMA As Integer = 44
Private Const CH_MINUS As Integer = 45
Private Const CH_DOT As Integer = 46
Private Const CH_COLON As Integer = 58
Private Const CH_0 As Integer = 48
Private Const CH_9 As Integer = 57
Private Const CH_E_UPPER As Integer = 69
Private Const CH_BACKSLASH As Integer = 92
Private Const CH_ARRAY_OPEN As Integer = 91
Private Const CH_ARRAY_CLOSE As Integer = 93
Private Const CH_E_LOWER As Integer = 101
Private Const CH_FALSE As Integer = 102
Private Const CH_NULL As Integer = 110
Private Const CH_TRUE As Integer = 116
Private Const CH_OBJECT_OPEN As Integer = 123
Private Const CH_OBJECT_CLOSE As Integer = 125

Private Const JSON_LITERAL_TRUE_LENGTH As Long = 4
Private Const JSON_LITERAL_FALSE_LENGTH As Long = 5
Private Const JSON_LITERAL_NULL_LENGTH As Long = 4

'/** @description SAFEARRAY descriptor values for aliasing UTF-16 string characters. */
Private Const SAFEARRAY_DIMS_1D As Integer = 1
Private Const SAFEARRAY_FEATURES_NONE As Integer = 0
Private Const SAFEARRAY_LOCKS_NONE As Long = 0
Private Const UTF16_CODE_UNIT_BYTES As Long = 2

'/** @description CompareStringOrdinal return value for equal strings. */
Private Const COMPARE_STRING_CASE_SENSITIVE As Long = 0
Private Const COMPARE_STRING_EQUAL As Long = 2

'/**
' * @struct SAFEARRAY1D
' * @brief Minimal one-dimensional SAFEARRAY descriptor used to alias a VBA String as Integer().
' */
Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    #If VBA7 Then
        pvData As LongPtr
    #Else
        pvData As Long
    #End If
    cElements As Long
    lLbound As Long
End Type

'/**
' * @enum JSType
' * @brief Internal JSON token type identifiers.
' */
Private Enum JSType
    jsNone = 0
    jsObject = 1
    jsArray = 2
    jsString = 3
    jsNumber = 4
    jsBool = 5
    jsNull = 6
End Enum

'/**
' * @struct JSToken
' * @brief Compact parsed JSON token storing hierarchy links and key/value slices.
' */
Private Type JSToken
    Type As Integer
    Parent As Long
    NextSibling As Long
    FirstChild As Long
    LastChild As Long
    ChildCount As Long
    KeyStart As Long
    KeyLen As Long
    ValStart As Long
    ValLen As Long
End Type

Private Type JSWriter
    Text As String
    Pos As Long
End Type

Private m_Chars() As Integer
Private m_Text As String
Private m_Length As Long
Private m_Index As Long
Private m_Tokens() As JSToken
Private m_TokenCount As Long
Private m_TokenCap As Long
Private m_RootId As Long
Private m_NodeId As Long
Private m_Document As JSON
Private m_CharAliasActive As Boolean

'/**
' * @method Class_Terminate
' * @brief Releases the active SAFEARRAY alias before the object is destroyed.
' */
Private Sub Class_Terminate()
    ClearCharAlias
End Sub

'/**
' * @method Parse
' * @brief Parses JSON text into a tokenized JSON document.
' * @param Text Source JSON string.
' * @return Parsed JSON document.
' */
Public Function Parse(ByRef Text As String) As JSON
    Dim doc As JSON
    Set doc = New JSON
    doc.LoadText Text
    Set Parse = doc
End Function

'/**
' * @method Stringify
' * @brief Serializes the current JSON document or node to JSON text.
' * @param Pretty Enables formatted output with line breaks and indentation.
' * @param IndentSize Number of spaces per indentation level.
' * @return JSON text.
' */
Public Function Stringify(Optional ByVal Pretty As Boolean = False, Optional ByVal IndentSize As Long = JSON_DEFAULT_INDENT_SIZE) As String
    Stringify = StringifyCurrent(Pretty, Space$(IndentSize))
End Function

'/**
' * @method StringifyWithIndent
' * @brief Serializes the current JSON document or node using a custom indentation string.
' * @param Pretty Enables formatted output with line breaks and indentation.
' * @param IndentText Text used for each indentation level, such as spaces or vbTab.
' * @return JSON text.
' */
Public Function StringifyWithIndent(Optional ByVal Pretty As Boolean = False, Optional ByVal IndentText As String = "  ") As String
    StringifyWithIndent = StringifyCurrent(Pretty, IndentText)
End Function

'/**
' * @method StringifyValue
' * @brief Serializes an external VBA value to JSON text.
' * @param Value VBA value, array, Collection, Dictionary, or JSON node.
' * @param Pretty Enables formatted output with line breaks and indentation.
' * @param IndentSize Number of spaces per indentation level.
' * @return JSON text.
' */
Public Function StringifyValue(ByVal Value As Variant, Optional ByVal Pretty As Boolean = False, Optional ByVal IndentSize As Long = JSON_DEFAULT_INDENT_SIZE) As String
    StringifyValue = StringifyAny(Value, Pretty, Space$(IndentSize), 0)
End Function

'/**
' * @method StringifyValueWithIndent
' * @brief Serializes an external VBA value using a custom indentation string.
' * @param Value VBA value, array, Collection, Dictionary, or JSON node.
' * @param Pretty Enables formatted output with line breaks and indentation.
' * @param IndentText Text used for each indentation level, such as spaces or vbTab.
' * @return JSON text.
' */
Public Function StringifyValueWithIndent(ByVal Value As Variant, Optional ByVal Pretty As Boolean = False, Optional ByVal IndentText As String = "  ") As String
    StringifyValueWithIndent = StringifyAny(Value, Pretty, IndentText, 0)
End Function

'/**
' * @method LoadText
' * @brief Initializes this instance as a root JSON document.
' * @param Text Source JSON string.
' */
Friend Sub LoadText(ByRef Text As String)
    ClearCharAlias

    m_Text = Text
    m_Length = Len(m_Text)
    m_Index = 0
    m_RootId = 0
    m_NodeId = 0
    Set m_Document = Nothing

    Erase m_Chars
    Erase m_Tokens

    If m_Length = 0 Then Exit Sub

    m_TokenCap = m_Length \ JSON_TOKEN_CAPACITY_DIVISOR

    If m_TokenCap < JSON_MIN_TOKEN_CAPACITY Then
        m_TokenCap = JSON_MIN_TOKEN_CAPACITY
    End If

    If m_TokenCap > JSON_MAX_INITIAL_TOKEN_CAPACITY Then
        m_TokenCap = JSON_MAX_INITIAL_TOKEN_CAPACITY
    End If

    m_TokenCount = 0
    ReDim m_Tokens(1 To m_TokenCap)

    Dim SA As SAFEARRAY1D

    With SA
        .cDims = SAFEARRAY_DIMS_1D
        .fFeatures = SAFEARRAY_FEATURES_NONE
        .cbElements = UTF16_CODE_UNIT_BYTES
        .cLocks = SAFEARRAY_LOCKS_NONE
        .pvData = StrPtr(m_Text)
        .cElements = m_Length
        .lLbound = 0
    End With

    #If VBA7 Then
        Dim pSA As LongPtr
    #Else
        Dim pSA As Long
    #End If

    pSA = VarPtr(SA)
    CopyMemory ByVal VarPtrArray(m_Chars), pSA, PTR_SIZE
    m_CharAliasActive = True

    m_RootId = ParseValue(0, 0, 0)
End Sub

'/**
' * @method InitNode
' * @brief Initializes this instance as a lightweight wrapper around a token from another document.
' * @param TokenId Token to wrap.
' * @param Document Root document that owns the token buffer.
' */
Friend Sub InitNode(ByVal TokenId As Long, ByVal Document As JSON)
    m_NodeId = TokenId
    Set m_Document = Document
End Sub

'/**
' * @property Item
' * @brief Gets a child value by object key or array index.
' * @param Key Object key as String or array index as Long.
' * @return Primitive Variant or JSON node wrapper.
' */
Public Property Get Item(ByVal key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Property

    id = doc.ResolveToken(baseId, key)
    If id = 0 Then Exit Property

    If doc.TokenTypeOf(id) = jsObject Or doc.TokenTypeOf(id) = jsArray Then
        Dim node As JSON
        Set node = New JSON
        node.InitNode id, doc
        Set Item = node
    Else
        Item = doc.PrimitiveValue(id)
    End If
End Property

'/**
' * @property Value
' * @brief Gets the current node value. Objects and arrays return the node itself.
' * @return Primitive Variant or current JSON node.
' */
Public Property Get Value() As Variant
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Property

    If doc.TokenTypeOf(baseId) = jsObject Or doc.TokenTypeOf(baseId) = jsArray Then
        Set Value = Me
    Else
        Value = doc.PrimitiveValue(baseId)
    End If
End Property

'/**
' * @property Count
' * @brief Gets the amount of direct children in the current object or array.
' * @return Direct child count.
' */
Public Property Get Count() As Long
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Property

    Count = doc.TokenChildCountOf(baseId)
End Property

'/**
' * @property JsonType
' * @brief Gets the JSON type name of the current node.
' * @return object, array, string, number, boolean, null, or empty string.
' */
Public Property Get JsonType() As String
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Property

    Select Case doc.TokenTypeOf(baseId)
        Case jsObject
            JsonType = JSON_TYPE_OBJECT
        Case jsArray
            JsonType = JSON_TYPE_ARRAY
        Case jsString
            JsonType = JSON_TYPE_STRING
        Case jsNumber
            JsonType = JSON_TYPE_NUMBER
        Case jsBool
            JsonType = JSON_TYPE_BOOLEAN
        Case jsNull
            JsonType = JSON_TYPE_NULL
    End Select
End Property

'/**
' * @property IsObject
' * @brief Returns True when the current node is a JSON object.
' * @return Boolean state.
' */
Public Property Get IsObject() As Boolean
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Property

    IsObject = (doc.TokenTypeOf(baseId) = jsObject)
End Property

'/**
' * @property IsArray
' * @brief Returns True when the current node is a JSON array.
' * @return Boolean state.
' */
Public Property Get IsArray() As Boolean
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Property

    IsArray = (doc.TokenTypeOf(baseId) = jsArray)
End Property

'/**
' * @property IsNull
' * @brief Returns True when the current node is JSON null.
' * @return Boolean state.
' */
Public Property Get IsNull() As Boolean
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Property

    IsNull = (doc.TokenTypeOf(baseId) = jsNull)
End Property

'/**
' * @method Exists
' * @brief Returns True when an object key or array index exists.
' * @param Key Object key as String or array index as Long.
' * @return Boolean state.
' */
Public Function Exists(ByVal key As Variant) As Boolean
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    Exists = (doc.ResolveToken(baseId, key) <> 0)
End Function

'/**
' * @method Node
' * @brief Gets a child object or array by key/index without returning through Variant.
' * @param Key Object key as String or array index as Long.
' * @return JSON node wrapper.
' */
Public Function node(ByVal key As Variant) As JSON
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    id = doc.ResolveToken(baseId, key)
    If id = 0 Then Exit Function

    If doc.TokenTypeOf(id) = jsObject Or doc.TokenTypeOf(id) = jsArray Then
        Dim nodeObj As JSON
        Set nodeObj = New JSON
        nodeObj.InitNode id, doc
        Set node = nodeObj
    End If
End Function

'/**
' * @method NodeAt
' * @brief Gets an array/object child node by zero-based child index.
' * @param Index Zero-based child index.
' * @return JSON node wrapper.
' */
Public Function NodeAt(ByVal Index As Long) As JSON
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    id = doc.FindIndexChild(baseId, Index)
    If id = 0 Then Exit Function

    If doc.TokenTypeOf(id) = jsObject Or doc.TokenTypeOf(id) = jsArray Then
        Dim nodeObj As JSON
        Set nodeObj = New JSON
        nodeObj.InitNode id, doc
        Set NodeAt = nodeObj
    End If
End Function

'/**
' * @method ValueAt
' * @brief Gets any child value by zero-based child index.
' * @param Index Zero-based child index.
' * @return Primitive Variant or JSON node wrapper.
' */
Public Function ValueAt(ByVal Index As Long) As Variant
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    id = doc.FindIndexChild(baseId, Index)
    If id = 0 Then Exit Function

    If doc.TokenTypeOf(id) = jsObject Or doc.TokenTypeOf(id) = jsArray Then
        Dim node As JSON
        Set node = New JSON
        node.InitNode id, doc
        Set ValueAt = node
    Else
        ValueAt = doc.PrimitiveValue(id)
    End If
End Function

'/**
' * @method KeyAt
' * @brief Gets the key of an object child by zero-based child index.
' * @param Index Zero-based child index.
' * @return Object key text.
' */
Public Function KeyAt(ByVal Index As Long) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    id = doc.FindIndexChild(baseId, Index)
    If id = 0 Then Exit Function

    KeyAt = doc.GetRawKey(id)
End Function

'/**
' * @method StringValue
' * @brief Gets a child value as String by key/index, applying basic JSON unescape.
' * @param Key Object key as String or array index as Long.
' * @return String value.
' */
Public Function StringValue(ByVal key As Variant) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    id = doc.ResolveToken(baseId, key)
    If id = 0 Then Exit Function

    StringValue = doc.ValueAsString(id)
End Function

'/**
' * @method NumberValue
' * @brief Gets a child value as Double by key/index.
' * @param Key Object key as String or array index as Long.
' * @return Double value.
' */
Public Function NumberValue(ByVal key As Variant) As Double
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    id = doc.ResolveToken(baseId, key)
    If id = 0 Then Exit Function

    NumberValue = doc.ValueAsDouble(id)
End Function

'/**
' * @method BoolValue
' * @brief Gets a child value as Boolean by key/index.
' * @param Key Object key as String or array index as Long.
' * @return Boolean value.
' */
Public Function BoolValue(ByVal key As Variant) As Boolean
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    id = doc.ResolveToken(baseId, key)
    If id = 0 Then Exit Function

    BoolValue = doc.ValueAsBool(id)
End Function

'/**
' * @method RawStringValue
' * @brief Gets a child value as raw String by key/index without unescaping.
' * @param Key Object key as String or array index as Long.
' * @return Raw string slice.
' */
Public Function RawStringValue(ByVal key As Variant) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    id = doc.ResolveToken(baseId, key)
    If id = 0 Then Exit Function

    RawStringValue = doc.GetRawSlice(id)
End Function

'/**
' * @method StringAt
' * @brief Gets an indexed child value as String, applying basic JSON unescape.
' * @param Index Zero-based child index.
' * @return String value.
' */
Public Function StringAt(ByVal Index As Long) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    id = doc.FindIndexChild(baseId, Index)
    If id = 0 Then Exit Function

    StringAt = doc.ValueAsString(id)
End Function

'/**
' * @method NumberAt
' * @brief Gets an indexed child value as Double.
' * @param Index Zero-based child index.
' * @return Double value.
' */
Public Function NumberAt(ByVal Index As Long) As Double
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    id = doc.FindIndexChild(baseId, Index)
    If id = 0 Then Exit Function

    NumberAt = doc.ValueAsDouble(id)
End Function

'/**
' * @method BoolAt
' * @brief Gets an indexed child value as Boolean.
' * @param Index Zero-based child index.
' * @return Boolean value.
' */
Public Function BoolAt(ByVal Index As Long) As Boolean
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    id = doc.FindIndexChild(baseId, Index)
    If id = 0 Then Exit Function

    BoolAt = doc.ValueAsBool(id)
End Function

'/**
' * @method RawStringAt
' * @brief Gets an indexed child value as raw String without unescaping.
' * @param Index Zero-based child index.
' * @return Raw string slice.
' */
Public Function RawStringAt(ByVal Index As Long) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    id = doc.FindIndexChild(baseId, Index)
    If id = 0 Then Exit Function

    RawStringAt = doc.GetRawSlice(id)
End Function

'/**
' * @method FirstChildToken
' * @brief Gets the first direct child token of the current node.
' * @return First child token id.
' */
Public Function FirstChildToken() As Long
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    FirstChildToken = doc.TokenFirstChildOf(baseId)
End Function

'/**
' * @method LastChildToken
' * @brief Gets the last direct child token of the current node.
' * @return Last child token id.
' */
Public Function LastChildToken() As Long
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function

    LastChildToken = doc.TokenLastChildOf(baseId)
End Function

'/**
' * @method NextToken
' * @brief Gets the next sibling token for a token.
' * @param TokenId Current token id.
' * @return Next sibling token id.
' */
Public Function NextToken(ByVal TokenId As Long) As Long
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function

    NextToken = doc.TokenNextSiblingOf(TokenId)
End Function

'/**
' * @method NodeFromToken
' * @brief Wraps a token as a JSON node when the token is an object or array.
' * @param TokenId Token id.
' * @return JSON node wrapper.
' */
Public Function NodeFromToken(ByVal TokenId As Long) As JSON
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function

    If doc.TokenTypeOf(TokenId) = jsObject Or doc.TokenTypeOf(TokenId) = jsArray Then
        Dim node As JSON
        Set node = New JSON
        node.InitNode TokenId, doc
        Set NodeFromToken = node
    End If
End Function

'/**
' * @method TokenKey
' * @brief Gets the object key associated with a token.
' * @param TokenId Token id.
' * @return Token key text.
' */
Public Function TokenKey(ByVal TokenId As Long) As String
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function

    TokenKey = doc.GetRawKey(TokenId)
End Function

'/**
' * @method TokenValue
' * @brief Gets the token value as Variant.
' * @param TokenId Token id.
' * @return Primitive Variant or JSON node wrapper.
' */
Public Function TokenValue(ByVal TokenId As Long) As Variant
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function

    If doc.TokenTypeOf(TokenId) = jsObject Or doc.TokenTypeOf(TokenId) = jsArray Then
        Dim node As JSON
        Set node = New JSON
        node.InitNode TokenId, doc
        Set TokenValue = node
    Else
        TokenValue = doc.PrimitiveValue(TokenId)
    End If
End Function

'/**
' * @method TokenStringValue
' * @brief Gets the token value as String, applying basic JSON unescape.
' * @param TokenId Token id.
' * @return String value.
' */
Public Function TokenStringValue(ByVal TokenId As Long) As String
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function

    TokenStringValue = doc.ValueAsString(TokenId)
End Function

'/**
' * @method TokenRawStringValue
' * @brief Gets the token value as raw String without unescaping.
' * @param TokenId Token id.
' * @return Raw string slice.
' */
Public Function TokenRawStringValue(ByVal TokenId As Long) As String
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function

    TokenRawStringValue = doc.GetRawSlice(TokenId)
End Function

'/**
' * @method TokenNumberValue
' * @brief Gets the token value as Double.
' * @param TokenId Token id.
' * @return Double value.
' */
Public Function TokenNumberValue(ByVal TokenId As Long) As Double
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function

    TokenNumberValue = doc.ValueAsDouble(TokenId)
End Function

'/**
' * @method TokenBoolValue
' * @brief Gets the token value as Boolean.
' * @param TokenId Token id.
' * @return Boolean value.
' */
Public Function TokenBoolValue(ByVal TokenId As Long) As Boolean
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function

    TokenBoolValue = doc.ValueAsBool(TokenId)
End Function

'/**
' * @method TokenString
' * @brief Gets a field from an object token as String, applying basic JSON unescape.
' * @param TokenId Object token id.
' * @param Key Field key.
' * @return String value.
' */
Public Function TokenString(ByVal TokenId As Long, ByVal key As Variant) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function

    id = doc.ResolveToken(TokenId, key)
    If id = 0 Then Exit Function

    TokenString = doc.ValueAsString(id)
End Function

'/**
' * @method TokenRawString
' * @brief Gets a field from an object token as raw String without unescaping.
' * @param TokenId Object token id.
' * @param Key Field key.
' * @return Raw string slice.
' */
Public Function TokenRawString(ByVal TokenId As Long, ByVal key As Variant) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function

    id = doc.ResolveToken(TokenId, key)
    If id = 0 Then Exit Function

    TokenRawString = doc.GetRawSlice(id)
End Function

'/**
' * @method TokenRawField
' * @brief Gets a raw string field from an object token using a schema-known String key.
' * @param TokenId Object token id.
' * @param Key Field key.
' * @return Raw string slice.
' */
Public Function TokenRawField(ByVal TokenId As Long, ByVal key As String) As String
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function

    id = doc.FindChildFast(TokenId, key)
    If id = 0 Then Exit Function

    TokenRawField = doc.GetRawSlice(id)
End Function

'/**
' * @method TokenNumber
' * @brief Gets a field from an object token as Double.
' * @param TokenId Object token id.
' * @param Key Field key.
' * @return Double value.
' */
Public Function TokenNumber(ByVal TokenId As Long, ByVal key As Variant) As Double
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function

    id = doc.ResolveToken(TokenId, key)
    If id = 0 Then Exit Function

    TokenNumber = doc.ValueAsDouble(id)
End Function

'/**
' * @method TokenBool
' * @brief Gets a field from an object token as Boolean.
' * @param TokenId Object token id.
' * @param Key Field key.
' * @return Boolean value.
' */
Public Function TokenBool(ByVal TokenId As Long, ByVal key As Variant) As Boolean
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function

    id = doc.ResolveToken(TokenId, key)
    If id = 0 Then Exit Function

    TokenBool = doc.ValueAsBool(id)
End Function

'/**
' * @method TokenNode
' * @brief Gets a field from an object token as a JSON node.
' * @param TokenId Object token id.
' * @param Key Field key.
' * @return JSON node wrapper.
' */
Public Function TokenNode(ByVal TokenId As Long, ByVal key As Variant) As JSON
    Dim doc As JSON
    Dim baseId As Long
    Dim id As Long

    ResolveBase doc, baseId

    If baseId = 0 Then Exit Function
    If TokenId < 1 Or TokenId > doc.TokenCount Then Exit Function

    id = doc.ResolveToken(TokenId, key)
    If id = 0 Then Exit Function

    If doc.TokenTypeOf(id) = jsObject Or doc.TokenTypeOf(id) = jsArray Then
        Dim node As JSON
        Set node = New JSON
        node.InitNode id, doc
        Set TokenNode = node
    End If
End Function


'/**
' * @property TokenCount
' * @brief Gets the total parsed token count for internal wrapper calls.
' * @return Total token count.
' */
Friend Property Get TokenCount() As Long
    TokenCount = m_TokenCount
End Property

'/**
' * @method TokenTypeOf
' * @brief Gets the internal type of a token.
' * @param TokenId Token id.
' * @return Internal token type id.
' */
Friend Function TokenTypeOf(ByVal TokenId As Long) As Integer
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    TokenTypeOf = m_Tokens(TokenId).Type
End Function

'/**
' * @method TokenChildCountOf
' * @brief Gets the direct child count of a token.
' * @param TokenId Token id.
' * @return Direct child count.
' */
Friend Function TokenChildCountOf(ByVal TokenId As Long) As Long
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    TokenChildCountOf = m_Tokens(TokenId).ChildCount
End Function

'/**
' * @method TokenFirstChildOf
' * @brief Gets the first child token of a token.
' * @param TokenId Token id.
' * @return First child token id.
' */
Friend Function TokenFirstChildOf(ByVal TokenId As Long) As Long
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    TokenFirstChildOf = m_Tokens(TokenId).FirstChild
End Function

'/**
' * @method TokenLastChildOf
' * @brief Gets the last child token of a token.
' * @param TokenId Token id.
' * @return Last child token id.
' */
Friend Function TokenLastChildOf(ByVal TokenId As Long) As Long
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    TokenLastChildOf = m_Tokens(TokenId).LastChild
End Function

'/**
' * @method TokenNextSiblingOf
' * @brief Gets the next sibling token of a token.
' * @param TokenId Token id.
' * @return Next sibling token id.
' */
Friend Function TokenNextSiblingOf(ByVal TokenId As Long) As Long
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    TokenNextSiblingOf = m_Tokens(TokenId).NextSibling
End Function

'/**
' * @method ClearCharAlias
' * @brief Clears the character SAFEARRAY alias if it is active.
' */
Private Sub ClearCharAlias()
    If Not m_CharAliasActive Then Exit Sub

    #If VBA7 Then
        Dim nullPtr As LongPtr
    #Else
        Dim nullPtr As Long
    #End If

    nullPtr = 0
    CopyMemory ByVal VarPtrArray(m_Chars), nullPtr, PTR_SIZE
    m_CharAliasActive = False
End Sub

'/**
' * @method ResolveBase
' * @brief Resolves the root document and base token for root instances or node wrappers.
' * @param Document Receives the real root document.
' * @param BaseId Receives the active token id.
' */
Private Sub ResolveBase(ByRef Document As JSON, ByRef baseId As Long)
    If m_NodeId > 0 Then
        Set Document = m_Document
        baseId = m_NodeId
    Else
        Set Document = Me
        baseId = m_RootId
    End If
End Sub

'/**
' * @method ResolveToken
' * @brief Resolves a direct child token by object key or array index.
' * @param ParentId Parent token id.
' * @param Key Object key as String or array index as Long.
' * @return Child token id.
' */
Friend Function ResolveToken(ByVal ParentId As Long, ByVal key As Variant) As Long
    If ParentId < 1 Or ParentId > m_TokenCount Then Exit Function

    If VarType(key) = vbString Then
        ResolveToken = FindChildFast(ParentId, CStr(key))
    Else
        ResolveToken = FindIndexChild(ParentId, CLng(key))
    End If
End Function

'/**
' * @method PrimitiveValue
' * @brief Converts a primitive token into a Variant.
' * @param TokenId Primitive token id.
' * @return Primitive Variant value.
' */
Friend Function PrimitiveValue(ByVal TokenId As Long) As Variant
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function

    Select Case m_Tokens(TokenId).Type
        Case jsString
            PrimitiveValue = GetRawValue(TokenId)
        Case jsNumber
            PrimitiveValue = val(GetRawSlice(TokenId))
        Case jsBool
            PrimitiveValue = ValueAsBool(TokenId)
        Case jsNull
            PrimitiveValue = Null
    End Select
End Function

'/**
' * @method ValueAsString
' * @brief Reads a primitive token as String.
' * @param TokenId Token id.
' * @return String value.
' */
Friend Function ValueAsString(ByVal TokenId As Long) As String
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function

    Select Case m_Tokens(TokenId).Type
        Case jsString
            ValueAsString = GetRawValue(TokenId)
        Case jsNumber, jsBool
            ValueAsString = GetRawSlice(TokenId)
        Case jsNull
            ValueAsString = vbNullString
    End Select
End Function

'/**
' * @method ValueAsDouble
' * @brief Reads a number token as Double.
' * @param TokenId Token id.
' * @return Double value.
' */
Friend Function ValueAsDouble(ByVal TokenId As Long) As Double
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function

    If m_Tokens(TokenId).Type = jsNumber Then
        ValueAsDouble = val(GetRawSlice(TokenId))
    End If
End Function

'/**
' * @method ValueAsBool
' * @brief Reads a boolean token as Boolean.
' * @param TokenId Token id.
' * @return Boolean value.
' */
Friend Function ValueAsBool(ByVal TokenId As Long) As Boolean
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function

    If m_Tokens(TokenId).Type = jsBool Then
        ValueAsBool = (m_Chars(m_Tokens(TokenId).ValStart - 1) = CH_TRUE)
    End If
End Function

'/**
' * @method AddToken
' * @brief Adds a new token and grows the token buffer when required.
' * @return Created token id.
' */
Private Function AddToken() As Long
    m_TokenCount = m_TokenCount + 1

    If m_TokenCount > m_TokenCap Then
        If m_TokenCap < JSON_LARGE_TOKEN_CAPACITY Then
            m_TokenCap = m_TokenCap * 2
        Else
            m_TokenCap = m_TokenCap + (m_TokenCap \ 2)
        End If

        ReDim Preserve m_Tokens(1 To m_TokenCap)
    End If

    AddToken = m_TokenCount
End Function

'/**
' * @method ParseValue
' * @brief Parses any JSON value and returns the created token id.
' * @param ParentId Parent token id.
' * @param KeyStart One-based key start position in the original text.
' * @param KeyLen Key length.
' * @return Created token id.
' */
Private Function ParseValue(ByVal ParentId As Long, ByVal KeyStart As Long, ByVal KeyLen As Long) As Long
    SkipWhitespace

    If m_Index >= m_Length Then Exit Function

    Dim id As Long
    id = AddToken()
    ParseValue = id

    With m_Tokens(id)
        .Parent = ParentId
        .KeyStart = KeyStart
        .KeyLen = KeyLen
    End With

    If ParentId > 0 Then
        If m_Tokens(ParentId).FirstChild = 0 Then
            m_Tokens(ParentId).FirstChild = id
        Else
            m_Tokens(m_Tokens(ParentId).LastChild).NextSibling = id
        End If

        m_Tokens(ParentId).LastChild = id
        m_Tokens(ParentId).ChildCount = m_Tokens(ParentId).ChildCount + 1
    End If

    Select Case m_Chars(m_Index)
        Case CH_OBJECT_OPEN
            m_Tokens(id).Type = jsObject
            ParseObject id

        Case CH_ARRAY_OPEN
            m_Tokens(id).Type = jsArray
            ParseArray id

        Case CH_QUOTE
            m_Tokens(id).Type = jsString
            ParseString id

        Case CH_TRUE
            m_Tokens(id).Type = jsBool
            m_Tokens(id).ValStart = m_Index + 1
            m_Tokens(id).ValLen = JSON_LITERAL_TRUE_LENGTH
            m_Index = m_Index + JSON_LITERAL_TRUE_LENGTH

        Case CH_FALSE
            m_Tokens(id).Type = jsBool
            m_Tokens(id).ValStart = m_Index + 1
            m_Tokens(id).ValLen = JSON_LITERAL_FALSE_LENGTH
            m_Index = m_Index + JSON_LITERAL_FALSE_LENGTH

        Case CH_NULL
            m_Tokens(id).Type = jsNull
            m_Tokens(id).ValStart = m_Index + 1
            m_Tokens(id).ValLen = JSON_LITERAL_NULL_LENGTH
            m_Index = m_Index + JSON_LITERAL_NULL_LENGTH

        Case CH_MINUS, CH_0 To CH_9
            m_Tokens(id).Type = jsNumber
            ParseNumber id
    End Select
End Function

'/**
' * @method SkipWhitespace
' * @brief Skips JSON whitespace characters.
' */
Private Sub SkipWhitespace()
    Do While m_Index < m_Length
        Select Case m_Chars(m_Index)
            Case CH_SPACE, CH_TAB, CH_LF, CH_CR
                m_Index = m_Index + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

'/**
' * @method ParseObject
' * @brief Parses a JSON object.
' * @param id Object token id.
' */
Private Sub ParseObject(ByVal id As Long)
    m_Index = m_Index + 1

    Do While m_Index < m_Length
        SkipWhitespace

        If m_Index >= m_Length Then Exit Do

        If m_Chars(m_Index) = CH_OBJECT_CLOSE Then
            m_Index = m_Index + 1
            Exit Sub
        End If

        Dim kStart As Long
        Dim kLen As Long

        If m_Chars(m_Index) = CH_QUOTE Then
            ScanJSONString kStart, kLen
        End If

        SkipWhitespace

        If m_Index >= m_Length Then Exit Do

        If m_Chars(m_Index) = CH_COLON Then
            m_Index = m_Index + 1
        End If

        ParseValue id, kStart, kLen

        SkipWhitespace

        If m_Index >= m_Length Then Exit Do

        If m_Chars(m_Index) = CH_COMMA Then
            m_Index = m_Index + 1
        End If
    Loop
End Sub

'/**
' * @method ParseArray
' * @brief Parses a JSON array.
' * @param id Array token id.
' */
Private Sub ParseArray(ByVal id As Long)
    m_Index = m_Index + 1

    Do While m_Index < m_Length
        SkipWhitespace

        If m_Index >= m_Length Then Exit Do

        If m_Chars(m_Index) = CH_ARRAY_CLOSE Then
            m_Index = m_Index + 1
            Exit Sub
        End If

        ParseValue id, 0, 0

        SkipWhitespace

        If m_Index >= m_Length Then Exit Do

        If m_Chars(m_Index) = CH_COMMA Then
            m_Index = m_Index + 1
        End If
    Loop
End Sub

'/**
' * @method ParseString
' * @brief Parses a JSON string token.
' * @param id String token id.
' */
Private Sub ParseString(ByVal id As Long)
    ScanJSONString m_Tokens(id).ValStart, m_Tokens(id).ValLen
End Sub

'/**
' * @method ScanJSONString
' * @brief Scans a JSON string using native InStr for quote search and m_Chars for escape checks.
' * @param StartPos Receives one-based string content start.
' * @param LengthOut Receives string content length.
' */
Private Sub ScanJSONString(ByRef StartPos As Long, ByRef LengthOut As Long)
    Dim quotePos As Long
    Dim searchFrom As Long
    Dim slashCount As Long
    Dim p As Long
    Dim firstChar As Long

    m_Index = m_Index + 1
    StartPos = m_Index + 1
    searchFrom = StartPos
    firstChar = StartPos - 1

    Do
        quotePos = InStr(searchFrom, m_Text, """")

        If quotePos = 0 Then
            LengthOut = m_Length - StartPos + 1
            m_Index = m_Length
            Exit Sub
        End If

        slashCount = 0
        p = quotePos - 2

        Do While p >= firstChar
            If m_Chars(p) <> CH_BACKSLASH Then Exit Do
            slashCount = slashCount + 1
            p = p - 1
        Loop

        If (slashCount And 1) = 0 Then
            LengthOut = quotePos - StartPos
            m_Index = quotePos
            Exit Sub
        End If

        searchFrom = quotePos + 1
    Loop
End Sub

'/**
' * @method ParseNumber
' * @brief Parses a JSON number as a raw slice.
' * @param id Number token id.
' */
Private Sub ParseNumber(ByVal id As Long)
    m_Tokens(id).ValStart = m_Index + 1

    Do While m_Index < m_Length
        Select Case m_Chars(m_Index)
            Case CH_MINUS, CH_DOT, CH_0 To CH_9, CH_E_UPPER, CH_E_LOWER, CH_PLUS
                m_Index = m_Index + 1
            Case Else
                Exit Do
        End Select
    Loop

    m_Tokens(id).ValLen = m_Index - m_Tokens(id).ValStart + 1
End Sub

'/**
' * @method FindChildFast
' * @brief Finds an object child by key without allocating key strings.
' * @param ParentId Parent token id.
' * @param KeyName Key name.
' * @return Child token id.
' */
Friend Function FindChildFast(ByVal ParentId As Long, ByRef KeyName As String) As Long
    Dim child As Long

    child = m_Tokens(ParentId).FirstChild

    Do While child > 0
        If KeyEquals(child, KeyName) Then
            FindChildFast = child
            Exit Function
        End If

        child = m_Tokens(child).NextSibling
    Loop
End Function

'/**
' * @method KeyEquals
' * @brief Compares a token key with a VBA string using CompareStringOrdinal.
' * @param TokenId Token id.
' * @param KeyName Key name.
' * @return Boolean comparison result.
' */
Private Function KeyEquals(ByVal TokenId As Long, ByRef KeyName As String) As Boolean
    If m_Tokens(TokenId).KeyLen <> Len(KeyName) Then Exit Function
    If m_Tokens(TokenId).KeyLen = 0 Then Exit Function

    #If VBA7 Then
        Dim pJson As LongPtr
        Dim pKey As LongPtr
    #Else
        Dim pJson As Long
        Dim pKey As Long
    #End If

    pJson = StrPtr(m_Text) + ((m_Tokens(TokenId).KeyStart - 1) * UTF16_CODE_UNIT_BYTES)
    pKey = StrPtr(KeyName)

    KeyEquals = (CompareStringOrdinal(pJson, m_Tokens(TokenId).KeyLen, pKey, Len(KeyName), COMPARE_STRING_CASE_SENSITIVE) = COMPARE_STRING_EQUAL)
End Function

'/**
' * @method FindIndexChild
' * @brief Finds a direct child by zero-based child index.
' * @param ParentId Parent token id.
' * @param Index Zero-based child index.
' * @return Child token id.
' */
Friend Function FindIndexChild(ByVal ParentId As Long, ByVal Index As Long) As Long
    If Index < 0 Then Exit Function

    Dim child As Long
    Dim i As Long

    child = m_Tokens(ParentId).FirstChild

    Do While child > 0
        If i = Index Then
            FindIndexChild = child
            Exit Function
        End If

        i = i + 1
        child = m_Tokens(child).NextSibling
    Loop
End Function

'/**
' * @method GetRawKey
' * @brief Gets an object key string from the original JSON text.
' * @param TokenId Token id.
' * @return Key string.
' */
Friend Function GetRawKey(ByVal TokenId As Long) As String
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    If m_Tokens(TokenId).KeyLen <= 0 Then Exit Function

    GetRawKey = Mid$(m_Text, m_Tokens(TokenId).KeyStart, m_Tokens(TokenId).KeyLen)

    If InStr(1, GetRawKey, "\") > 0 Then
        GetRawKey = UnescapeJSONString(GetRawKey)
    End If
End Function

'/**
' * @method GetRawValue
' * @brief Gets a value string from the original JSON text, applying basic JSON unescape.
' * @param TokenId Token id.
' * @return Value string.
' */
Friend Function GetRawValue(ByVal TokenId As Long) As String
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    If m_Tokens(TokenId).ValLen <= 0 Then Exit Function

    GetRawValue = Mid$(m_Text, m_Tokens(TokenId).ValStart, m_Tokens(TokenId).ValLen)

    If InStr(1, GetRawValue, "\") > 0 Then
        GetRawValue = UnescapeJSONString(GetRawValue)
    End If
End Function

'/**
' * @method GetRawSlice
' * @brief Gets a value string slice without unescaping.
' * @param TokenId Token id.
' * @return Raw value string.
' */
Friend Function GetRawSlice(ByVal TokenId As Long) As String
    If TokenId < 1 Or TokenId > m_TokenCount Then Exit Function
    If m_Tokens(TokenId).ValLen <= 0 Then Exit Function

    GetRawSlice = Mid$(m_Text, m_Tokens(TokenId).ValStart, m_Tokens(TokenId).ValLen)
End Function

'/**
' * @method UnescapeJSONString
' * @brief Applies basic JSON string unescaping.
' * @param Text Escaped JSON string content.
' * @return Unescaped string content.
' */
Private Function UnescapeJSONString(ByVal Text As String) As String
    Text = Replace(Text, "\""", """")
    Text = Replace(Text, "\\", "\")
    Text = Replace(Text, "\/", "/")
    Text = Replace(Text, "\b", Chr$(8))
    Text = Replace(Text, "\f", Chr$(12))
    Text = Replace(Text, "\n", vbLf)
    Text = Replace(Text, "\r", vbCr)
    Text = Replace(Text, "\t", vbTab)

    UnescapeJSONString = Text
End Function

'/**
' * @method StringifyCurrent
' * @brief Serializes the current document or node to JSON text.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @return JSON text.
' */
Friend Function StringifyCurrent(ByVal Pretty As Boolean, ByVal IndentText As String) As String
    Dim doc As JSON
    Dim baseId As Long

    ResolveBase doc, baseId

    If baseId = 0 Then
        StringifyCurrent = JSON_LITERAL_NULL
        Exit Function
    End If

    StringifyCurrent = doc.StringifyToken(baseId, Pretty, IndentText, 0)
End Function

'/**
' * @method StringifyToken
' * @brief Serializes a parsed token to JSON text.
' * @param TokenId Token id.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' * @return JSON text.
' */
Friend Function StringifyToken(ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As String
    Dim writer As JSWriter

    If TokenId < 1 Or TokenId > m_TokenCount Then
        StringifyToken = JSON_LITERAL_NULL
        Exit Function
    End If

    InitWriter writer, SizeToken(TokenId, Pretty, IndentText, Depth)
    WriteToken writer, TokenId, Pretty, IndentText, Depth
    StringifyToken = FinishWriter(writer)
End Function

Private Sub InitWriter(ByRef Writer As JSWriter, ByVal Size As Long)
    If Size <= 0 Then
        Writer.Text = vbNullString
    Else
        Writer.Text = String$(Size, vbNullChar)
    End If

    Writer.Pos = 1
End Sub

Private Function FinishWriter(ByRef Writer As JSWriter) As String
    If Writer.Pos <= Len(Writer.Text) Then
        FinishWriter = Left$(Writer.Text, Writer.Pos - 1)
    Else
        FinishWriter = Writer.Text
    End If
End Function

Private Sub WriteText(ByRef Writer As JSWriter, ByVal Text As String)
    Dim n As Long

    n = Len(Text)
    If n = 0 Then Exit Sub

    Mid$(Writer.Text, Writer.Pos, n) = Text
    Writer.Pos = Writer.Pos + n
End Sub

Private Function SizeToken(ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As Long
    If TokenId < 1 Or TokenId > m_TokenCount Then
        SizeToken = Len(JSON_LITERAL_NULL)
        Exit Function
    End If

    Select Case m_Tokens(TokenId).Type
        Case jsObject
            SizeToken = SizeObjectToken(TokenId, Pretty, IndentText, Depth)
        Case jsArray
            SizeToken = SizeArrayToken(TokenId, Pretty, IndentText, Depth)
        Case jsString
            SizeToken = SizeQuotedJSONString(GetRawValue(TokenId))
        Case jsNumber, jsBool
            SizeToken = m_Tokens(TokenId).ValLen
        Case jsNull
            SizeToken = Len(JSON_LITERAL_NULL)
        Case Else
            SizeToken = Len(JSON_LITERAL_NULL)
    End Select
End Function

Private Sub WriteToken(ByRef Writer As JSWriter, ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    Dim text As String

    If TokenId < 1 Or TokenId > m_TokenCount Then
        WriteText Writer, JSON_LITERAL_NULL
        Exit Sub
    End If

    Select Case m_Tokens(TokenId).Type
        Case jsObject
            WriteObjectToken Writer, TokenId, Pretty, IndentText, Depth
        Case jsArray
            WriteArrayToken Writer, TokenId, Pretty, IndentText, Depth
        Case jsString
            text = GetRawValue(TokenId)
            WriteQuotedJSONString Writer, text
        Case jsNumber, jsBool
            text = GetRawSlice(TokenId)
            WriteText Writer, text
        Case jsNull
            WriteText Writer, JSON_LITERAL_NULL
        Case Else
            WriteText Writer, JSON_LITERAL_NULL
    End Select
End Sub

Private Function SizeObjectToken(ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As Long
    Dim child As Long
    Dim first As Boolean

    If m_Tokens(TokenId).ChildCount = 0 Then
        SizeObjectToken = Len(JSON_EMPTY_OBJECT)
        Exit Function
    End If

    SizeObjectToken = Len(JSON_OBJECT_OPEN)
    first = True
    child = m_Tokens(TokenId).FirstChild

    Do While child > 0
        If first Then
            first = False
        Else
            SizeObjectToken = SizeObjectToken + Len(JSON_COMMA)
        End If

        If Pretty Then SizeObjectToken = SizeObjectToken + Len(vbCrLf) + IndentSize(IndentText, Depth + 1)
        SizeObjectToken = SizeObjectToken + SizeQuotedJSONString(GetRawKey(child))

        If Pretty Then
            SizeObjectToken = SizeObjectToken + Len(JSON_COLON_PRETTY)
        Else
            SizeObjectToken = SizeObjectToken + Len(JSON_COLON)
        End If

        SizeObjectToken = SizeObjectToken + SizeToken(child, Pretty, IndentText, Depth + 1)
        child = m_Tokens(child).NextSibling
    Loop

    If Pretty Then SizeObjectToken = SizeObjectToken + Len(vbCrLf) + IndentSize(IndentText, Depth)
    SizeObjectToken = SizeObjectToken + Len(JSON_OBJECT_CLOSE)
End Function

Private Sub WriteObjectToken(ByRef Writer As JSWriter, ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    Dim child As Long
    Dim first As Boolean
    Dim keyText As String

    If m_Tokens(TokenId).ChildCount = 0 Then
        WriteText Writer, JSON_EMPTY_OBJECT
        Exit Sub
    End If

    WriteText Writer, JSON_OBJECT_OPEN
    first = True
    child = m_Tokens(TokenId).FirstChild

    Do While child > 0
        If first Then
            first = False
        Else
            WriteText Writer, JSON_COMMA
        End If

        If Pretty Then
            WriteText Writer, vbCrLf
            WriteIndent Writer, IndentText, Depth + 1
        End If

        keyText = GetRawKey(child)
        WriteQuotedJSONString Writer, keyText

        If Pretty Then
            WriteText Writer, JSON_COLON_PRETTY
        Else
            WriteText Writer, JSON_COLON
        End If

        WriteToken Writer, child, Pretty, IndentText, Depth + 1
        child = m_Tokens(child).NextSibling
    Loop

    If Pretty Then
        WriteText Writer, vbCrLf
        WriteIndent Writer, IndentText, Depth
    End If

    WriteText Writer, JSON_OBJECT_CLOSE
End Sub

Private Function SizeArrayToken(ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As Long
    Dim child As Long
    Dim first As Boolean

    If m_Tokens(TokenId).ChildCount = 0 Then
        SizeArrayToken = Len(JSON_EMPTY_ARRAY)
        Exit Function
    End If

    SizeArrayToken = Len(JSON_ARRAY_OPEN)
    first = True
    child = m_Tokens(TokenId).FirstChild

    Do While child > 0
        If first Then
            first = False
        Else
            SizeArrayToken = SizeArrayToken + Len(JSON_COMMA)
        End If

        If Pretty Then SizeArrayToken = SizeArrayToken + Len(vbCrLf) + IndentSize(IndentText, Depth + 1)
        SizeArrayToken = SizeArrayToken + SizeToken(child, Pretty, IndentText, Depth + 1)
        child = m_Tokens(child).NextSibling
    Loop

    If Pretty Then SizeArrayToken = SizeArrayToken + Len(vbCrLf) + IndentSize(IndentText, Depth)
    SizeArrayToken = SizeArrayToken + Len(JSON_ARRAY_CLOSE)
End Function

Private Sub WriteArrayToken(ByRef Writer As JSWriter, ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    Dim child As Long
    Dim first As Boolean

    If m_Tokens(TokenId).ChildCount = 0 Then
        WriteText Writer, JSON_EMPTY_ARRAY
        Exit Sub
    End If

    WriteText Writer, JSON_ARRAY_OPEN
    first = True
    child = m_Tokens(TokenId).FirstChild

    Do While child > 0
        If first Then
            first = False
        Else
            WriteText Writer, JSON_COMMA
        End If

        If Pretty Then
            WriteText Writer, vbCrLf
            WriteIndent Writer, IndentText, Depth + 1
        End If

        WriteToken Writer, child, Pretty, IndentText, Depth + 1
        child = m_Tokens(child).NextSibling
    Loop

    If Pretty Then
        WriteText Writer, vbCrLf
        WriteIndent Writer, IndentText, Depth
    End If

    WriteText Writer, JSON_ARRAY_CLOSE
End Sub

'/**
' * @method StringifyObjectToken
' * @brief Serializes an object token.
' * @param TokenId Object token id.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' * @return JSON object text.
' */
Private Function StringifyObjectToken(ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As String
    Dim writer As JSWriter

    InitWriter writer, SizeObjectToken(TokenId, Pretty, IndentText, Depth)
    WriteObjectToken writer, TokenId, Pretty, IndentText, Depth
    StringifyObjectToken = FinishWriter(writer)
End Function

'/**
' * @method StringifyArrayToken
' * @brief Serializes an array token.
' * @param TokenId Array token id.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' * @return JSON array text.
' */
Private Function StringifyArrayToken(ByVal TokenId As Long, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As String
    Dim writer As JSWriter

    InitWriter writer, SizeArrayToken(TokenId, Pretty, IndentText, Depth)
    WriteArrayToken writer, TokenId, Pretty, IndentText, Depth
    StringifyArrayToken = FinishWriter(writer)
End Function

'/**
' * @method StringifyAny
' * @brief Serializes a VBA value to JSON text.
' * @param Value VBA value.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' * @return JSON text.
' */
Friend Function StringifyAny(ByVal Value As Variant, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As String
    Dim writer As JSWriter

    InitWriter writer, SizeAny(Value, Pretty, IndentText, Depth)
    WriteAny writer, Value, Pretty, IndentText, Depth
    StringifyAny = FinishWriter(writer)
End Function

Private Function SizeAny(ByVal Value As Variant, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As Long
    Dim valueType As Integer
    Dim text As String

    If VBA.IsObject(Value) Then
        Dim objectValue As Object
        Set objectValue = Value
        SizeAny = SizeObjectValue(objectValue, Pretty, IndentText, Depth)
        Exit Function
    End If

    valueType = VarType(Value)

    If (valueType And vbArray) = vbArray Then
        SizeAny = SizeArrayValue(Value, Pretty, IndentText, Depth)
        Exit Function
    End If

    If valueType = vbNull Or valueType = vbEmpty Then
        SizeAny = Len(JSON_LITERAL_NULL)
        Exit Function
    End If

    Select Case valueType
        Case vbString
            SizeAny = SizeQuotedJSONString(CStr(Value))
        Case vbBoolean
            If CBool(Value) Then
                SizeAny = Len(JSON_LITERAL_TRUE)
            Else
                SizeAny = Len(JSON_LITERAL_FALSE)
            End If
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            text = NormalizeNumber(CStr(Value))
            SizeAny = Len(text)
        Case vbDate
            text = Format$(CDate(Value), JSON_DATE_FORMAT)
            SizeAny = SizeQuotedJSONString(text)
        Case Else
            SizeAny = SizeQuotedJSONString(CStr(Value))
    End Select
End Function

Private Sub WriteAny(ByRef Writer As JSWriter, ByVal Value As Variant, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    Dim valueType As Integer
    Dim text As String

    If VBA.IsObject(Value) Then
        Dim objectValue As Object
        Set objectValue = Value
        WriteObjectValue Writer, objectValue, Pretty, IndentText, Depth
        Exit Sub
    End If

    valueType = VarType(Value)

    If (valueType And vbArray) = vbArray Then
        WriteArrayValue Writer, Value, Pretty, IndentText, Depth
        Exit Sub
    End If

    If valueType = vbNull Or valueType = vbEmpty Then
        WriteText Writer, JSON_LITERAL_NULL
        Exit Sub
    End If

    Select Case valueType
        Case vbString
            text = CStr(Value)
            WriteQuotedJSONString Writer, text
        Case vbBoolean
            If CBool(Value) Then
                WriteText Writer, JSON_LITERAL_TRUE
            Else
                WriteText Writer, JSON_LITERAL_FALSE
            End If
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            text = NormalizeNumber(CStr(Value))
            WriteText Writer, text
        Case vbDate
            text = Format$(CDate(Value), JSON_DATE_FORMAT)
            WriteQuotedJSONString Writer, text
        Case Else
            text = CStr(Value)
            WriteQuotedJSONString Writer, text
    End Select
End Sub

Private Function SizeObjectValue(ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As Long
    If Value Is Nothing Then
        SizeObjectValue = Len(JSON_LITERAL_NULL)
        Exit Function
    End If

    Select Case TypeName(Value)
        Case VBA_TYPE_JSON
            Dim node As JSON
            Set node = Value
            SizeObjectValue = Len(node.StringifyCurrent(Pretty, IndentText))
        Case VBA_TYPE_COLLECTION
            SizeObjectValue = SizeCollection(Value, Pretty, IndentText, Depth)
        Case VBA_TYPE_DICTIONARY, VBA_TYPE_SCRIPTING_DICTIONARY
            SizeObjectValue = SizeDictionary(Value, Pretty, IndentText, Depth)
        Case Else
            SizeObjectValue = Len(JSON_LITERAL_NULL)
    End Select
End Function

Private Sub WriteObjectValue(ByRef Writer As JSWriter, ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    If Value Is Nothing Then
        WriteText Writer, JSON_LITERAL_NULL
        Exit Sub
    End If

    Select Case TypeName(Value)
        Case VBA_TYPE_JSON
            Dim node As JSON
            Dim text As String
            Set node = Value
            text = node.StringifyCurrent(Pretty, IndentText)
            WriteText Writer, text
        Case VBA_TYPE_COLLECTION
            WriteCollection Writer, Value, Pretty, IndentText, Depth
        Case VBA_TYPE_DICTIONARY, VBA_TYPE_SCRIPTING_DICTIONARY
            WriteDictionary Writer, Value, Pretty, IndentText, Depth
        Case Else
            WriteText Writer, JSON_LITERAL_NULL
    End Select
End Sub

Private Function SizeArrayValue(ByVal Value As Variant, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As Long
    Dim i As Long
    Dim first As Boolean

    SizeArrayValue = Len(JSON_ARRAY_OPEN)
    first = True

    For i = LBound(Value) To UBound(Value)
        If first Then
            first = False
        Else
            SizeArrayValue = SizeArrayValue + Len(JSON_COMMA)
        End If

        If Pretty Then SizeArrayValue = SizeArrayValue + Len(vbCrLf) + IndentSize(IndentText, Depth + 1)
        SizeArrayValue = SizeArrayValue + SizeAny(Value(i), Pretty, IndentText, Depth + 1)
    Next i

    If Pretty And Not first Then SizeArrayValue = SizeArrayValue + Len(vbCrLf) + IndentSize(IndentText, Depth)
    SizeArrayValue = SizeArrayValue + Len(JSON_ARRAY_CLOSE)
End Function

Private Sub WriteArrayValue(ByRef Writer As JSWriter, ByVal Value As Variant, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    Dim i As Long
    Dim first As Boolean

    WriteText Writer, JSON_ARRAY_OPEN
    first = True

    For i = LBound(Value) To UBound(Value)
        If first Then
            first = False
        Else
            WriteText Writer, JSON_COMMA
        End If

        If Pretty Then
            WriteText Writer, vbCrLf
            WriteIndent Writer, IndentText, Depth + 1
        End If

        WriteAny Writer, Value(i), Pretty, IndentText, Depth + 1
    Next i

    If Pretty And Not first Then
        WriteText Writer, vbCrLf
        WriteIndent Writer, IndentText, Depth
    End If

    WriteText Writer, JSON_ARRAY_CLOSE
End Sub

Private Function SizeCollection(ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As Long
    Dim i As Long

    If Value.Count = 0 Then
        SizeCollection = Len(JSON_EMPTY_ARRAY)
        Exit Function
    End If

    SizeCollection = Len(JSON_ARRAY_OPEN)

    For i = 1 To Value.Count
        If i > 1 Then SizeCollection = SizeCollection + Len(JSON_COMMA)
        If Pretty Then SizeCollection = SizeCollection + Len(vbCrLf) + IndentSize(IndentText, Depth + 1)
        SizeCollection = SizeCollection + SizeAny(Value.Item(i), Pretty, IndentText, Depth + 1)
    Next i

    If Pretty Then SizeCollection = SizeCollection + Len(vbCrLf) + IndentSize(IndentText, Depth)
    SizeCollection = SizeCollection + Len(JSON_ARRAY_CLOSE)
End Function

Private Sub WriteCollection(ByRef Writer As JSWriter, ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    Dim i As Long

    If Value.Count = 0 Then
        WriteText Writer, JSON_EMPTY_ARRAY
        Exit Sub
    End If

    WriteText Writer, JSON_ARRAY_OPEN

    For i = 1 To Value.Count
        If i > 1 Then WriteText Writer, JSON_COMMA

        If Pretty Then
            WriteText Writer, vbCrLf
            WriteIndent Writer, IndentText, Depth + 1
        End If

        WriteAny Writer, Value.Item(i), Pretty, IndentText, Depth + 1
    Next i

    If Pretty Then
        WriteText Writer, vbCrLf
        WriteIndent Writer, IndentText, Depth
    End If

    WriteText Writer, JSON_ARRAY_CLOSE
End Sub

Private Function SizeDictionary(ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As Long
    Dim keys As Variant
    Dim key As Variant
    Dim i As Long

    If Value.Count = 0 Then
        SizeDictionary = Len(JSON_EMPTY_OBJECT)
        Exit Function
    End If

    keys = Value.keys
    SizeDictionary = Len(JSON_OBJECT_OPEN)

    For i = LBound(keys) To UBound(keys)
        If i > LBound(keys) Then SizeDictionary = SizeDictionary + Len(JSON_COMMA)
        If Pretty Then SizeDictionary = SizeDictionary + Len(vbCrLf) + IndentSize(IndentText, Depth + 1)

        key = keys(i)
        SizeDictionary = SizeDictionary + SizeQuotedJSONString(CStr(key))

        If Pretty Then
            SizeDictionary = SizeDictionary + Len(JSON_COLON_PRETTY)
        Else
            SizeDictionary = SizeDictionary + Len(JSON_COLON)
        End If

        SizeDictionary = SizeDictionary + SizeAny(Value.Item(key), Pretty, IndentText, Depth + 1)
    Next i

    If Pretty Then SizeDictionary = SizeDictionary + Len(vbCrLf) + IndentSize(IndentText, Depth)
    SizeDictionary = SizeDictionary + Len(JSON_OBJECT_CLOSE)
End Function

Private Sub WriteDictionary(ByRef Writer As JSWriter, ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long)
    Dim keys As Variant
    Dim key As Variant
    Dim i As Long
    Dim keyText As String

    If Value.Count = 0 Then
        WriteText Writer, JSON_EMPTY_OBJECT
        Exit Sub
    End If

    keys = Value.keys
    WriteText Writer, JSON_OBJECT_OPEN

    For i = LBound(keys) To UBound(keys)
        If i > LBound(keys) Then WriteText Writer, JSON_COMMA

        If Pretty Then
            WriteText Writer, vbCrLf
            WriteIndent Writer, IndentText, Depth + 1
        End If

        key = keys(i)
        keyText = CStr(key)
        WriteQuotedJSONString Writer, keyText

        If Pretty Then
            WriteText Writer, JSON_COLON_PRETTY
        Else
            WriteText Writer, JSON_COLON
        End If

        WriteAny Writer, Value.Item(key), Pretty, IndentText, Depth + 1
    Next i

    If Pretty Then
        WriteText Writer, vbCrLf
        WriteIndent Writer, IndentText, Depth
    End If

    WriteText Writer, JSON_OBJECT_CLOSE
End Sub

Private Function IndentSize(ByVal IndentText As String, ByVal Depth As Long) As Long
    If Depth <= 0 Then Exit Function
    IndentSize = Len(IndentText) * Depth
End Function

Private Sub WriteIndent(ByRef Writer As JSWriter, ByVal IndentText As String, ByVal Depth As Long)
    Dim text As String

    text = RepeatText(IndentText, Depth)
    WriteText Writer, text
End Sub

Private Function SizeQuotedJSONString(ByVal Text As String) As Long
    Dim i As Long
    Dim n As Long
    Dim ch As Integer

    n = Len(Text)
    SizeQuotedJSONString = 2
    i = 1

    Do While i <= n
        ch = AscW(Mid$(Text, i, 1))

        Select Case ch
            Case CH_QUOTE, CH_BACKSLASH, CH_TAB, CH_BACKSPACE, CH_FORM_FEED, CH_LF
                SizeQuotedJSONString = SizeQuotedJSONString + 2
                i = i + 1
            Case CH_CR
                SizeQuotedJSONString = SizeQuotedJSONString + 2
                If i < n And Mid$(Text, i + 1, 1) = vbLf Then
                    i = i + 2
                Else
                    i = i + 1
                End If
            Case Else
                SizeQuotedJSONString = SizeQuotedJSONString + 1
                i = i + 1
        End Select
    Loop
End Function

Private Sub WriteQuotedJSONString(ByRef Writer As JSWriter, ByVal Text As String)
    Dim i As Long
    Dim n As Long
    Dim startPos As Long
    Dim ch As Integer
    Dim chunk As String

    WriteText Writer, """"
    n = Len(Text)

    If n = 0 Then
        WriteText Writer, """"
        Exit Sub
    End If

    i = 1
    startPos = 1

    Do While i <= n
        ch = AscW(Mid$(Text, i, 1))

        Select Case ch
            Case CH_QUOTE, CH_BACKSLASH, CH_TAB, CH_BACKSPACE, CH_FORM_FEED, CH_LF, CH_CR
                If i > startPos Then
                    chunk = Mid$(Text, startPos, i - startPos)
                    WriteText Writer, chunk
                End If

                Select Case ch
                    Case CH_QUOTE
                        WriteText Writer, "\"""
                        i = i + 1
                    Case CH_BACKSLASH
                        WriteText Writer, "\\"
                        i = i + 1
                    Case CH_TAB
                        WriteText Writer, "\t"
                        i = i + 1
                    Case CH_BACKSPACE
                        WriteText Writer, "\b"
                        i = i + 1
                    Case CH_FORM_FEED
                        WriteText Writer, "\f"
                        i = i + 1
                    Case CH_LF
                        WriteText Writer, "\n"
                        i = i + 1
                    Case CH_CR
                        If i < n And Mid$(Text, i + 1, 1) = vbLf Then
                            WriteText Writer, "\n"
                            i = i + 2
                        Else
                            WriteText Writer, "\r"
                            i = i + 1
                        End If
                End Select

                startPos = i
            Case Else
                i = i + 1
        End Select
    Loop

    If startPos <= n Then
        chunk = Mid$(Text, startPos, n - startPos + 1)
        WriteText Writer, chunk
    End If

    WriteText Writer, """"
End Sub

'/**
' * @method StringifyObjectValue
' * @brief Serializes supported VBA object values to JSON text.
' * @param Value Object value.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' * @return JSON text.
' */
Private Function StringifyObjectValue(ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As String
    Dim writer As JSWriter

    InitWriter writer, SizeObjectValue(Value, Pretty, IndentText, Depth)
    WriteObjectValue writer, Value, Pretty, IndentText, Depth
    StringifyObjectValue = FinishWriter(writer)
End Function

'/**
' * @method StringifyArrayValue
' * @brief Serializes a one-dimensional VBA array to JSON text.
' * @param Value VBA array.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' * @return JSON array text.
' */
Private Function StringifyArrayValue(ByVal Value As Variant, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As String
    Dim writer As JSWriter

    InitWriter writer, SizeArrayValue(Value, Pretty, IndentText, Depth)
    WriteArrayValue writer, Value, Pretty, IndentText, Depth
    StringifyArrayValue = FinishWriter(writer)
End Function

'/**
' * @method StringifyCollection
' * @brief Serializes a VBA Collection as a JSON array.
' * @param Value Collection object.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' * @return JSON array text.
' */
Private Function StringifyCollection(ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As String
    Dim writer As JSWriter

    InitWriter writer, SizeCollection(Value, Pretty, IndentText, Depth)
    WriteCollection writer, Value, Pretty, IndentText, Depth
    StringifyCollection = FinishWriter(writer)
End Function

'/**
' * @method StringifyDictionary
' * @brief Serializes a Scripting.Dictionary as a JSON object.
' * @param Value Dictionary object.
' * @param Pretty Enables formatted output.
' * @param IndentText Text used for each indentation level.
' * @param Depth Current nesting depth.
' * @return JSON object text.
' */
Private Function StringifyDictionary(ByVal Value As Object, ByVal Pretty As Boolean, ByVal IndentText As String, ByVal Depth As Long) As String
    Dim writer As JSWriter

    InitWriter writer, SizeDictionary(Value, Pretty, IndentText, Depth)
    WriteDictionary writer, Value, Pretty, IndentText, Depth
    StringifyDictionary = FinishWriter(writer)
End Function

'/**
' * @method RepeatText
' * @brief Repeats an indentation string a fixed number of times.
' * @param Text Text to repeat.
' * @param Count Number of repetitions.
' * @return Repeated text.
' */
Private Function RepeatText(ByVal Text As String, ByVal Count As Long) As String
    Dim i As Long
    Dim n As Long
    Dim pos As Long

    If Count <= 0 Then Exit Function
    If Len(Text) = 1 Then
        RepeatText = String$(Count, Text)
        Exit Function
    End If

    n = Len(Text)
    RepeatText = String$(n * Count, vbNullChar)
    pos = 1

    For i = 1 To Count
        Mid$(RepeatText, pos, n) = Text
        pos = pos + n
    Next i
End Function

'/**
' * @method QuoteJSONString
' * @brief Escapes and quotes a string as JSON string text.
' * @param Text Raw string text.
' * @return Quoted JSON string.
' */
Private Function QuoteJSONString(ByVal Text As String) As String
    Dim writer As JSWriter

    InitWriter writer, SizeQuotedJSONString(Text)
    WriteQuotedJSONString writer, Text
    QuoteJSONString = FinishWriter(writer)
End Function

'/**
' * @method NormalizeNumber
' * @brief Normalizes a VBA numeric string to JSON decimal format.
' * @param Text Numeric text.
' * @return JSON-compatible number text.
' */
Private Function NormalizeNumber(ByVal Text As String) As String
    NormalizeNumber = Replace(Text, VBA_DECIMAL_SEPARATOR, JSON_DECIMAL_SEPARATOR)
End Function
