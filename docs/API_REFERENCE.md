# JSON API Reference

> Current target: **JSON v1.0.1**. **JSON.cls** is a high-performance, zero-copy JSON reader and lightweight JSON writer for VBA Office hosts.

It supports x86 and x64 VBA, compact token-tree parsing, lazy node wrappers, typed accessors, raw field access, token iteration for large arrays and objects, and `Stringify` support for parsed JSON, primitive VBA values, arrays, Collections, Dictionaries, and nested JSON nodes.

This reference documents the current public API exposed by **JSON.cls**.

## Table of Contents

- [Core Concepts](#core-concepts)
- [Mental Model](#mental-model)
- [Document and Node Model](#document-and-node-model)
- [Token Tree](#token-tree)
- [Zero-Copy Parsing](#zero-copy-parsing)
- [Lazy Node Wrappers](#lazy-node-wrappers)
- [Typed Accessors](#typed-accessors)
- [Raw Access](#raw-access)
- [Token Iteration](#token-iteration)
- [Stringify Pipeline](#stringify-pipeline)
- [Compatibility Strategy](#compatibility-strategy)
- [Parsing](#parsing)
- [Serialization](#serialization)
- [Core Node Properties](#core-node-properties)
- [Child Lookup](#child-lookup)
- [Indexed Access](#indexed-access)
- [Typed Field Access](#typed-field-access)
- [Typed Indexed Access](#typed-indexed-access)
- [Token Traversal](#token-traversal)
- [Token Value Access](#token-value-access)
- [Token Field Access](#token-field-access)
- [Practical Recipes](#practical-recipes)
- [Best Practices](#best-practices)
- [Troubleshooting](#troubleshooting)
- [Complete Public API Index](#complete-public-api-index)

## Core Concepts

### Mental Model

JSON has three main layers:

1. **Document**: created with `JSON.Parse(text)`. The document owns the original JSON text and the internal token buffer.
2. **Nodes**: lightweight wrappers around tokens. Objects and arrays can be accessed as `JSON` objects without copying their contents.
3. **Tokens**: internal compact records that store JSON type, parent/child/sibling links, key slices, value slices, and child counts.

The key distinction is:

```txt
Document = parsed JSON root and source text owner
Node     = lightweight wrapper around one token
Token    = internal parsed JSON entry
```

A parsed JSON document does not eagerly allocate Dictionaries, Collections, or one object per node during parsing. It builds a compact token tree and creates node wrappers only when requested.

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""user"":{""name"":""Ueslei""},""score"":100}")

Debug.Print doc.StringValue("score")

Dim user As JSON
Set user = doc.Node("user")

Debug.Print user.StringValue("name")
```

### Document and Node Model

The root result returned by `JSON.Parse` is a document object. Calling `Node`, `NodeAt`, `TokenNode`, or `NodeFromToken` creates a lightweight wrapper pointing back to the original document.

This means child nodes are cheap to create, but they depend on the root document staying alive.

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""items"":[1,2,3]}")

Dim items As JSON
Set items = doc.Node("items")

Debug.Print items.Count
```

Keep the root document variable alive for as long as you use any child node wrappers.

### Token Tree

Internally, JSON stores parsed data as tokens. A token stores:

- JSON type.
- Parent token.
- First child token.
- Last child token.
- Next sibling token.
- Child count.
- Key slice position.
- Value slice position.

This enables fast traversal without materializing the full JSON into nested VBA containers.

```txt
object
 ├─ "name"  -> string
 ├─ "score" -> number
 └─ "tags"  -> array
              ├─ string
              └─ string
```

### Zero-Copy Parsing

The parser keeps slices into the original source text. Keys, strings, numbers, booleans, and raw fields are not copied during parsing.

Values are converted only when requested:

```vb
Debug.Print doc.StringValue("name")
Debug.Print doc.NumberValue("score")
Debug.Print doc.BoolValue("active")
```

For best performance, parse once and reuse the document.

### Lazy Node Wrappers

Objects and arrays are wrapped lazily.

```vb
Dim profile As JSON
Set profile = doc.Node("profile")
```

No wrapper is created for every parsed object or array automatically. This avoids large object-allocation overhead when parsing big API responses.

### Typed Accessors

Typed accessors are direct helpers for known schemas.

```vb
Debug.Print doc.StringValue("name")
Debug.Print doc.NumberValue("id")
Debug.Print doc.BoolValue("verified")
```

They are useful when you know the expected field type and want compact user code.

### Raw Access

Raw access returns the original JSON slice without additional traversal or conversion.

```vb
Dim rawProfile As String
rawProfile = users.TokenRawField(t, "profile")
```

This is useful for forwarding nested JSON, caching fragments, or delaying expensive traversal.

### Token Iteration

Token iteration is the fastest style for large arrays.

```vb
Dim rows As JSON
Set rows = doc.Node("rows")

Dim t As Long
t = rows.FirstChildToken()

Do While t <> 0
    Debug.Print rows.TokenString(t, "name"), rows.TokenNumber(t, "score")
    t = rows.NextToken(t)
Loop
```

This avoids creating a `JSON` wrapper for every array element.

### Stringify Pipeline

JSON can serialize:

- Parsed JSON documents.
- Parsed JSON nodes.
- Strings.
- Booleans.
- Numbers.
- Dates.
- Null and Empty values.
- One-dimensional arrays.
- Collections.
- Scripting.Dictionary objects.
- Nested JSON objects.

```vb
Debug.Print doc.Stringify()
Debug.Print doc.Stringify(True)
Debug.Print JSON.StringifyValue(dict, True)
```

### Compatibility Strategy

JSON is distributed as a single predeclared `.cls` class named `JSON`.

The intended usage style is:

```vb
Dim doc As JSON
Set doc = JSON.Parse(jsonText)

Debug.Print doc.StringValue("name")
Debug.Print JSON.StringifyValue(value, True)
```

The class is compatible with 32-bit and 64-bit Office through conditional compilation.

## Parsing

### Parse

```vb
Public Function Parse(ByRef Text As String) As JSON
```

Parses JSON text into a tokenized `JSON` document.

Returns a new `JSON` document instance.

```vb
Public Sub ParseExample()
    Dim text As String
    text = "{""name"":""Ueslei"",""age"":18,""active"":true}"

    Dim doc As JSON
    Set doc = JSON.Parse(text)

    Debug.Print doc.StringValue("name")
    Debug.Print doc.NumberValue("age")
    Debug.Print doc.BoolValue("active")
End Sub
```

The parser is optimized for well-formed JSON. It is intended for fast parsing and traversal of trusted or already validated JSON payloads.

## Serialization

### Stringify

```vb
Public Function Stringify( _
    Optional ByVal Pretty As Boolean = False, _
    Optional ByVal IndentSize As Long = 2 _
) As String
```

Serializes the current JSON document or node to JSON text.

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""name"":""JSON"",""language"":""VBA""}")

Debug.Print doc.Stringify()
```

Pretty output:

```vb
Debug.Print doc.Stringify(True)
```

Custom space count:

```vb
Debug.Print doc.Stringify(True, 4)
```

### StringifyWithIndent

```vb
Public Function StringifyWithIndent( _
    Optional ByVal Pretty As Boolean = False, _
    Optional ByVal IndentText As String = "  " _
) As String
```

Serializes the current JSON document or node using a custom indentation string.

```vb
Debug.Print doc.StringifyWithIndent(True, vbTab)
```

Use this when you want tabs or a custom formatting style.

### StringifyValue

```vb
Public Function StringifyValue( _
    ByVal Value As Variant, _
    Optional ByVal Pretty As Boolean = False, _
    Optional ByVal IndentSize As Long = 2 _
) As String
```

Serializes an external VBA value to JSON text.

Supported values include:

| VBA Value | JSON Output |
| `String` | JSON string |
| `Boolean` | `true` / `false` |
| Numeric types | JSON number |
| `Null` | `null` |
| `Empty` | `null` |
| `Date` | ISO-like JSON string |
| One-dimensional array | JSON array |
| `Collection` | JSON array |
| `Dictionary` / `Scripting.Dictionary` | JSON object |
| `JSON` object/node | Serialized JSON |

```vb
Public Sub StringifyDictionaryExample()
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")

    data("name") = "JSON"
    data("language") = "VBA"
    data("fast") = True

    Debug.Print JSON.StringifyValue(data, True)
End Sub
```

### StringifyValueWithIndent

```vb
Public Function StringifyValueWithIndent( _
    ByVal Value As Variant, _
    Optional ByVal Pretty As Boolean = False, _
    Optional ByVal IndentText As String = "  " _
) As String
```

Serializes an external VBA value using custom indentation text.

```vb
Debug.Print JSON.StringifyValueWithIndent(data, True, vbTab)
```

## Core Node Properties

### Item

```vb
Public Property Get Item(ByVal key As Variant) As Variant
```

Gets a child value by object key or array index.

`Item` is the default member, so these are equivalent:

```vb
Debug.Print doc.Item("name")
Debug.Print doc("name")
```

Default-member access can be chained through nested objects and arrays. This is an important usage style for JSON because it lets code read known payload shapes without repeatedly calling `Node`, `ValueAt`, `StringAt`, or other accessor functions.

```vb
Dim myJson As JSON
Set myJson = JSON.Parse("{""names"":[""Ana"",""Bia"",""Caio""]}")

Debug.Print myJson("names")(0)
Debug.Print myJson("names")(1)
```

This style is not always highlighted in VBA JSON documentation, but it is intentional in this class. Use it when the JSON shape is known and concise traversal is more useful than explicit typed accessors.

For primitive values, it returns a `Variant`.

For objects and arrays, it returns a `JSON` node wrapper.

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""user"":{""name"":""Ueslei""},""score"":100}")

Debug.Print doc("score")

Dim user As JSON
Set user = doc("user")

Debug.Print user("name")
```

### Value

```vb
Public Property Get Value() As Variant
```

Gets the current node value.

For primitive nodes, returns the primitive `Variant`.

For objects and arrays, returns the current `JSON` node.

```vb
Dim item As Variant
item = doc.Item("score")

Debug.Print item
```

### Count

```vb
Public Property Get Count() As Long
```

Returns the amount of direct children in the current object or array.

```vb
Dim arr As JSON
Set arr = JSON.Parse("[10,20,30]")

Debug.Print arr.Count
```

For primitive nodes, the count is `0`.

### JsonType

```vb
Public Property Get JsonType() As String
```

Returns the JSON type name of the current node.

Possible values:

| Value | Meaning |
| `object` | JSON object |
| `array` | JSON array |
| `string` | JSON string |
| `number` | JSON number |
| `boolean` | JSON boolean |
| `null` | JSON null |
| empty string | Empty or invalid wrapper |

```vb
Debug.Print doc.JsonType
```

### IsObject

```vb
Public Property Get IsObject() As Boolean
```

Returns `True` when the current node is a JSON object.

```vb
If doc.IsObject Then
    Debug.Print "Root is object"
End If
```

### IsArray

```vb
Public Property Get IsArray() As Boolean
```

Returns `True` when the current node is a JSON array.

```vb
If items.IsArray Then
    Debug.Print items.Count
End If
```

### IsNull

```vb
Public Property Get IsNull() As Boolean
```

Returns `True` when the current node is JSON `null`.

```vb
If doc.Node("meta").IsNull Then
    Debug.Print "Meta is null"
End If
```

## Child Lookup

### Exists

```vb
Public Function Exists(ByVal key As Variant) As Boolean
```

Returns `True` when an object key or array index exists.

```vb
If doc.Exists("data") Then
    Debug.Print "data exists"
End If
```

For arrays:

```vb
If arr.Exists(0) Then
    Debug.Print arr.ValueAt(0)
End If
```

### Node

```vb
Public Function Node(ByVal key As Variant) As JSON
```

Gets a child object or array by object key or array index.

Returns `Nothing` if the child does not exist or is not an object/array.

```vb
Dim user As JSON
Set user = doc.Node("user")

If Not user Is Nothing Then
    Debug.Print user.StringValue("name")
End If
```

For array indexes:

```vb
Dim first As JSON
Set first = users.Node(0)
```

Use `Node` when you know the target is an object or array.

## Indexed Access

### NodeAt

```vb
Public Function NodeAt(ByVal Index As Long) As JSON
```

Gets an object or array child by zero-based child index.

Returns `Nothing` if the child does not exist or is not an object/array.

```vb
Dim firstUser As JSON
Set firstUser = users.NodeAt(0)
```

### ValueAt

```vb
Public Function ValueAt(ByVal Index As Long) As Variant
```

Gets any child value by zero-based child index.

Primitive values are returned as `Variant`.

Objects and arrays are returned as `JSON` node wrappers.

```vb
Dim arr As JSON
Set arr = JSON.Parse("[10,20,30]")

Debug.Print arr.ValueAt(0)
Debug.Print arr.ValueAt(1)
Debug.Print arr.ValueAt(2)
```

### KeyAt

```vb
Public Function KeyAt(ByVal Index As Long) As String
```

Gets the key of an object child by zero-based child index.

```vb
Dim i As Long

For i = 0 To doc.Count - 1
    Debug.Print doc.KeyAt(i), doc.ValueAt(i)
Next
```

For arrays, keys are empty because array items do not have object keys.

## Typed Field Access

### StringValue

```vb
Public Function StringValue(ByVal key As Variant) As String
```

Gets a child value as `String` by object key or array index.

For JSON strings, it returns the string value with basic JSON escapes decoded.

For numbers and booleans, it returns the raw value text.

For null, it returns an empty string.

```vb
Debug.Print doc.StringValue("name")
```

### NumberValue

```vb
Public Function NumberValue(ByVal key As Variant) As Double
```

Gets a child value as `Double` by object key or array index.

Returns `0` when the field is missing or not a number.

```vb
Debug.Print doc.NumberValue("score")
```

### BoolValue

```vb
Public Function BoolValue(ByVal key As Variant) As Boolean
```

Gets a child value as `Boolean` by object key or array index.

Returns `False` when the field is missing or not a boolean.

```vb
If doc.BoolValue("active") Then
    Debug.Print "Active"
End If
```

### RawStringValue

```vb
Public Function RawStringValue(ByVal key As Variant) As String
```

Gets a child string value without unescaping.

This is useful when you want the exact raw string slice stored in the JSON text.

```vb
Debug.Print doc.RawStringValue("message")
```

## Typed Indexed Access

### StringAt

```vb
Public Function StringAt(ByVal Index As Long) As String
```

Gets an indexed child value as `String`.

```vb
Debug.Print arr.StringAt(0)
```

### NumberAt

```vb
Public Function NumberAt(ByVal Index As Long) As Double
```

Gets an indexed child value as `Double`.

```vb
Debug.Print arr.NumberAt(0)
```

### BoolAt

```vb
Public Function BoolAt(ByVal Index As Long) As Boolean
```

Gets an indexed child value as `Boolean`.

```vb
Debug.Print arr.BoolAt(0)
```

### RawStringAt

```vb
Public Function RawStringAt(ByVal Index As Long) As String
```

Gets an indexed string value without unescaping.

```vb
Debug.Print arr.RawStringAt(0)
```

## Token Traversal

Token traversal is intended for high-performance loops over large arrays or objects.

### FirstChildToken

```vb
Public Function FirstChildToken() As Long
```

Returns the first direct child token of the current node.

Returns `0` when there is no child.

```vb
Dim t As Long
t = arr.FirstChildToken()
```

### LastChildToken

```vb
Public Function LastChildToken() As Long
```

Returns the last direct child token of the current node.

```vb
Debug.Print arr.LastChildToken()
```

### NextToken

```vb
Public Function NextToken(ByVal TokenId As Long) As Long
```

Returns the next sibling token for a token.

Returns `0` when there is no next sibling.

```vb
Dim t As Long
t = arr.FirstChildToken()

Do While t <> 0
    Debug.Print arr.TokenValue(t)
    t = arr.NextToken(t)
Loop
```

### NodeFromToken

```vb
Public Function NodeFromToken(ByVal TokenId As Long) As JSON
```

Wraps a token as a `JSON` node when the token is an object or array.

Returns `Nothing` for primitive tokens.

```vb
Dim row As JSON
Set row = rows.NodeFromToken(t)
```

Use this when you need object-style access for a specific token.

## Token Value Access

### TokenKey

```vb
Public Function TokenKey(ByVal TokenId As Long) As String
```

Gets the object key associated with a token.

```vb
Dim t As Long
t = doc.FirstChildToken()

Do While t <> 0
    Debug.Print doc.TokenKey(t), doc.TokenValue(t)
    t = doc.NextToken(t)
Loop
```

### TokenValue

```vb
Public Function TokenValue(ByVal TokenId As Long) As Variant
```

Gets a token value as `Variant`.

Primitive tokens return primitive values.

Object and array tokens return `JSON` node wrappers.

```vb
Debug.Print arr.TokenValue(t)
```

### TokenStringValue

```vb
Public Function TokenStringValue(ByVal TokenId As Long) As String
```

Gets a token value as `String`, applying basic JSON unescape for strings.

```vb
Debug.Print arr.TokenStringValue(t)
```

### TokenRawStringValue

```vb
Public Function TokenRawStringValue(ByVal TokenId As Long) As String
```

Gets a token value as a raw string without unescaping.

```vb
Debug.Print arr.TokenRawStringValue(t)
```

### TokenNumberValue

```vb
Public Function TokenNumberValue(ByVal TokenId As Long) As Double
```

Gets a token value as `Double`.

```vb
Debug.Print arr.TokenNumberValue(t)
```

### TokenBoolValue

```vb
Public Function TokenBoolValue(ByVal TokenId As Long) As Boolean
```

Gets a token value as `Boolean`.

```vb
Debug.Print arr.TokenBoolValue(t)
```

## Token Field Access

Token field helpers are designed for arrays of objects.

Given a JSON payload like this:

```json
{
  "users": [
    { "name": "Ana", "score": 10, "active": true },
    { "name": "Bia", "score": 20, "active": false }
  ]
}
```

You can iterate without creating a node wrapper for every user:

```vb
Dim doc As JSON
Set doc = JSON.Parse(responseText)

Dim users As JSON
Set users = doc.Node("users")

Dim t As Long
t = users.FirstChildToken()

Do While t <> 0
    Debug.Print users.TokenString(t, "name")
    Debug.Print users.TokenNumber(t, "score")
    Debug.Print users.TokenBool(t, "active")

    t = users.NextToken(t)
Loop
```

### TokenString

```vb
Public Function TokenString(ByVal TokenId As Long, ByVal key As Variant) As String
```

Gets a field from an object token as `String`.

```vb
Debug.Print users.TokenString(t, "name")
```

### TokenRawString

```vb
Public Function TokenRawString(ByVal TokenId As Long, ByVal key As Variant) As String
```

Gets a field from an object token as raw string without unescaping.

```vb
Debug.Print users.TokenRawString(t, "name")
```

### TokenRawField

```vb
Public Function TokenRawField(ByVal TokenId As Long, ByVal key As String) As String
```

Gets a raw field slice from an object token using a schema-known string key.

```vb
Dim rawProfile As String
rawProfile = users.TokenRawField(t, "profile")
```

This is useful for nested objects and arrays:

```json
{
  "profile": {
    "level": 12,
    "rank": "S"
  }
}
```

`TokenRawField(t, "profile")` returns the raw JSON text of the nested object.

### TokenNumber

```vb
Public Function TokenNumber(ByVal TokenId As Long, ByVal key As Variant) As Double
```

Gets a field from an object token as `Double`.

```vb
Debug.Print users.TokenNumber(t, "score")
```

### TokenBool

```vb
Public Function TokenBool(ByVal TokenId As Long, ByVal key As Variant) As Boolean
```

Gets a field from an object token as `Boolean`.

```vb
Debug.Print users.TokenBool(t, "active")
```

### TokenNode

```vb
Public Function TokenNode(ByVal TokenId As Long, ByVal key As Variant) As JSON
```

Gets a field from an object token as a `JSON` node.

Returns `Nothing` if the field does not exist or is not an object/array.

```vb
Dim profile As JSON
Set profile = users.TokenNode(t, "profile")

If Not profile Is Nothing Then
    Debug.Print profile.StringValue("rank")
End If
```

## Practical Recipes

### Parse an API Response

```vb
Public Sub ReadApiResponse(ByVal responseText As String)
    Dim doc As JSON
    Set doc = JSON.Parse(responseText)

    If Not doc.Exists("data") Then Exit Sub

    Dim data As JSON
    Set data = doc.Node("data")

    Debug.Print data.StringValue("name")
    Debug.Print data.NumberValue("id")
    Debug.Print data.BoolValue("active")
End Sub
```

### Read a Nested Object

```vb
Public Sub ReadNestedObject()
    Dim text As String
    text = "{""user"":{""name"":""Ueslei"",""role"":""developer""}}"

    Dim doc As JSON
    Set doc = JSON.Parse(text)

    Dim user As JSON
    Set user = doc.Node("user")

    If Not user Is Nothing Then
        Debug.Print user.StringValue("name")
        Debug.Print user.StringValue("role")
    End If
End Sub
```

### Read a Simple Array

```vb
Public Sub ReadSimpleArray()
    Dim arr As JSON
    Set arr = JSON.Parse("[""VBA"",""Rust"",""JavaScript""]")

    Dim i As Long
    For i = 0 To arr.Count - 1
        Debug.Print arr.StringAt(i)
    Next i
End Sub
```

### Read an Array of Objects

```vb
Public Sub ReadArrayOfObjects(ByVal responseText As String)
    Dim doc As JSON
    Set doc = JSON.Parse(responseText)

    Dim users As JSON
    Set users = doc.Node("users")

    If users Is Nothing Then Exit Sub

    Dim i As Long
    Dim user As JSON

    For i = 0 To users.Count - 1
        Set user = users.NodeAt(i)

        If Not user Is Nothing Then
            Debug.Print user.StringValue("name")
            Debug.Print user.NumberValue("score")
        End If
    Next i
End Sub
```

### Fast Array-of-Objects Iteration

```vb
Public Sub ReadLargeArrayFast(ByVal responseText As String)
    Dim doc As JSON
    Set doc = JSON.Parse(responseText)

    Dim rows As JSON
    Set rows = doc.Node("rows")

    If rows Is Nothing Then Exit Sub

    Dim t As Long
    t = rows.FirstChildToken()

    Do While t <> 0
        Debug.Print rows.TokenString(t, "name"), rows.TokenNumber(t, "score")
        t = rows.NextToken(t)
    Loop
End Sub
```

### Extract a Raw Nested Field

```vb
Public Sub ExtractRawPayload(ByVal responseText As String)
    Dim doc As JSON
    Set doc = JSON.Parse(responseText)

    Dim rows As JSON
    Set rows = doc.Node("rows")

    If rows Is Nothing Then Exit Sub

    Dim t As Long
    t = rows.FirstChildToken()

    Do While t <> 0
        Debug.Print rows.TokenRawField(t, "payload")
        t = rows.NextToken(t)
    Loop
End Sub
```

### Build JSON from a Dictionary

```vb
Public Sub BuildJsonObject()
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")

    data("name") = "JSON"
    data("version") = "1.0.1"
    data("language") = "VBA"
    data("fast") = True

    Debug.Print JSON.StringifyValue(data, True)
End Sub
```

### Build JSON from a Collection

```vb
Public Sub BuildJsonArray()
    Dim list As Collection
    Set list = New Collection

    list.Add "Excel"
    list.Add "PowerPoint"
    list.Add "Access"

    Debug.Print JSON.StringifyValue(list, True)
End Sub
```

### Build JSON from a VBA Array

```vb
Public Sub BuildJsonFromArray()
    Dim values(0 To 2) As Variant

    values(0) = "VBA"
    values(1) = "JSON"
    values(2) = True

    Debug.Print JSON.StringifyValue(values, True)
End Sub
```

### Pretty Print Existing JSON

```vb
Public Sub PrettyPrintJson(ByVal text As String)
    Dim doc As JSON
    Set doc = JSON.Parse(text)

    Debug.Print doc.Stringify(True)
End Sub
```

### Use Tab Indentation

```vb
Public Sub PrettyPrintWithTabs(ByVal text As String)
    Dim doc As JSON
    Set doc = JSON.Parse(text)

    Debug.Print doc.StringifyWithIndent(True, vbTab)
End Sub
```

## Best Practices

### Keep the Root Document Alive

Node wrappers point back to the document that owns the token buffer.

Good:

```vb
Dim doc As JSON
Set doc = JSON.Parse(text)

Dim data As JSON
Set data = doc.Node("data")

Debug.Print data.Count
```

Avoid returning only a child node from a short-lived local document unless you also keep the root alive.

### Use Typed Accessors for Known Schemas

When you know the expected JSON structure, typed accessors keep code short and avoid repeated `Variant` handling.

```vb
Debug.Print doc.StringValue("name")
Debug.Print doc.NumberValue("id")
Debug.Print doc.BoolValue("active")
```

### Use Token Iteration for Large Arrays

For huge arrays, prefer token traversal.

```vb
Dim t As Long
t = rows.FirstChildToken()

Do While t <> 0
    Debug.Print rows.TokenString(t, "name")
    t = rows.NextToken(t)
Loop
```

### Use Raw Access for Forwarding Data

If you only need to forward a nested object or array, use raw field access.

```vb
Dim raw As String
raw = rows.TokenRawField(t, "payload")
```

### Prefer StringifyValue for External VBA Values

Use `Stringify` for parsed JSON documents and nodes.

Use `StringifyValue` for regular VBA values.

```vb
Debug.Print doc.Stringify(True)
Debug.Print JSON.StringifyValue(dict, True)
```

### Use Dictionaries for JSON Objects

`Scripting.Dictionary` maps naturally to a JSON object.

```vb
Dim obj As Object
Set obj = CreateObject("Scripting.Dictionary")

obj("name") = "JSON"
obj("ok") = True

Debug.Print JSON.StringifyValue(obj)
```

### Use Collections or Arrays for JSON Arrays

Collections are convenient when the length is dynamic.

```vb
Dim arr As Collection
Set arr = New Collection

arr.Add "a"
arr.Add "b"

Debug.Print JSON.StringifyValue(arr)
```

## Notes and Limitations

### Parser Validation

The parser is optimized for speed and low allocation. It assumes normal, well-formed JSON input.

For untrusted input, validate or sanitize upstream if strict error reporting is required.

### String Escapes

String reading applies basic JSON unescape behavior for:

- `\"`
- `\\`
- `\/`
- `\b`
- `\f`
- `\n`
- `\r`
- `\t`

Unicode escape sequences such as `\uXXXX` are not expanded into Unicode characters by the current lightweight unescape helper.

### Number Parsing

Numbers are converted through VBA numeric conversion behavior.

```vb
Debug.Print doc.NumberValue("price")
```

For exact decimal/financial handling, keep raw numeric text if needed.

### Missing Values

Most accessors return the default VBA value when a field is missing or has an unexpected type.

| Accessor | Missing Result |
| `StringValue` | `""` |
| `NumberValue` | `0` |
| `BoolValue` | `False` |
| `Node` | `Nothing` |
| `ValueAt` | Empty Variant |
| `Token...` helpers | Default VBA value |

Use `Exists` when you need to distinguish missing fields from default values.

```vb
If doc.Exists("score") Then
    Debug.Print doc.NumberValue("score")
End If
```

### Object Key Comparison

Object key lookup is case-sensitive and uses ordinal comparison.

```vb
doc.Exists("name")
doc.Exists("Name")
```

These are different keys.

## Troubleshooting

### `Node` Returns Nothing

`Node` only returns objects and arrays.

This returns `Nothing` if `"name"` is a string:

```vb
Set value = doc.Node("name")
```

Use `StringValue` instead:

```vb
Debug.Print doc.StringValue("name")
```

### `NumberValue` Returns 0

Possible causes:

- The field is missing.
- The field is not a number.
- The number text cannot be converted as expected by VBA.

Check existence first:

```vb
If doc.Exists("score") Then
    Debug.Print doc.NumberValue("score")
End If
```

### `BoolValue` Returns False

`False` can mean either the JSON value is actually `false`, the field is missing, or the field is not a boolean.

Use `Exists` if needed:

```vb
If doc.Exists("active") Then
    Debug.Print doc.BoolValue("active")
End If
```

### Large Arrays Feel Slow

Avoid this pattern for huge arrays:

```vb
For i = 0 To rows.Count - 1
    Set row = rows.NodeAt(i)
    Debug.Print row.StringValue("name")
Next
```

Prefer token iteration:

```vb
Dim t As Long
t = rows.FirstChildToken()

Do While t <> 0
    Debug.Print rows.TokenString(t, "name")
    t = rows.NextToken(t)
Loop
```

### Dictionary Output Requires Scripting.Dictionary Object

You can use late binding:

```vb
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
```

No explicit reference is required when using late binding.

### Pretty Output Is Bigger and Slower

Pretty output adds line breaks and indentation.

Use compact output for storage, network payloads, or performance-sensitive paths:

```vb
Debug.Print doc.Stringify(False)
```

Use pretty output for debugging:

```vb
Debug.Print doc.Stringify(True)
```

## Complete Public API Index

### Parsing

| API | Signature | Description |
| `Parse` | `Parse(ByRef Text As String) As JSON` | Parses JSON text into a tokenized document. |

### Serialization

| API | Signature | Description |
| `Stringify` | `Stringify(Optional Pretty As Boolean = False, Optional IndentSize As Long = 2) As String` | Serializes the current document or node. |
| `StringifyWithIndent` | `StringifyWithIndent(Optional Pretty As Boolean = False, Optional IndentText As String = "  ") As String` | Serializes the current document or node with custom indentation. |
| `StringifyValue` | `StringifyValue(Value As Variant, Optional Pretty As Boolean = False, Optional IndentSize As Long = 2) As String` | Serializes an external VBA value. |
| `StringifyValueWithIndent` | `StringifyValueWithIndent(Value As Variant, Optional Pretty As Boolean = False, Optional IndentText As String = "  ") As String` | Serializes an external VBA value with custom indentation. |

### Core Properties

| API | Signature | Description |
| `Item` | `Item(key As Variant) As Variant` | Default member. Gets a child value by key or index. |
| `Value` | `Value() As Variant` | Gets the current node value. |
| `Count` | `Count() As Long` | Gets the direct child count. |
| `JsonType` | `JsonType() As String` | Gets the current JSON type name. |
| `IsObject` | `IsObject() As Boolean` | Returns whether the node is an object. |
| `IsArray` | `IsArray() As Boolean` | Returns whether the node is an array. |
| `IsNull` | `IsNull() As Boolean` | Returns whether the node is null. |

### Child Lookup

| API | Signature | Description |
| `Exists` | `Exists(key As Variant) As Boolean` | Checks whether a field or array index exists. |
| `Node` | `Node(key As Variant) As JSON` | Gets a child object or array as a node. |

### Indexed Access

| API | Signature | Description |
| `NodeAt` | `NodeAt(Index As Long) As JSON` | Gets an object/array child by zero-based child index. |
| `ValueAt` | `ValueAt(Index As Long) As Variant` | Gets any child value by zero-based child index. |
| `KeyAt` | `KeyAt(Index As Long) As String` | Gets an object child key by zero-based child index. |

### Typed Field Access

| API | Signature | Description |
| `StringValue` | `StringValue(key As Variant) As String` | Gets a child value as string. |
| `NumberValue` | `NumberValue(key As Variant) As Double` | Gets a child value as double. |
| `BoolValue` | `BoolValue(key As Variant) As Boolean` | Gets a child value as boolean. |
| `RawStringValue` | `RawStringValue(key As Variant) As String` | Gets a child string without unescaping. |

### Typed Indexed Access

| API | Signature | Description |
| `StringAt` | `StringAt(Index As Long) As String` | Gets an indexed child value as string. |
| `NumberAt` | `NumberAt(Index As Long) As Double` | Gets an indexed child value as double. |
| `BoolAt` | `BoolAt(Index As Long) As Boolean` | Gets an indexed child value as boolean. |
| `RawStringAt` | `RawStringAt(Index As Long) As String` | Gets an indexed child string without unescaping. |

### Token Traversal

| API | Signature | Description |
| `FirstChildToken` | `FirstChildToken() As Long` | Gets the first direct child token. |
| `LastChildToken` | `LastChildToken() As Long` | Gets the last direct child token. |
| `NextToken` | `NextToken(TokenId As Long) As Long` | Gets the next sibling token. |
| `NodeFromToken` | `NodeFromToken(TokenId As Long) As JSON` | Wraps an object/array token as a node. |

### Token Value Access

| API | Signature | Description |
| `TokenKey` | `TokenKey(TokenId As Long) As String` | Gets the key associated with a token. |
| `TokenValue` | `TokenValue(TokenId As Long) As Variant` | Gets a token value as Variant. |
| `TokenStringValue` | `TokenStringValue(TokenId As Long) As String` | Gets a token value as string. |
| `TokenRawStringValue` | `TokenRawStringValue(TokenId As Long) As String` | Gets a token string without unescaping. |
| `TokenNumberValue` | `TokenNumberValue(TokenId As Long) As Double` | Gets a token value as double. |
| `TokenBoolValue` | `TokenBoolValue(TokenId As Long) As Boolean` | Gets a token value as boolean. |

### Token Field Access

| API | Signature | Description |
| `TokenString` | `TokenString(TokenId As Long, key As Variant) As String` | Gets a field from an object token as string. |
| `TokenRawString` | `TokenRawString(TokenId As Long, key As Variant) As String` | Gets a field from an object token as raw string. |
| `TokenRawField` | `TokenRawField(TokenId As Long, key As String) As String` | Gets a raw field slice from an object token. |
| `TokenNumber` | `TokenNumber(TokenId As Long, key As Variant) As Double` | Gets a field from an object token as double. |
| `TokenBool` | `TokenBool(TokenId As Long, key As Variant) As Boolean` | Gets a field from an object token as boolean. |
| `TokenNode` | `TokenNode(TokenId As Long, key As Variant) As JSON` | Gets a field from an object token as a JSON node. |

## License

MIT. Designed for fast JSON parsing, clean traversal, low allocation, and practical data automation inside Microsoft Office.
