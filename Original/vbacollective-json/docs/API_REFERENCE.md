# JSON API Reference

**JSON** is a single-file JSON parser and writer for VBA. It is designed for Office projects that need fast parsing, low allocation, typed access, lightweight traversal, token iteration, and practical JSON serialization without requiring external references.

This reference documents the current public API exposed by **JSON.cls**, including the newer typed helpers such as `StringKey`, `NumberKey`, `BoolKey`, `NodeKey`, `StringIndex`, `NodeIndex`, `ExistsKey`, `ExistsIndex`, and `Keys`.

## Table of Contents

- [Overview](#overview)
- [Core Concepts](#core-concepts)
- [Recommended Usage](#recommended-usage)
- [Parsing](#parsing)
- [Serialization](#serialization)
- [Core Node Properties](#core-node-properties)
- [Generic Child Access](#generic-child-access)
- [Keys and Existence](#keys-and-existence)
- [Typed Object Access](#typed-object-access)
- [Typed Array Access](#typed-array-access)
- [Legacy Typed Accessors](#legacy-typed-accessors)
- [Token Traversal](#token-traversal)
- [Token Value Access](#token-value-access)
- [Token Field Access](#token-field-access)
- [Practical Recipes](#practical-recipes)
- [Validation and Tests](#validation-and-tests)
- [Best Practices](#best-practices)
- [Notes and Limitations](#notes-and-limitations)
- [Troubleshooting](#troubleshooting)
- [Complete Public API Index](#complete-public-api-index)
- [License](#license)

## Overview

`JSON.cls` provides a compact JSON document model for VBA.

Main goals:

- Parse JSON text into a compact token tree.
- Avoid building nested `Scripting.Dictionary` or `Collection` trees during parse.
- Keep node wrappers lazy and lightweight.
- Provide typed accessors for common paths.
- Keep compatibility with older generic access patterns.
- Support `Stringify` for parsed documents and normal VBA values.
- Work in 32-bit and 64-bit Office.

The intended basic usage is:

```vb
Dim doc As JSON
Set doc = JSON.Parse(jsonText)

If doc.ExistsKey("name") Then
    Debug.Print doc.StringKey("name")
End If
```

For arrays:

```vb
Dim arr As JSON
Set arr = JSON.Parse("[10,20,30]")

If arr.ExistsIndex(0) Then
    Debug.Print arr.NumberIndex(0)
End If
```

For nested nodes:

```vb
Dim user As JSON
Set user = doc.NodeKey("user")

If Not user Is Nothing Then
    Debug.Print user.StringKey("role")
End If
```

## Core Concepts

### Mental Model

The class has three important layers:

1. **Document**: the root `JSON` object returned by `JSON.Parse(text)`.
2. **Node**: a lightweight `JSON` wrapper around an object or array token.
3. **Token**: an internal parsed entry representing an object, array, string, number, boolean, or null.

```txt
Document = owns parsed source and token buffer
Node     = wrapper around one token inside the document
Token    = internal parsed JSON entry
```

The parser does not eagerly materialize every object into dictionaries. It builds a compact token tree and creates node wrappers only when code asks for them.

### Document and Node Lifetime

Child nodes depend on the root document staying alive.

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""items"":[1,2,3]}")

Dim items As JSON
Set items = doc.NodeKey("items")

Debug.Print items.Count
```

Keep `doc` alive while using `items`, child nodes, or token helpers.

### Token Tree

Internally, each parsed value is stored as a token containing information such as:

- JSON type.
- Parent token.
- First child token.
- Last child token.
- Next sibling token.
- Child count.
- Key slice.
- Value slice.

Example tree:

```txt
object
 ├─ "name"  -> string
 ├─ "score" -> number
 └─ "tags"  -> array
              ├─ string
              └─ string
```

This model allows fast traversal without allocating a full object graph.

### Lazy Node Wrappers

Objects and arrays are wrapped only when requested:

```vb
Dim profile As JSON
Set profile = doc.NodeKey("profile")
```

This avoids creating one VBA object for every JSON object or array during parsing.

### Typed Accessors

The preferred modern API uses explicit object-key and array-index helpers.

Object access:

```vb
Debug.Print doc.StringKey("name")
Debug.Print doc.NumberKey("id")
Debug.Print doc.BoolKey("active")
```

Array access:

```vb
Debug.Print arr.StringIndex(0)
Debug.Print arr.NumberIndex(1)
Debug.Print arr.BoolIndex(2)
```

These helpers are cleaner and avoid relying on `Variant` for the common read path.

### Generic Access Still Exists

The generic/default API is still available for compatibility:

```vb
Debug.Print doc("name")
Debug.Print doc.Item("name")
Debug.Print doc.ValueAt(0)
```

Use it when you want dynamic access. Use the typed API when the schema is known.

## Recommended Usage

### Recommended Object Style

```vb
Public Sub ReadObject()
    Dim doc As JSON
    Set doc = JSON.Parse("{""name"":""Ueslei"",""age"":18,""active"":true}")

    If doc.ExistsKey("name") Then
        Debug.Print doc.StringKey("name")
    End If

    Debug.Print doc.NumberKey("age")
    Debug.Print doc.BoolKey("active")
End Sub
```

### Recommended Array Style

```vb
Public Sub ReadArray()
    Dim arr As JSON
    Set arr = JSON.Parse("[""Excel"",""PowerPoint"",""Access""]")

    Dim i As Long
    For i = 0 To arr.Count - 1
        Debug.Print arr.StringIndex(i)
    Next i
End Sub
```

### Recommended Nested Style

```vb
Public Sub ReadNested()
    Dim doc As JSON
    Set doc = JSON.Parse("{""user"":{""name"":""Ueslei"",""role"":""developer""}}")

    Dim user As JSON
    Set user = doc.NodeKey("user")

    If Not user Is Nothing Then
        Debug.Print user.StringKey("name")
        Debug.Print user.StringKey("role")
    End If
End Sub
```

### Recommended Large Array Style

For very large arrays, prefer token iteration:

```vb
Public Sub ReadRows(ByVal responseText As String)
    Dim doc As JSON
    Set doc = JSON.Parse(responseText)

    Dim rows As JSON
    Set rows = doc.NodeKey("rows")

    If rows Is Nothing Then Exit Sub

    Dim t As Long
    t = rows.FirstChildToken()

    Do While t <> 0
        Debug.Print rows.TokenString(t, "name")
        Debug.Print rows.TokenNumber(t, "score")
        Debug.Print rows.TokenBool(t, "active")

        t = rows.NextToken(t)
    Loop
End Sub
```

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

    Debug.Print doc.StringKey("name")
    Debug.Print doc.NumberKey("age")
    Debug.Print doc.BoolKey("active")
End Sub
```

The current preferred return pattern is strongly typed:

```vb
Dim doc As JSON
Set doc = JSON.Parse(jsonText)
```

This avoids treating the parse result as a generic `Variant`.

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
Debug.Print doc.Stringify(True)
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

Use this when you want tabs or a custom indentation string.

### StringifyValue

```vb
Public Function StringifyValue( _
    ByVal Value As Variant, _
    Optional ByVal Pretty As Boolean = False, _
    Optional ByVal IndentSize As Long = 2 _
) As String
```

Serializes a normal VBA value to JSON text.

Supported values include:

| VBA Value | JSON Output |
|:--|:--|
| `String` | JSON string |
| `Boolean` | `true` / `false` |
| Numeric types | JSON number |
| `Null` | `null` |
| `Empty` | `null` |
| `Date` | JSON string |
| One-dimensional array | JSON array |
| `Collection` | JSON array |
| `Dictionary` / `Scripting.Dictionary` | JSON object |
| `JSON` object/node | Serialized JSON |

Example:

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

Serializes a normal VBA value using a custom indentation string.

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

Default member chaining is supported:

```vb
Dim myJson As JSON
Set myJson = JSON.Parse("{""names"":[""Ana"",""Bia"",""Caio""]}")

Debug.Print myJson("names")(0)
Debug.Print myJson("names")(1)
```

For primitive values, `Item` returns a `Variant`.

For objects and arrays, `Item` returns a `JSON` node wrapper.

```vb
Dim user As JSON
Set user = doc("user")

Debug.Print user("name")
```

Use `Item` when convenience matters. Prefer `StringKey`, `NumberKey`, `BoolKey`, `NodeKey`, `StringIndex`, `NumberIndex`, `BoolIndex`, and `NodeIndex` when you know the expected schema.

### Value

```vb
Public Property Get Value() As Variant
```

Gets the current node value.

For primitive nodes, returns a primitive `Variant`.

For object and array nodes, returns the current `JSON` node.

```vb
Debug.Print doc.NodeKey("user").Value
```

### Count

```vb
Public Property Get Count() As Long
```

Returns the number of direct children in the current object or array.

```vb
Dim arr As JSON
Set arr = JSON.Parse("[10,20,30]")

Debug.Print arr.Count
```

Primitive nodes return `0`.

### JsonType

```vb
Public Property Get JsonType() As String
```

Returns the JSON type name of the current node.

Possible values:

| Value | Meaning |
|:--|:--|
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
If arr.IsArray Then
    Debug.Print arr.Count
End If
```

### IsNull

```vb
Public Property Get IsNull() As Boolean
```

Returns `True` when the current node is JSON `null`.

```vb
Dim meta As JSON
Set meta = doc.NodeKey("meta")

If Not meta Is Nothing Then
    If meta.IsNull Then Debug.Print "Meta is null"
End If
```

## Generic Child Access

The generic access API exists for compatibility and dynamic scenarios.

### Node

```vb
Public Function Node(ByVal key As Variant) As JSON
```

Gets a child object or array by object key or array index.

Returns `Nothing` if the child does not exist or is not an object/array.

```vb
Dim user As JSON
Set user = doc.Node("user")
```

Modern equivalent:

```vb
Set user = doc.NodeKey("user")
```

For arrays:

```vb
Set first = arr.Node(0)
```

Modern equivalent:

```vb
Set first = arr.NodeIndex(0)
```

### NodeAt

```vb
Public Function NodeAt(ByVal Index As Long) As JSON
```

Gets an object or array child by zero-based child position.

Returns `Nothing` if the child does not exist or is not an object/array.

```vb
Dim firstUser As JSON
Set firstUser = users.NodeAt(0)
```

Modern equivalent:

```vb
Set firstUser = users.NodeIndex(0)
```

### ValueAt

```vb
Public Function ValueAt(ByVal Index As Long) As Variant
```

Gets any child value by zero-based child position.

Primitive values are returned as `Variant`.

Objects and arrays are returned as `JSON` node wrappers.

```vb
Debug.Print arr.ValueAt(0)
Debug.Print arr.ValueAt(1)
Debug.Print arr.ValueAt(2)
```

### KeyAt

```vb
Public Function KeyAt(ByVal Index As Long) As String
```

Gets the key of an object child by zero-based child position.

```vb
Dim i As Long

For i = 0 To doc.Count - 1
    Debug.Print doc.KeyAt(i), doc.ValueAt(i)
Next i
```

For arrays, `KeyAt` returns an empty string because array values do not have object keys.

## Keys and Existence

### Keys

```vb
Public Function Keys() As Variant
```

Returns the direct keys of the current node.

For objects, it returns a zero-based `Variant` array containing property names.

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""name"":""Ueslei"",""age"":18,""active"":true}")

Dim keys As Variant
keys = doc.Keys

Debug.Print keys(0) ' name
Debug.Print keys(1) ' age
Debug.Print keys(2) ' active
```

For arrays, it returns a zero-based `Variant` array containing numeric indexes.

```vb
Dim arr As JSON
Set arr = JSON.Parse("[""a"",""b"",""c""]")

Dim keys As Variant
keys = arr.Keys

Debug.Print keys(0) ' 0
Debug.Print keys(1) ' 1
Debug.Print keys(2) ' 2
```

For primitives or empty nodes, it returns an empty array.

### Exists

```vb
Public Function Exists(ByVal key As Variant) As Boolean
```

Compatibility helper that checks whether an object key or array index exists.

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

Prefer `ExistsKey` and `ExistsIndex` when the schema is known.

### ExistsKey

```vb
Public Function ExistsKey(ByRef key As String) As Boolean
```

Checks whether an object field exists.

```vb
If doc.ExistsKey("name") Then
    Debug.Print doc.StringKey("name")
End If
```

Use this for object lookup instead of generic `Exists` when you know the input is a key.

### ExistsIndex

```vb
Public Function ExistsIndex(ByVal Index As Long) As Boolean
```

Checks whether an array index exists.

```vb
If arr.ExistsIndex(2) Then
    Debug.Print arr.StringIndex(2)
End If
```

Use this for array lookup instead of generic `Exists` when you know the input is an index.

## Typed Object Access

Typed object accessors are the recommended API for known object schemas.

### StringKey

```vb
Public Function StringKey(ByRef key As String) As String
```

Gets an object field as `String`.

For JSON strings, it returns the decoded string value.

For numbers and booleans, it returns the raw value text.

For missing fields or `null`, it returns an empty string.

```vb
Debug.Print doc.StringKey("name")
```

### NumberKey

```vb
Public Function NumberKey(ByRef key As String) As Double
```

Gets an object field as `Double`.

Returns `0` when the field is missing or not a number.

```vb
Debug.Print doc.NumberKey("score")
```

### BoolKey

```vb
Public Function BoolKey(ByRef key As String) As Boolean
```

Gets an object field as `Boolean`.

Returns `False` when the field is missing or not a boolean.

```vb
If doc.BoolKey("active") Then
    Debug.Print "Active"
End If
```

### RawStringKey

```vb
Public Function RawStringKey(ByRef key As String) As String
```

Gets an object string field without unescaping.

```vb
Debug.Print doc.RawStringKey("message")
```

This is useful when you want the exact raw text inside the JSON string value.

### NodeKey

```vb
Public Function NodeKey(ByRef key As String) As JSON
```

Gets an object field as a `JSON` node.

Returns `Nothing` when the field is missing or is not an object/array.

```vb
Dim user As JSON
Set user = doc.NodeKey("user")

If Not user Is Nothing Then
    Debug.Print user.StringKey("name")
End If
```

Use `NodeKey` for nested objects and arrays.

## Typed Array Access

Typed array accessors are the recommended API for known array schemas.

### StringIndex

```vb
Public Function StringIndex(ByVal Index As Long) As String
```

Gets an array item as `String`.

```vb
Debug.Print arr.StringIndex(0)
```

### NumberIndex

```vb
Public Function NumberIndex(ByVal Index As Long) As Double
```

Gets an array item as `Double`.

```vb
Debug.Print arr.NumberIndex(0)
```

### BoolIndex

```vb
Public Function BoolIndex(ByVal Index As Long) As Boolean
```

Gets an array item as `Boolean`.

```vb
Debug.Print arr.BoolIndex(0)
```

### RawStringIndex

```vb
Public Function RawStringIndex(ByVal Index As Long) As String
```

Gets an array string item without unescaping.

```vb
Debug.Print arr.RawStringIndex(0)
```

### NodeIndex

```vb
Public Function NodeIndex(ByVal Index As Long) As JSON
```

Gets an array item as a `JSON` node.

Returns `Nothing` when the item is missing or is not an object/array.

```vb
Dim first As JSON
Set first = users.NodeIndex(0)

If Not first Is Nothing Then
    Debug.Print first.StringKey("name")
End If
```

## Legacy Typed Accessors

The older typed accessor names remain useful for compatibility and mixed key/index access.

### StringValue

```vb
Public Function StringValue(ByVal key As Variant) As String
```

Gets a child value as `String` by object key or array index.

```vb
Debug.Print doc.StringValue("name")
Debug.Print arr.StringValue(0)
```

Modern object equivalent:

```vb
Debug.Print doc.StringKey("name")
```

Modern array equivalent:

```vb
Debug.Print arr.StringIndex(0)
```

### NumberValue

```vb
Public Function NumberValue(ByVal key As Variant) As Double
```

Gets a child value as `Double` by object key or array index.

```vb
Debug.Print doc.NumberValue("score")
Debug.Print arr.NumberValue(0)
```

Modern equivalents:

```vb
Debug.Print doc.NumberKey("score")
Debug.Print arr.NumberIndex(0)
```

### BoolValue

```vb
Public Function BoolValue(ByVal key As Variant) As Boolean
```

Gets a child value as `Boolean` by object key or array index.

```vb
Debug.Print doc.BoolValue("active")
Debug.Print arr.BoolValue(0)
```

Modern equivalents:

```vb
Debug.Print doc.BoolKey("active")
Debug.Print arr.BoolIndex(0)
```

### RawStringValue

```vb
Public Function RawStringValue(ByVal key As Variant) As String
```

Gets a child string value without unescaping by object key or array index.

```vb
Debug.Print doc.RawStringValue("message")
Debug.Print arr.RawStringValue(0)
```

Modern equivalents:

```vb
Debug.Print doc.RawStringKey("message")
Debug.Print arr.RawStringIndex(0)
```

### StringAt

```vb
Public Function StringAt(ByVal Index As Long) As String
```

Gets an indexed child value as `String`.

```vb
Debug.Print arr.StringAt(0)
```

Modern equivalent:

```vb
Debug.Print arr.StringIndex(0)
```

### NumberAt

```vb
Public Function NumberAt(ByVal Index As Long) As Double
```

Gets an indexed child value as `Double`.

```vb
Debug.Print arr.NumberAt(0)
```

Modern equivalent:

```vb
Debug.Print arr.NumberIndex(0)
```

### BoolAt

```vb
Public Function BoolAt(ByVal Index As Long) As Boolean
```

Gets an indexed child value as `Boolean`.

```vb
Debug.Print arr.BoolAt(0)
```

Modern equivalent:

```vb
Debug.Print arr.BoolIndex(0)
```

### RawStringAt

```vb
Public Function RawStringAt(ByVal Index As Long) As String
```

Gets an indexed string value without unescaping.

```vb
Debug.Print arr.RawStringAt(0)
```

Modern equivalent:

```vb
Debug.Print arr.RawStringIndex(0)
```

## Token Traversal

Token traversal is intended for high-performance loops over large arrays or objects.

It avoids creating a `JSON` wrapper for every child.

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

Use this only when you need object-style access for a specific token. For maximum speed, prefer `TokenString`, `TokenNumber`, `TokenBool`, and `TokenNode` directly.

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

Gets a token value as `String`, applying JSON string unescape for string tokens.

```vb
Debug.Print arr.TokenStringValue(t)
```

### TokenRawStringValue

```vb
Public Function TokenRawStringValue(ByVal TokenId As Long) As String
```

Gets a token string value without unescaping.

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

Given this JSON:

```json
{
  "users": [
    { "name": "Ana", "score": 10, "active": true },
    { "name": "Bia", "score": 20, "active": false }
  ]
}
```

You can iterate without creating one wrapper per user:

```vb
Dim doc As JSON
Set doc = JSON.Parse(responseText)

Dim users As JSON
Set users = doc.NodeKey("users")

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

This is useful for nested objects and arrays:

```vb
Dim rawProfile As String
rawProfile = users.TokenRawField(t, "profile")
```

If `profile` is an object or array, the raw JSON fragment is returned.

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
    Debug.Print profile.StringKey("rank")
End If
```

## Practical Recipes

### Parse an API Response

```vb
Public Sub ReadApiResponse(ByVal responseText As String)
    Dim doc As JSON
    Set doc = JSON.Parse(responseText)

    If Not doc.ExistsKey("data") Then Exit Sub

    Dim data As JSON
    Set data = doc.NodeKey("data")

    If data Is Nothing Then Exit Sub

    Debug.Print data.StringKey("name")
    Debug.Print data.NumberKey("id")
    Debug.Print data.BoolKey("active")
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
    Set user = doc.NodeKey("user")

    If Not user Is Nothing Then
        Debug.Print user.StringKey("name")
        Debug.Print user.StringKey("role")
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
        Debug.Print arr.StringIndex(i)
    Next i
End Sub
```

### Read an Array of Objects

```vb
Public Sub ReadArrayOfObjects(ByVal responseText As String)
    Dim doc As JSON
    Set doc = JSON.Parse(responseText)

    Dim users As JSON
    Set users = doc.NodeKey("users")

    If users Is Nothing Then Exit Sub

    Dim i As Long
    Dim user As JSON

    For i = 0 To users.Count - 1
        Set user = users.NodeIndex(i)

        If Not user Is Nothing Then
            Debug.Print user.StringKey("name")
            Debug.Print user.NumberKey("score")
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
    Set rows = doc.NodeKey("rows")

    If rows Is Nothing Then Exit Sub

    Dim t As Long
    t = rows.FirstChildToken()

    Do While t <> 0
        Debug.Print rows.TokenString(t, "name"), rows.TokenNumber(t, "score")
        t = rows.NextToken(t)
    Loop
End Sub
```

### Extract Keys from an Object

```vb
Public Sub PrintKeys()
    Dim doc As JSON
    Set doc = JSON.Parse("{""name"":""JSON"",""language"":""VBA"",""fast"":true}")

    Dim keys As Variant
    keys = doc.Keys

    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Debug.Print keys(i)
    Next i
End Sub
```

### Extract a Raw Nested Field

```vb
Public Sub ExtractRawPayload(ByVal responseText As String)
    Dim doc As JSON
    Set doc = JSON.Parse(responseText)

    Dim rows As JSON
    Set rows = doc.NodeKey("rows")

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

### Default Member Chaining

```vb
Public Sub ChainRead()
    Dim doc As JSON
    Set doc = JSON.Parse("{""user"":{""badges"" : [""admin"",""dev""]}}")

    Debug.Print doc("user")("badges")(0)
    Debug.Print doc("user")("badges")(1)
End Sub
```

This is concise, but it uses the generic `Variant` path. Prefer typed helpers in hot loops.

## Validation and Tests

A companion validation module can be used to test the parser behavior.

Recommended safe entry points:

```vb
RunJSONSafeSmokeTests
RunAllJSONTests
```

The current validation suite checks common valid JSON cases, including:

- Objects.
- Arrays.
- Primitive roots.
- Escaped strings.
- Unicode escapes.
- Numbers.
- `Stringify` roundtrip.
- `StringifyValue`.
- `Keys`.
- `Exists`, `ExistsKey`, and `ExistsIndex`.
- Token iteration.

A successful run should look like:

```txt
Total:   75
Passed:  75
Failed:  0
```

Invalid JSON tests should be used carefully unless the parser has strict error guards enabled. Invalid-input suites can reveal where validation needs to be improved, but they should not be allowed to freeze the VBE.

## Best Practices

### Keep the Root Document Alive

Node wrappers point back to the document that owns the token buffer.

Good:

```vb
Dim doc As JSON
Set doc = JSON.Parse(text)

Dim data As JSON
Set data = doc.NodeKey("data")

Debug.Print data.Count
```

Avoid returning only a child node from a short-lived local document unless you also keep the root document alive.

### Prefer the New Typed API

Use this style for known schemas:

```vb
Debug.Print doc.StringKey("name")
Debug.Print doc.NumberKey("id")
Debug.Print doc.BoolKey("active")
```

And for arrays:

```vb
Debug.Print arr.StringIndex(0)
Debug.Print arr.NumberIndex(1)
Debug.Print arr.BoolIndex(2)
```

This is clearer than using the generic `Item` or `ValueAt` path everywhere.

### Use ExistsKey and ExistsIndex

Typed accessors return default VBA values when a field is missing.

Use existence checks when default values are ambiguous.

```vb
If doc.ExistsKey("score") Then
    Debug.Print doc.NumberKey("score")
End If
```

For arrays:

```vb
If arr.ExistsIndex(3) Then
    Debug.Print arr.StringIndex(3)
End If
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

### Use StringifyValue for External VBA Values

Use `Stringify` for parsed JSON documents and nodes.

Use `StringifyValue` for normal VBA values.

```vb
Debug.Print doc.Stringify(True)
Debug.Print JSON.StringifyValue(dict, True)
```

### Avoid Generic Variant Access in Hot Loops

This is convenient:

```vb
Debug.Print doc("user")("name")
```

This is better in hot paths:

```vb
Dim user As JSON
Set user = doc.NodeKey("user")
Debug.Print user.StringKey("name")
```

### Use Dictionaries Only for Writing External Objects

The parser itself does not require `Scripting.Dictionary`.

For `StringifyValue`, dictionaries are useful when building JSON objects manually:

```vb
Dim obj As Object
Set obj = CreateObject("Scripting.Dictionary")

obj("name") = "JSON"
obj("ok") = True

Debug.Print JSON.StringifyValue(obj)
```

## Notes and Limitations

### Parser Validation

The parser is optimized for speed and low allocation. It is best used with well-formed JSON input.

For fully untrusted input, strict invalid-input validation should be tested carefully. A separate strict mode can be added later if rejecting every malformed JSON edge case becomes a priority.

### String Escapes

String reading supports common JSON escapes such as:

- `\"`
- `\\`
- `\/`
- `\b`
- `\f`
- `\n`
- `\r`
- `\t`
- `\uXXXX`

Use the normal string helpers for decoded string values:

```vb
Debug.Print doc.StringKey("message")
```

Use raw helpers when you want the stored string slice without unescaping:

```vb
Debug.Print doc.RawStringKey("message")
```

### Number Parsing

Numbers are exposed as `Double` through the numeric helpers.

```vb
Debug.Print doc.NumberKey("price")
```

For exact decimal or financial handling, preserve raw numeric text if needed.

### Missing Values

Most typed accessors return default VBA values when a field is missing or has an unexpected type.

| Accessor | Missing Result |
|:--|:--|
| `StringKey`, `StringIndex` | `""` |
| `NumberKey`, `NumberIndex` | `0` |
| `BoolKey`, `BoolIndex` | `False` |
| `NodeKey`, `NodeIndex` | `Nothing` |
| `ValueAt` | Empty `Variant` |
| `Token...` helpers | Default VBA value |

Use `ExistsKey` or `ExistsIndex` when you need to distinguish missing data from real default values.

### Object Key Comparison

Object key lookup is case-sensitive.

```vb
doc.ExistsKey("name")
doc.ExistsKey("Name")
```

These are different keys.

### Root Primitive Values

JSON root values can be primitives:

```vb
Set doc = JSON.Parse("123")
Debug.Print doc.JsonType
Debug.Print doc.Value
```

Supported root types include object, array, string, number, boolean, and null.

## Troubleshooting

### `NodeKey` Returns Nothing

`NodeKey` only returns objects and arrays.

This returns `Nothing` if `name` is a string:

```vb
Set value = doc.NodeKey("name")
```

Use `StringKey` instead:

```vb
Debug.Print doc.StringKey("name")
```

### `NodeIndex` Returns Nothing

`NodeIndex` only returns object or array items.

If the array item is primitive, use a typed index helper:

```vb
Debug.Print arr.StringIndex(0)
Debug.Print arr.NumberIndex(0)
Debug.Print arr.BoolIndex(0)
```

### `NumberKey` Returns 0

Possible causes:

- The field is missing.
- The field is not a number.
- The number text cannot be converted as expected by VBA.
- The real JSON value is actually `0`.

Check existence first:

```vb
If doc.ExistsKey("score") Then
    Debug.Print doc.NumberKey("score")
End If
```

### `BoolKey` Returns False

`False` can mean the JSON value is actually `false`, the field is missing, or the field is not a boolean.

Check existence first:

```vb
If doc.ExistsKey("active") Then
    Debug.Print doc.BoolKey("active")
End If
```

### `Keys` Seems Empty

`Keys` returns direct keys or direct indexes only.

For primitive nodes, it returns an empty array.

For nested keys, get the nested node first:

```vb
Dim user As JSON
Set user = doc.NodeKey("user")

If Not user Is Nothing Then
    Debug.Print user.Keys()(0)
End If
```

### Large Arrays Feel Slow

Avoid creating a node wrapper for every item in huge arrays:

```vb
For i = 0 To rows.Count - 1
    Set row = rows.NodeIndex(i)
    Debug.Print row.StringKey("name")
Next i
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

### Dictionary Output Requires a Dictionary Object

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
|:--|:--|:--|
| `Parse` | `Parse(ByRef Text As String) As JSON` | Parses JSON text into a tokenized document. |

### Serialization

| API | Signature | Description |
|:--|:--|:--|
| `Stringify` | `Stringify(Optional Pretty As Boolean = False, Optional IndentSize As Long = 2) As String` | Serializes the current document or node. |
| `StringifyWithIndent` | `StringifyWithIndent(Optional Pretty As Boolean = False, Optional IndentText As String = "  ") As String` | Serializes the current document or node with custom indentation. |
| `StringifyValue` | `StringifyValue(Value As Variant, Optional Pretty As Boolean = False, Optional IndentSize As Long = 2) As String` | Serializes an external VBA value. |
| `StringifyValueWithIndent` | `StringifyValueWithIndent(Value As Variant, Optional Pretty As Boolean = False, Optional IndentText As String = "  ") As String` | Serializes an external VBA value with custom indentation. |

### Core Properties

| API | Signature | Description |
|:--|:--|:--|
| `Item` | `Item(key As Variant) As Variant` | Default member. Gets a child value by key or index. |
| `Value` | `Value() As Variant` | Gets the current node value. |
| `Count` | `Count() As Long` | Gets the direct child count. |
| `JsonType` | `JsonType() As String` | Gets the current JSON type name. |
| `IsObject` | `IsObject() As Boolean` | Returns whether the node is an object. |
| `IsArray` | `IsArray() As Boolean` | Returns whether the node is an array. |
| `IsNull` | `IsNull() As Boolean` | Returns whether the node is null. |

### Keys and Existence

| API | Signature | Description |
|:--|:--|:--|
| `Keys` | `Keys() As Variant` | Returns object keys or array indexes. |
| `Exists` | `Exists(key As Variant) As Boolean` | Compatibility helper. Checks whether a field or array index exists. |
| `ExistsKey` | `ExistsKey(key As String) As Boolean` | Checks whether an object key exists. |
| `ExistsIndex` | `ExistsIndex(Index As Long) As Boolean` | Checks whether an array index exists. |

### Generic Child Access

| API | Signature | Description |
|:--|:--|:--|
| `Node` | `Node(key As Variant) As JSON` | Gets a child object or array by key or index. |
| `NodeAt` | `NodeAt(Index As Long) As JSON` | Gets an object/array child by zero-based child position. |
| `ValueAt` | `ValueAt(Index As Long) As Variant` | Gets any child value by zero-based child position. |
| `KeyAt` | `KeyAt(Index As Long) As String` | Gets an object child key by zero-based child position. |

### Typed Object Access

| API | Signature | Description |
|:--|:--|:--|
| `StringKey` | `StringKey(key As String) As String` | Gets an object field as string. |
| `NumberKey` | `NumberKey(key As String) As Double` | Gets an object field as double. |
| `BoolKey` | `BoolKey(key As String) As Boolean` | Gets an object field as boolean. |
| `RawStringKey` | `RawStringKey(key As String) As String` | Gets an object string field without unescaping. |
| `NodeKey` | `NodeKey(key As String) As JSON` | Gets an object field as a JSON node. |

### Typed Array Access

| API | Signature | Description |
|:--|:--|:--|
| `StringIndex` | `StringIndex(Index As Long) As String` | Gets an array item as string. |
| `NumberIndex` | `NumberIndex(Index As Long) As Double` | Gets an array item as double. |
| `BoolIndex` | `BoolIndex(Index As Long) As Boolean` | Gets an array item as boolean. |
| `RawStringIndex` | `RawStringIndex(Index As Long) As String` | Gets an array string item without unescaping. |
| `NodeIndex` | `NodeIndex(Index As Long) As JSON` | Gets an array item as a JSON node. |

### Legacy Typed Accessors

| API | Signature | Description |
|:--|:--|:--|
| `StringValue` | `StringValue(key As Variant) As String` | Gets a child value as string by key or index. |
| `NumberValue` | `NumberValue(key As Variant) As Double` | Gets a child value as double by key or index. |
| `BoolValue` | `BoolValue(key As Variant) As Boolean` | Gets a child value as boolean by key or index. |
| `RawStringValue` | `RawStringValue(key As Variant) As String` | Gets a child string without unescaping by key or index. |
| `StringAt` | `StringAt(Index As Long) As String` | Gets an indexed child value as string. |
| `NumberAt` | `NumberAt(Index As Long) As Double` | Gets an indexed child value as double. |
| `BoolAt` | `BoolAt(Index As Long) As Boolean` | Gets an indexed child value as boolean. |
| `RawStringAt` | `RawStringAt(Index As Long) As String` | Gets an indexed child string without unescaping. |

### Token Traversal

| API | Signature | Description |
|:--|:--|:--|
| `FirstChildToken` | `FirstChildToken() As Long` | Gets the first direct child token. |
| `LastChildToken` | `LastChildToken() As Long` | Gets the last direct child token. |
| `NextToken` | `NextToken(TokenId As Long) As Long` | Gets the next sibling token. |
| `NodeFromToken` | `NodeFromToken(TokenId As Long) As JSON` | Wraps an object/array token as a node. |

### Token Value Access

| API | Signature | Description |
|:--|:--|:--|
| `TokenKey` | `TokenKey(TokenId As Long) As String` | Gets the key associated with a token. |
| `TokenValue` | `TokenValue(TokenId As Long) As Variant` | Gets a token value as Variant. |
| `TokenStringValue` | `TokenStringValue(TokenId As Long) As String` | Gets a token value as string. |
| `TokenRawStringValue` | `TokenRawStringValue(TokenId As Long) As String` | Gets a token string without unescaping. |
| `TokenNumberValue` | `TokenNumberValue(TokenId As Long) As Double` | Gets a token value as double. |
| `TokenBoolValue` | `TokenBoolValue(TokenId As Long) As Boolean` | Gets a token value as boolean. |

### Token Field Access

| API | Signature | Description |
|:--|:--|:--|
| `TokenString` | `TokenString(TokenId As Long, key As Variant) As String` | Gets a field from an object token as string. |
| `TokenRawString` | `TokenRawString(TokenId As Long, key As Variant) As String` | Gets a field from an object token as raw string. |
| `TokenRawField` | `TokenRawField(TokenId As Long, key As String) As String` | Gets a raw field slice from an object token. |
| `TokenNumber` | `TokenNumber(TokenId As Long, key As Variant) As Double` | Gets a field from an object token as double. |
| `TokenBool` | `TokenBool(TokenId As Long, key As Variant) As Boolean` | Gets a field from an object token as boolean. |
| `TokenNode` | `TokenNode(TokenId As Long, key As Variant) As JSON` | Gets a field from an object token as a JSON node. |

## License

MIT. Designed for fast JSON parsing, clean traversal, low allocation, and practical data automation inside Microsoft Office.
