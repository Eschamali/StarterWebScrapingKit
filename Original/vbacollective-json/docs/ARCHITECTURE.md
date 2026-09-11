# JSON Architecture

JSON is a single-file VBA class (`JSON.cls`) that implements a fast JSON reader and lightweight JSON writer for Microsoft Office hosts.

It is designed around compact token-tree parsing, lazy node wrappers, typed accessors, raw field access, token iteration, `Keys`/`Exists` helpers, and `Stringify` support while staying fully contained inside one importable `.cls` file.

This document describes the architecture used by the current `JSON_typed_api.cls` version.

## Table of Contents

- [High-Level Overview](#high-level-overview)
- [Design Goals](#design-goals)
- [Runtime Model](#runtime-model)
- [Class Layout](#class-layout)
- [Core Data Structures](#core-data-structures)
- [Document State](#document-state)
- [Node Wrapper State](#node-wrapper-state)
- [Parsing Pipeline](#parsing-pipeline)
- [Token Allocation](#token-allocation)
- [Object and Array Parsing](#object-and-array-parsing)
- [String Parsing](#string-parsing)
- [Number and Primitive Parsing](#number-and-primitive-parsing)
- [Token Tree Linking](#token-tree-linking)
- [Lookup Architecture](#lookup-architecture)
- [Modern Typed API Layer](#modern-typed-api-layer)
- [Compatibility API Layer](#compatibility-api-layer)
- [Keys and Exists](#keys-and-exists)
- [Value Conversion](#value-conversion)
- [Raw Slice Access](#raw-slice-access)
- [Lazy Node Wrappers](#lazy-node-wrappers)
- [Token Iteration](#token-iteration)
- [Stringify Architecture](#stringify-architecture)
- [External Value Serialization](#external-value-serialization)
- [Memory Model](#memory-model)
- [Performance Strategy](#performance-strategy)
- [Validation and Testing Strategy](#validation-and-testing-strategy)
- [Compatibility Strategy](#compatibility-strategy)
- [Shutdown and Cleanup](#shutdown-and-cleanup)
- [Known Architectural Boundaries](#known-architectural-boundaries)
- [Summary](#summary)

## High-Level Overview

JSON parses text into a compact internal token tree instead of immediately converting the entire payload into nested `Scripting.Dictionary`, `Collection`, or one VBA object per node.

The public usage is intentionally simple:

```vb
Dim doc As JSON
Set doc = JSON.Parse(responseText)

Debug.Print doc.StringKey("name")
Debug.Print doc.NumberKey("id")
Debug.Print doc.BoolKey("active")
```

Internally, the parser stores each JSON value as a token. A token records the value type, parent/child links, sibling links, key slice, value slice, and child count.

```txt
JSON text
  -> parser cursor
  -> token buffer
  -> root JSON document
  -> typed access / lazy nodes / token iteration / Stringify
```

The main architectural idea is:

```txt
Do the minimum work during parsing.
Convert, wrap, copy, or allocate only when user code asks for data.
```

## Design Goals

| Goal | Architectural Choice |
|:---|:---|
| Single-file deployment | Everything lives inside `JSON.cls`. |
| No required references | The parser does not depend on `Scripting.Dictionary`. |
| Fast parsing | Builds compact tokens instead of nested containers. |
| Low allocation | Stores key/value positions and lengths instead of eagerly copying values. |
| Comfortable API | Provides `StringKey`, `NumberKey`, `BoolKey`, `NodeKey`, `Keys`, and `Exists`. |
| Compatibility | Keeps older generic methods such as `Item`, `Value`, `ValueAt`, `StringValue`, and `Node`. |
| Large-array support | Exposes token iteration to avoid wrapper allocation per row. |
| Practical writing | Supports `Stringify` for parsed nodes and `StringifyValue` for normal VBA values. |
| Office compatibility | Works in normal VBA hosts and supports x86/x64 conditional declarations. |

JSON is not meant to be a schema validator. It is a fast practical JSON reader/writer for API responses, local config files, automation payloads, and Office/VBA tooling.

## Runtime Model

The runtime has two object modes:

```txt
Root document = owns the parsed text and token buffer
Node wrapper  = points to a token owned by the root document
```

A root document owns:

```txt
m_Text       = original JSON text
m_Tokens()   = parsed token tree
m_RootId     = root token id
```

A child node wrapper owns:

```txt
m_NodeId     = wrapped token id
m_Document   = reference to the root document
```

Example:

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""user"":{""name"":""Ueslei""}}")

Dim user As JSON
Set user = doc.NodeKey("user")

Debug.Print user.StringKey("name")
```

The `user` object does not duplicate the nested object. It only stores a reference back to `doc` plus the token id for the `user` object.

## Class Layout

`JSON.cls` is intended to be a predeclared class:

```vb
Attribute VB_Name = "JSON"
Attribute VB_PredeclaredId = True
```

That enables factory-style usage:

```vb
Set doc = JSON.Parse(text)
Debug.Print JSON.StringifyValue(value, True)
```

The same class contains:

1. Public API methods.
2. Root document state.
3. Node wrapper state.
4. Parser implementation.
5. Token lookup implementation.
6. Stringify writer implementation.

This is why distribution stays simple: import one `JSON.cls` file.

## Core Data Structures

### JSType

The internal JSON type enum represents parsed value types.

```vb
Private Enum JSType
    jsNone = 0
    jsObject = 1
    jsArray = 2
    jsString = 3
    jsNumber = 4
    jsBool = 5
    jsNull = 6
End Enum
```

The public `JsonType` property converts these to user-facing strings:

| Internal Type | Public Type |
|:---|:---|
| `jsObject` | `object` |
| `jsArray` | `array` |
| `jsString` | `string` |
| `jsNumber` | `number` |
| `jsBool` | `boolean` |
| `jsNull` | `null` |

### JSToken

`JSToken` is the core internal record.

```vb
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
```

Each token stores hierarchy and text slice metadata.

| Field | Purpose |
|:---|:---|
| `Type` | Internal JSON type. |
| `Parent` | Parent token id. |
| `NextSibling` | Next token under the same parent. |
| `FirstChild` | First direct child token. |
| `LastChild` | Last direct child token. |
| `ChildCount` | Number of direct children. |
| `KeyStart` | One-based start position of object key. |
| `KeyLen` | Object key length. |
| `ValStart` | One-based start position of value text. |
| `ValLen` | Value text length. |

Example token layout:

```json
{
  "name": "Ueslei",
  "age": 18
}
```

```txt
Token 1: object
  FirstChild -> 2
  LastChild  -> 3

Token 2: string
  Parent      -> 1
  KeyStart    -> slice for name
  ValStart    -> slice for Ueslei
  NextSibling -> 3

Token 3: number
  Parent      -> 1
  KeyStart    -> slice for age
  ValStart    -> slice for 18
```

## Document State

The root document stores parser and token state.

Typical fields include:

```vb
Private m_Text As String
Private m_Length As Long
Private m_Index As Long
Private m_Tokens() As JSToken
Private m_TokenCount As Long
Private m_TokenCap As Long
Private m_RootId As Long
Private m_NodeId As Long
Private m_Document As JSON
```

Depending on the exact build, the class may also include a character buffer or SAFEARRAY alias fields used for faster character access.

| Field | Purpose |
|:---|:---|
| `m_Text` | Original JSON source text. |
| `m_Length` | Length of source text. |
| `m_Index` | Current parser cursor. |
| `m_Tokens()` | Token buffer. |
| `m_TokenCount` | Number of active tokens. |
| `m_TokenCap` | Current token capacity. |
| `m_RootId` | Root token id. |
| `m_NodeId` | Wrapped token id for child node wrappers. |
| `m_Document` | Root document reference for wrappers. |

The root document must stay alive while child node wrappers are being used.

## Node Wrapper State

A node wrapper is a lightweight `JSON` instance that points to a token inside another root document.

Conceptually:

```vb
Friend Sub InitNode(ByVal TokenId As Long, ByVal Document As JSON)
    m_NodeId = TokenId
    Set m_Document = Document
End Sub
```

The wrapper does not own the source text or token buffer.

Most public methods resolve their execution target like this:

```txt
If this object is a node wrapper:
    document = m_Document
    baseId   = m_NodeId
Else:
    document = Me
    baseId   = m_RootId
```

That lets the same API work on both root documents and nested nodes:

```vb
Debug.Print doc.StringKey("name")
Debug.Print user.StringKey("name")
Debug.Print items.StringIndex(0)
```

## Parsing Pipeline

Parsing starts at `Parse`:

```vb
Public Function Parse(ByRef Text As String) As JSON
    Dim doc As JSON
    Set doc = New JSON
    doc.LoadText Text
    Set Parse = doc
End Function
```

The full pipeline is:

```txt
JSON.Parse(text)
  -> create new JSON document
  -> reset state
  -> store source text
  -> allocate token buffer
  -> parse root value
  -> return document
```

The parser does not build Dictionaries or Collections. It only builds the token tree.

## Token Allocation

Tokens are allocated sequentially.

Conceptual implementation:

```vb
Private Function AddToken() As Long
    m_TokenCount = m_TokenCount + 1

    If m_TokenCount > m_TokenCap Then
        m_TokenCap = m_TokenCap * 2
        ReDim Preserve m_Tokens(1 To m_TokenCap)
    End If

    AddToken = m_TokenCount
End Function
```

The important performance rule is to avoid repeated small `ReDim Preserve` operations. The parser uses an estimated initial capacity and grows in larger blocks.

Why this matters:

```txt
Few large resizes = good
Many tiny resizes = slow
```

## Object and Array Parsing

### Objects

Object parsing follows this model:

```txt
consume {
repeat:
    skip whitespace
    if } then finish
    scan key string
    skip whitespace
    consume :
    parse child value
    link child token to object token
    skip whitespace
    if , continue
    if } finish
```

Object keys are stored as text slices:

```txt
KeyStart
KeyLen
```

The key is not converted to a VBA string unless user code asks for it through methods such as `KeyAt`, `Keys`, `TokenKey`, or serialization.

### Arrays

Array parsing follows this model:

```txt
consume [
repeat:
    skip whitespace
    if ] then finish
    parse child value
    link child token to array token
    skip whitespace
    if , continue
    if ] finish
```

Array items are sibling tokens under the array token.

```txt
array token
  FirstChild -> item 0
  item 0 NextSibling -> item 1
  item 1 NextSibling -> item 2
```

## String Parsing

String tokens store the slice inside the quotes.

```txt
"Ueslei"
 ^    ^
 value slice only
```

Stored token fields:

```txt
ValStart = first character after opening quote
ValLen   = number of characters before closing quote
```

When the string is read through `StringKey`, `StringIndex`, `StringValue`, or `TokenStringValue`, the raw slice is copied and then unescaped if needed.

This keeps parse time cheaper because strings are not fully processed unless user code reads them.

## Number and Primitive Parsing

Numbers are stored as raw slices.

```txt
-123.45e+6
 ^       ^
 raw numeric slice
```

Conversion to `Double` happens only when requested:

```vb
Debug.Print doc.NumberKey("score")
```

Booleans and null are also represented as tokens:

| JSON Literal | Internal Type |
|:---|:---|
| `true` | `jsBool` |
| `false` | `jsBool` |
| `null` | `jsNull` |

This means parsing remains focused on structure and tokenization, not on eagerly converting every value.

## Token Tree Linking

Every child token is linked to its parent.

Conceptual link operation:

```txt
If parent has no first child:
    parent.FirstChild = child
Else:
    parent.LastChild.NextSibling = child

parent.LastChild = child
parent.ChildCount += 1
child.Parent = parent
```

This gives every object and array a forward-linked list of direct children.

Benefits:

- very cheap append during parsing;
- no container object allocation;
- simple traversal;
- token iteration can walk large arrays without wrapper objects.

## Lookup Architecture

Lookup resolves a child token by either object key or array index.

Generic compatibility methods use a `Variant` key:

```vb
Debug.Print doc.Item("name")
Debug.Print arr.Item(0)
Debug.Print doc.Exists("name")
Debug.Print arr.Exists(0)
```

Modern methods avoid ambiguity by splitting object-key access from array-index access:

```vb
Debug.Print doc.StringKey("name")
Debug.Print arr.StringIndex(0)
Debug.Print doc.ExistsKey("name")
Debug.Print arr.ExistsIndex(0)
```

This is cleaner and avoids using `Variant` in the common path.

## Modern Typed API Layer

The modern API added to the class is the preferred usage style.

### Object-key methods

```vb
ExistsKey(ByRef Key As String) As Boolean
StringKey(ByRef Key As String) As String
NumberKey(ByRef Key As String) As Double
BoolKey(ByRef Key As String) As Boolean
RawStringKey(ByRef Key As String) As String
NodeKey(ByRef Key As String) As JSON
```

Example:

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""name"":""Ueslei"",""age"":18,""active"":true}")

If doc.ExistsKey("name") Then
    Debug.Print doc.StringKey("name")
End If

Debug.Print doc.NumberKey("age")
Debug.Print doc.BoolKey("active")
```

### Array-index methods

```vb
ExistsIndex(ByVal Index As Long) As Boolean
StringIndex(ByVal Index As Long) As String
NumberIndex(ByVal Index As Long) As Double
BoolIndex(ByVal Index As Long) As Boolean
RawStringIndex(ByVal Index As Long) As String
NodeIndex(ByVal Index As Long) As JSON
```

Example:

```vb
Dim arr As JSON
Set arr = JSON.Parse("[""VBA"",""JSON"",true]")

Debug.Print arr.StringIndex(0)
Debug.Print arr.StringIndex(1)
Debug.Print arr.BoolIndex(2)
```

### Why these methods exist

The original generic API is flexible, but it uses `Variant` because a JSON child can be read by string key or numeric index.

The typed API makes common usage faster and clearer:

```txt
Known object field -> use *Key
Known array index  -> use *Index
Unknown/generic    -> use Item/Value/Exists compatibility methods
```

## Compatibility API Layer

The class keeps older generic methods so existing code continues to work.

Examples:

```vb
Debug.Print doc("name")
Debug.Print doc.StringValue("name")
Debug.Print doc.NumberValue("age")
Debug.Print doc.BoolValue("active")

Dim user As JSON
Set user = doc.Node("user")
```

These methods are still useful when writing quick scripts or when the caller naturally has a `Variant` key.

The recommended style for new code is:

```vb
Debug.Print doc.StringKey("name")
Set user = doc.NodeKey("user")
Debug.Print arr.StringIndex(0)
```

## Keys and Exists

### Keys

`Keys()` returns a `Variant` array of direct child identifiers.

For objects, it returns object keys:

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""name"":""Ueslei"",""age"":18}")

Dim keys As Variant
keys = doc.Keys

Debug.Print keys(0)  ' name
Debug.Print keys(1)  ' age
```

For arrays, it returns indexes:

```vb
Dim arr As JSON
Set arr = JSON.Parse("[10,20,30]")

Dim keys As Variant
keys = arr.Keys

Debug.Print keys(0)  ' 0
Debug.Print keys(1)  ' 1
Debug.Print keys(2)  ' 2
```

Why this exists:

- the parser does not materialize `Scripting.Dictionary` objects;
- users still need a familiar way to inspect object fields;
- it mirrors one of the most useful `Dictionary` helpers without depending on Dictionary internally.

### Exists

There are three existence helpers:

```vb
Exists(key As Variant) As Boolean
ExistsKey(ByRef Key As String) As Boolean
ExistsIndex(ByVal Index As Long) As Boolean
```

Use `ExistsKey` for objects:

```vb
If doc.ExistsKey("user") Then
    Set user = doc.NodeKey("user")
End If
```

Use `ExistsIndex` for arrays:

```vb
If arr.ExistsIndex(0) Then
    Debug.Print arr.StringIndex(0)
End If
```

Use `Exists` for compatibility or generic code:

```vb
Debug.Print doc.Exists("name")
Debug.Print arr.Exists(0)
```

## Value Conversion

Primitive values are converted lazily.

### String

String access copies the stored slice and applies JSON unescaping if needed.

```vb
Debug.Print doc.StringKey("message")
```

Raw string access skips unescaping:

```vb
Debug.Print doc.RawStringKey("message")
```

### Number

Number access converts the stored numeric slice to `Double`.

```vb
Debug.Print doc.NumberKey("price")
```

Missing or incompatible values return `0`.

### Boolean

Boolean access checks the token type and literal.

```vb
If doc.BoolKey("active") Then
    Debug.Print "active"
End If
```

Missing or incompatible values return `False`.

### Null

`IsNull` returns whether the current node is JSON null.

String/numeric/boolean accessors return default values when pointed at `null`.

## Raw Slice Access

Raw access returns text based on the original stored value slice.

Used by:

- `RawStringKey`
- `RawStringIndex`
- `RawStringValue`
- `RawStringAt`
- `TokenRawStringValue`
- `TokenRawString`
- `TokenRawField`

Raw access is useful when you want to forward or cache nested JSON without walking or re-serializing the subtree.

Example:

```vb
Dim rawPayload As String
rawPayload = rows.TokenRawField(t, "payload")
```

## Lazy Node Wrappers

Objects and arrays are returned as lightweight wrappers.

```vb
Dim user As JSON
Set user = doc.NodeKey("user")
```

Internally:

```txt
resolve token for user
if token is object or array:
    create new JSON wrapper
    wrapper.m_NodeId = token id
    wrapper.m_Document = root document
```

Wrappers are created only on demand.

Methods that can return wrappers include:

- `Item`
- `Value`
- `ValueAt`
- `Node`
- `NodeAt`
- `NodeKey`
- `NodeIndex`
- `TokenValue`
- `TokenNode`
- `NodeFromToken`

Keep the root document alive while using child wrappers.

## Token Iteration

Token iteration is the fastest traversal model for large arrays.

```vb
Dim rows As JSON
Set rows = doc.NodeKey("rows")

Dim t As Long
t = rows.FirstChildToken()

Do While t <> 0
    Debug.Print rows.TokenString(t, "name")
    Debug.Print rows.TokenNumber(t, "score")
    Debug.Print rows.TokenBool(t, "active")

    t = rows.NextToken(t)
Loop
```

This avoids:

- creating a `JSON` wrapper for every row;
- repeatedly scanning array indexes with `NodeIndex(i)`;
- returning generic `Variant` objects in tight loops.

Use token iteration when processing large arrays of objects.

## Stringify Architecture

The writer has two main paths:

```txt
Parsed JSON document/node -> Stringify
External VBA value        -> StringifyValue
```

Public entry points:

```vb
doc.Stringify()
doc.Stringify(True)
doc.StringifyWithIndent(True, vbTab)

JSON.StringifyValue(value)
JSON.StringifyValue(value, True)
JSON.StringifyValueWithIndent(value, True, vbTab)
```

### Parsed token serialization

`Stringify` resolves the current base token and recursively writes it.

| Token Type | Serialization |
|:---|:---|
| Object | `{ ... }` |
| Array | `[ ... ]` |
| String | quoted/escaped JSON string |
| Number | numeric text |
| Boolean | `true` / `false` |
| Null | `null` |

### Pretty printing

Pretty output adds:

- `vbCrLf`
- indentation
- spacing after `:`

Examples:

```vb
Debug.Print doc.Stringify(True)
Debug.Print doc.Stringify(True, 4)
Debug.Print doc.StringifyWithIndent(True, vbTab)
```

Compact output is recommended for storage, transport, or speed-sensitive paths.

## External Value Serialization

`StringifyValue` serializes normal VBA values.

Supported values include:

| VBA Value | JSON Output |
|:---|:---|
| `String` | JSON string |
| `Boolean` | `true` / `false` |
| Numeric types | JSON number |
| `Currency` | JSON number |
| `Date` | JSON string |
| `Null` | `null` |
| `Empty` | `null` |
| One-dimensional VBA array | JSON array |
| `Collection` | JSON array |
| `Dictionary` / `Scripting.Dictionary` | JSON object |
| `JSON` document/node | serialized JSON |
| Unsupported object | `null` |

Example:

```vb
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

dict("name") = "JSON"
dict("version") = "1.0.1"
dict("fast") = True

Debug.Print JSON.StringifyValue(dict, True)
```

The parser itself does not require Dictionary. Dictionary support here exists only for writing external VBA objects to JSON.

## Memory Model

The memory model is centered around:

```txt
Source text
Token buffer
Lazy wrappers
```

### Source text

The original JSON text remains stored in the document.

Key/value tokens reference positions and lengths inside this text.

### Token buffer

The token buffer is a dynamic array of compact `JSToken` records.

No Dictionary, Collection, or per-node wrapper is allocated during parse.

### Lazy wrappers

Wrappers are created only when code asks for nested objects/arrays.

This keeps parsing cheap and makes large payloads more practical in VBA.

## Performance Strategy

Main performance choices:

| Area | Strategy |
|:---|:---|
| Parse output | Compact token tree. |
| Object lookup | Scan direct children without materializing Dictionary. |
| Value conversion | Convert only when requested. |
| Object/array nodes | Create wrappers lazily. |
| Large arrays | Use token iteration. |
| Compatibility | Keep generic `Variant` methods, but add typed methods for common paths. |
| Writing | Serialize parsed tokens or external values directly. |

### Fast path: known object fields

```vb
Debug.Print doc.StringKey("name")
Debug.Print doc.NumberKey("score")
Debug.Print doc.BoolKey("active")
```

This avoids the generic `Variant` route where possible.

### Fast path: known array indexes

```vb
Debug.Print arr.StringIndex(0)
Debug.Print arr.NumberIndex(1)
Debug.Print arr.BoolIndex(2)
```

### Fast path: large array scan

```vb
Dim t As Long
t = rows.FirstChildToken()

Do While t <> 0
    Debug.Print rows.TokenString(t, "name")
    t = rows.NextToken(t)
Loop
```

This is preferable to repeatedly calling `NodeIndex(i)` on huge arrays.

## Validation and Testing Strategy

The validation suite added for this version focuses on proving the safe, normal JSON path.

Covered test groups:

- valid objects;
- valid arrays;
- primitive roots;
- escaped strings;
- unicode escape handling;
- numbers;
- `Stringify` roundtrip;
- `StringifyValue`;
- `Keys` and `Exists`;
- token iteration.

The safe validation run produced:

```txt
Total:   75
Passed:  75
Failed:  0
```

The invalid JSON suite should be treated carefully until the parser has stronger internal progress guards. Invalid-input tests must never be allowed to trap the VBE in an infinite parser loop.

Recommended testing order:

```vb
RunJSONSafeSmokeTests
RunAllJSONTests
```

Run strict invalid-input tests only after parser guard improvements.

## Compatibility Strategy

JSON is meant to work in normal Office VBA hosts.

The class is designed around:

- one `.cls` import;
- no required project references;
- x86/x64 conditional declarations where needed;
- late-bound `Scripting.Dictionary` support only for external value serialization;
- compatibility methods kept alongside the modern typed API.

This means users can start simple:

```vb
Set doc = JSON.Parse(text)
Debug.Print doc("name")
```

Then move to the more explicit typed API when performance and clarity matter:

```vb
Debug.Print doc.StringKey("name")
```

## Shutdown and Cleanup

If a build uses a native character alias or SAFEARRAY descriptor, cleanup must clear that alias before object destruction.

Conceptual cleanup:

```vb
Private Sub Class_Terminate()
    ClearCharAlias
End Sub
```

The purpose is to prevent VBA from trying to free memory it does not own.

For builds that copy characters instead of aliasing, cleanup is simpler, but the rule remains the same: root document state owns parse buffers; node wrappers only reference the root document.

## Known Architectural Boundaries

### Parser strictness

The parser is optimized for speed and normal well-formed JSON.

It is not currently a full strict diagnostic parser with detailed syntax errors.

For untrusted input, strict validation should be added or performed upstream.

### Invalid input safety

Some malformed inputs can expose parser edge cases if internal progress guards are missing.

A future hardening pass should ensure:

```txt
Every parser loop either advances the cursor or exits with failure.
```

This prevents infinite loops on invalid JSON.

### Unicode escapes

The current lightweight string handling supports common escapes and tested unicode cases. Full surrogate-pair normalization and advanced Unicode validation may require a dedicated stricter unescape path.

### Object lookup complexity

Object lookup scans direct children linearly.

This avoids building a Dictionary per object, which keeps parsing cheaper.

For repeated heavy lookup against the same very large object, a future optional hash/index layer could improve access speed.

### Array index complexity

`NodeIndex(i)`, `ValueAt(i)`, and similar indexed methods walk sibling tokens to find the requested child.

For sequential scans of large arrays, token iteration is better.

Recommended:

```vb
Dim t As Long
t = rows.FirstChildToken()

Do While t <> 0
    Debug.Print rows.TokenString(t, "name")
    t = rows.NextToken(t)
Loop
```

Avoid for huge arrays:

```vb
For i = 0 To rows.Count - 1
    Set row = rows.NodeIndex(i)
Next i
```

### Variant cannot disappear completely

JSON values are naturally mixed-type.

The class reduces `Variant` use in the preferred API, but it still keeps `Variant` where it is necessary:

- `Item` default member;
- `Value`;
- `ValueAt`;
- `TokenValue`;
- `StringifyValue`;
- compatibility methods that accept either string keys or numeric indexes.

That is intentional. The modern `*Key` and `*Index` methods exist so hot paths do not need to use the generic route.

### Threading

JSON is synchronous.

It does not create threads, background workers, timers, or async jobs.

All parsing, traversal, and serialization happen in the calling VBA procedure.

## Summary

JSON's architecture is built around one practical idea:

```txt
Parse once into compact tokens.
Read through typed helpers, lazy nodes, or token iteration.
Serialize only when needed.
```

The current design gives VBA code:

```txt
Single .cls file
    -> compact token tree
    -> no internal Dictionary dependency
    -> typed object-key API
    -> typed array-index API
    -> Keys and Exists helpers
    -> lazy node wrappers
    -> token iteration
    -> raw field extraction
    -> lightweight Stringify pipeline
```

This makes the class suitable for fast Office automation, API response parsing, configuration loading, data extraction, and practical JSON writing from VBA.

## License

MIT. Designed for fast JSON parsing, clean traversal, low allocation, and practical data automation inside Microsoft Office.
