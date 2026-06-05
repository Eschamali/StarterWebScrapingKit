<div align="center">
  <img src="resources/logo.png" width="160" alt="JSON logo" />
</div>

<h1 align="center">JSON</h1>

<p align="center">
  <b>A high-performance, zero-copy JSON parser and lightweight writer for VBA.</b><br/>
  Fast parsing, typed accessors, raw field access, token iteration, lazy node wrappers, and Stringify support.
</p>

<p align="center">
  <img src="https://github.com/vbacollective/json/actions/workflows/ci.yml/badge.svg" alt="CI" />
  <img src="https://img.shields.io/badge/version-v1.0.1-blue.svg" alt="Version" />
  <img src="https://img.shields.io/badge/language-VBA-867DB1.svg" alt="Language" />
  <img src="https://img.shields.io/badge/platform-Windows-0078D6.svg" alt="Platform" />
  <img src="https://img.shields.io/badge/arch-32%20%26%2064--bit-green.svg" alt="Architecture" />
  <img src="https://img.shields.io/badge/parser-zero--copy-success.svg" alt="Zero-copy parser" />
  <img src="https://img.shields.io/badge/dependencies-none-success.svg" alt="Dependencies" />
  <img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License" />
</p>

**JSON** is a compact, production-oriented JSON reader and writer for VBA. It is designed for developers who need fast JSON parsing, lightweight traversal, and reliable serialization inside Microsoft Office applications without external DLLs, ActiveX controls, installers, or additional references.

Unlike traditional VBA JSON libraries that eagerly allocate nested Dictionaries, Collections, and objects while parsing, this module parses JSON into a compact internal token tree. Nodes are created lazily only when requested, allowing large JSON documents to be traversed with much lower allocation pressure.

Whether you are building Excel automation tools, PowerPoint dashboards, Access integrations, Office add-ins, local data pipelines, or API clients, this module provides a practical JSON layer for parsing, reading, iterating, and writing JSON data directly from VBA.

## Key Capabilities

* **Zero Dependencies:** No external libraries, references, DLLs, or installers required.
* **Single-File Distribution:** Import one `.cls` file and use it immediately.
* **Zero-Copy Parsing:** Keeps slices into the original JSON text instead of allocating every key and value during parse.
* **Compact Token Tree:** Stores hierarchy, siblings, keys, values, and child counts in lightweight token structures.
* **Lazy Node Wrappers:** JSON node objects are only created when requested.
* **Typed Accessors:** Read strings, numbers, booleans, arrays, objects, and nulls through focused helper methods.
* **Raw Field Access:** Extract raw JSON fragments without fully materializing nested objects.
* **Token Iteration:** Iterate huge arrays and objects through token handles for better performance.
* **Object and Array Support:** Traverse JSON objects by key and arrays by index.
* **Stringify Support:** Serialize parsed JSON, primitive VBA values, arrays, Collections, Dictionaries, and nested JSON nodes.
* **Pretty Printing:** Generate compact or formatted JSON with custom indentation.
* **VBA-Friendly API:** Simple `Parse`, `Item`, `Value`, `Count`, `Exists`, `Node`, and `Stringify` methods.
* **Architecture Aware:** Compatible with both 32-bit and 64-bit Office through `#If VBA7` declarations.

## Getting Started

### Installation

1. Download the latest `JSON.cls`.
2. Open the VBA Editor with `Alt + F11`.
3. Choose **File > Import File...** and select `JSON.cls`.
4. No external references are required.
5. Save your Office document as a macro-enabled file, such as `.xlsm`, `.pptm`, `.docm`, or `.accdb`.

## Minimal Usage

Parse a JSON string and read values directly.

```vb
Public Sub ReadJson()
    Dim text As String
    text = "{""name"":""Ueslei"",""age"":18,""active"":true}"

    Dim doc As JSON
    Set doc = JSON.Parse(text)

    Debug.Print doc.StringValue("name")
    Debug.Print doc.NumberValue("age")
    Debug.Print doc.BoolValue("active")
End Sub
```

## Basic Usage

### Parse JSON

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""project"":""JSON"",""language"":""VBA""}")
```

### Read Object Fields

```vb
Debug.Print doc.StringValue("project")
Debug.Print doc.StringValue("language")
```

### Check if a Field Exists

```vb
If doc.Exists("project") Then
    Debug.Print "Project:", doc.StringValue("project")
End If
```

### Read Nested Objects

```vb
Dim text As String
text = "{""user"":{""name"":""Ueslei"",""role"":""developer""}}"

Dim doc As JSON
Set doc = JSON.Parse(text)

Dim user As JSON
Set user = doc.Node("user")

Debug.Print user.StringValue("name")
Debug.Print user.StringValue("role")
```

### Read Arrays

```vb
Dim text As String
text = "{""items"":[""sword"",""shield"",""potion""]}"

Dim doc As JSON
Set doc = JSON.Parse(text)

Dim items As JSON
Set items = doc.Node("items")

Dim i As Long
For i = 0 To items.Count - 1
    Debug.Print items.StringAt(i)
Next
```

### Read Mixed Values

```vb
Dim text As String
text = "{""name"":""Potion"",""price"":25.5,""stackable"":true,""meta"":null}"

Dim doc As JSON
Set doc = JSON.Parse(text)

Debug.Print doc.StringValue("name")
Debug.Print doc.NumberValue("price")
Debug.Print doc.BoolValue("stackable")
Debug.Print doc.Node("meta").IsNull
```

## Stringify

### Serialize Parsed JSON

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""name"":""Ueslei"",""skills"":[""VBA"",""Rust"",""JS""]}")

Debug.Print doc.Stringify()
```

### Pretty Print JSON

```vb
Debug.Print doc.Stringify(True)
```

Output:

```json
{
  "name": "Ueslei",
  "skills": [
    "VBA",
    "Rust",
    "JS"
  ]
}
```

### Custom Indentation

Use a custom indentation string, such as tabs.

```vb
Debug.Print doc.StringifyWithIndent(True, vbTab)
```

### Serialize VBA Values

```vb
Dim values(0 To 2) As Variant
values(0) = "VBA"
values(1) = "JSON"
values(2) = True

Debug.Print JSON.StringifyValue(values)
Debug.Print JSON.StringifyValue(values, True)
```

### Serialize Collections

```vb
Dim list As Collection
Set list = New Collection

list.Add "Excel"
list.Add "PowerPoint"
list.Add "Access"

Debug.Print JSON.StringifyValue(list, True)
```

### Serialize Dictionaries

```vb
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

dict("name") = "JSON"
dict("language") = "VBA"
dict("fast") = True

Debug.Print JSON.StringifyValue(dict, True)
```

## Traversal API

The standard object API is ideal for most projects.

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""user"":{""name"":""Ueslei""},""score"":100}")

Debug.Print doc.Item("score")
Debug.Print doc.Value
Debug.Print doc.Count
Debug.Print doc.JsonType
Debug.Print doc.IsObject
Debug.Print doc.IsArray
Debug.Print doc.IsNull
```

### Common Methods

| Method                                                | Description                                                         |
| :---------------------------------------------------- | :------------------------------------------------------------------ |
| `Parse(text)`                                         | Parses JSON text into a document.                                   |
| `Stringify(pretty, indentSize)`                       | Serializes the current document or node.                            |
| `StringifyWithIndent(pretty, indentText)`             | Serializes using a custom indentation string.                       |
| `StringifyValue(value, pretty, indentSize)`           | Serializes an external VBA value.                                   |
| `StringifyValueWithIndent(value, pretty, indentText)` | Serializes an external VBA value using a custom indentation string. |
| `Item(key)`                                           | Reads a child value by object key or array index.                   |
| `Value`                                               | Reads the current node as a VBA value.                              |
| `Count`                                               | Returns the number of children in an object or array.               |
| `JsonType`                                            | Returns the current node type.                                      |
| `Exists(key)`                                         | Checks whether an object field or array index exists.               |
| `Node(key)`                                           | Returns a child node by object key or array index.                  |
| `NodeAt(index)`                                       | Returns an array/object child node by position.                     |
| `ValueAt(index)`                                      | Reads a child value by position.                                    |
| `KeyAt(index)`                                        | Reads an object key by position.                                    |

## Typed Accessors

Typed accessors avoid repeated type checks in user code and keep JSON reading compact.

```vb
Debug.Print doc.StringValue("name")
Debug.Print doc.NumberValue("score")
Debug.Print doc.BoolValue("active")
```

### Object Field Accessors

| Method                | Description                                   |
| :-------------------- | :-------------------------------------------- |
| `StringValue(key)`    | Reads a field as `String`.                    |
| `NumberValue(key)`    | Reads a field as `Double`.                    |
| `BoolValue(key)`      | Reads a field as `Boolean`.                   |
| `RawStringValue(key)` | Reads a string field without JSON unescaping. |

### Array Accessors

| Method               | Description                                        |
| :------------------- | :------------------------------------------------- |
| `StringAt(index)`    | Reads an array item as `String`.                   |
| `NumberAt(index)`    | Reads an array item as `Double`.                   |
| `BoolAt(index)`      | Reads an array item as `Boolean`.                  |
| `RawStringAt(index)` | Reads a string array item without JSON unescaping. |

## Token Iteration

For very large arrays or objects, token iteration avoids creating wrapper nodes for every child.

```vb
Dim arr As JSON
Set arr = doc.Node("items")

Dim t As Long
t = arr.FirstChildToken()

Do While t <> 0
    Debug.Print arr.TokenStringValue(t)
    t = arr.NextToken(t)
Loop
```

### Token Helpers

| Method                         | Description                                        |
| :----------------------------- | :------------------------------------------------- |
| `FirstChildToken()`            | Returns the first child token of the current node. |
| `LastChildToken()`             | Returns the last child token of the current node.  |
| `NextToken(tokenId)`           | Returns the next sibling token.                    |
| `NodeFromToken(tokenId)`       | Wraps a token as a JSON node.                      |
| `TokenKey(tokenId)`            | Reads the key of an object child token.            |
| `TokenValue(tokenId)`          | Reads the token value as a Variant.                |
| `TokenStringValue(tokenId)`    | Reads the token value as a String.                 |
| `TokenRawStringValue(tokenId)` | Reads the raw string value without unescaping.     |
| `TokenNumberValue(tokenId)`    | Reads the token value as a Double.                 |
| `TokenBoolValue(tokenId)`      | Reads the token value as a Boolean.                |

## Fast Field Access from Tokens

When iterating arrays of objects, token field helpers let you read fields directly from each object token.

```vb
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

### Token Field Helpers

| Method                         | Description                                              |
| :----------------------------- | :------------------------------------------------------- |
| `TokenString(tokenId, key)`    | Reads a child field as String from an object token.      |
| `TokenRawString(tokenId, key)` | Reads a child string field without unescaping.           |
| `TokenRawField(tokenId, key)`  | Reads a raw JSON field slice.                            |
| `TokenNumber(tokenId, key)`    | Reads a child field as Double from an object token.      |
| `TokenBool(tokenId, key)`      | Reads a child field as Boolean from an object token.     |
| `TokenNode(tokenId, key)`      | Reads a child field as a JSON node from an object token. |

## Raw JSON Access

Raw access is useful when you want to extract a nested object or array as JSON text without walking its full structure.

```vb
Dim raw As String
raw = users.TokenRawField(t, "profile")

Debug.Print raw
```

This is useful for:

* Passing nested JSON to another layer.
* Caching raw fragments.
* Avoiding object allocation for fields you do not need immediately.
* Extracting large nested payloads from API responses.

## Feature Summary

| Category          | Features                                                                                          |
| :---------------- | :------------------------------------------------------------------------------------------------ |
| **Core**          | Single `.cls`, predeclared class API, x86/x64 support                                             |
| **Parsing**       | Zero-copy source slicing, compact token tree, lazy node wrappers                                  |
| **Traversal**     | Object keys, array indexes, child counts, node wrappers                                           |
| **Typed Reads**   | String, Double, Boolean, Null, Variant values                                                     |
| **Raw Reads**     | Raw string values, raw JSON field extraction, token slices                                        |
| **Iteration**     | Token-based traversal for large arrays and objects                                                |
| **Writing**       | Compact Stringify, pretty Stringify, custom indentation                                           |
| **VBA Values**    | Primitive values, arrays, Collections, Dictionaries, JSON nodes                                   |
| **Performance**   | No Dictionary/Collection allocation during parse, native string comparison, native quote scanning |
| **Compatibility** | 32-bit and 64-bit Office through conditional compilation                                          |

## Performance Notes

This module is optimized for fast reads and low allocation overhead.

For best performance:

* Parse once and reuse the document.
* Use typed accessors when reading known fields.
* Use token iteration for huge arrays.
* Use `TokenString`, `TokenNumber`, and `TokenBool` when walking arrays of objects.
* Use raw field access when you only need to forward or cache a nested JSON fragment.
* Avoid repeatedly wrapping every array item with `NodeAt` unless you need object-style traversal.
* Prefer `Stringify(False)` for compact output and `Stringify(True)` only when human-readable output is needed.

## Recommended Patterns

### API Response Parsing

```vb
Public Sub ReadApiResponse(ByVal responseText As String)
    Dim doc As JSON
    Set doc = JSON.Parse(responseText)

    If Not doc.Exists("data") Then Exit Sub

    Dim data As JSON
    Set data = doc.Node("data")

    Debug.Print data.StringValue("name")
    Debug.Print data.NumberValue("id")
End Sub
```

### Large Array Iteration

```vb
Public Sub ReadLargeArray(ByVal responseText As String)
    Dim doc As JSON
    Set doc = JSON.Parse(responseText)

    Dim rows As JSON
    Set rows = doc.Node("rows")

    Dim t As Long
    t = rows.FirstChildToken()

    Do While t <> 0
        Debug.Print rows.TokenString(t, "name"), rows.TokenNumber(t, "score")
        t = rows.NextToken(t)
    Loop
End Sub
```

### Build JSON Output

```vb
Public Sub BuildJson()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    dict("name") = "JSON"
    dict("version") = "1.0.1"
    dict("language") = "VBA"

    Debug.Print JSON.StringifyValue(dict, True)
End Sub
```

## Roadmap

### Current Version (v1.0.1)

* [x] Zero-copy JSON parsing.
* [x] Compact token tree storage.
* [x] Lazy node wrappers.
* [x] Object key access.
* [x] Array index access.
* [x] Typed accessors for strings, numbers, and booleans.
* [x] Null detection.
* [x] Raw string access.
* [x] Raw field extraction.
* [x] Token iteration.
* [x] Token field access.
* [x] Stringify for parsed JSON nodes.
* [x] Stringify for primitive VBA values.
* [x] Stringify for arrays.
* [x] Stringify for Collections.
* [x] Stringify for Dictionaries.
* [x] Pretty printing.
* [x] Custom indentation.
* [x] 32-bit and 64-bit Office compatibility.

### Planned

* [ ] Benchmark suite against common VBA JSON libraries.
* [ ] Optional stricter validation mode.
* [ ] Additional writer helpers for object/array construction.
* [ ] More real-world Excel, PowerPoint, Access, and Word integration examples.

## Documentation

* [**API Reference**](docs/API_REFERENCE.md) – Detailed guide to every public method, property, and traversal pattern.

## License

MIT. Designed for fast JSON parsing, clean traversal, and practical data automation inside Microsoft Office.
