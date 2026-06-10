<div align="center">
  <img src="resources/logo.png" width="160" alt="JSON logo" />
</div>

<h1 align="center">JSON</h1>

<p align="center">
  <b>A fast, single-file JSON parser and lightweight writer for VBA.</b><br/>
  Typed accessors, Keys/Exists helpers, raw field access, token iteration, lazy node wrappers, and Stringify support.
</p>

<p align="center">
  <img src="https://github.com/vbacollective/json/actions/workflows/ci.yml/badge.svg" alt="CI" />
  <img src="https://github.com/vbacollective/json/actions/workflows/release-assets.yml/badge.svg" alt="Release Assets" />
  <img src="https://img.shields.io/github/v/release/vbacollective/json" alt="Latest Version" />
  <img src="https://img.shields.io/badge/language-VBA-867DB1.svg" alt="Language" />
  <img src="https://img.shields.io/badge/platform-Windows-0078D6.svg" alt="Platform" />
  <img src="https://img.shields.io/badge/arch-32%20%26%2064--bit-green.svg" alt="Architecture" />
  <img src="https://img.shields.io/badge/dependencies-none-success.svg" alt="Dependencies" />
  <img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License" />
</p>

**JSON** is a single-file JSON reader and writer for VBA Office projects. Import `package/JSON.cls` into Excel, Word, PowerPoint, Access, or another VBA host and use the predeclared `JSON` class directly.

The parser stores a compact token tree instead of building nested `Dictionary` or `Collection` objects during parse. Child objects and arrays are exposed through lightweight `JSON` node wrappers only when your code asks for them.

This keeps the API simple while still allowing fast reads, typed access, token iteration, and normal VBA-friendly usage.

## Features

* **Single class:** Import only `JSON.cls`.
* **No references:** No required entries in **Tools > References**.
* **No Dictionary dependency:** Objects are stored internally and can be read with `Keys`, `Exists`, and typed key accessors.
* **Typed reads:** Use `StringKey`, `NumberKey`, `BoolKey`, `StringIndex`, `NumberIndex`, and `BoolIndex` when the schema is known.
* **Node traversal:** Use `NodeKey`, `NodeIndex`, `Count`, `ExistsKey`, `ExistsIndex`, and `JsonType`.
* **Compatibility access:** The generic `Item`, `Value`, `ValueAt`, `StringValue`, `NumberValue`, `BoolValue`, `StringAt`, `NumberAt`, and `BoolAt` style remains available for existing code.
* **Keys support:** Use `Keys()` to list object keys or array indexes.
* **Raw access:** Use raw string helpers when you need the original textual value without extra conversion.
* **Token iteration:** Walk large arrays with token handles instead of allocating a wrapper for every item.
* **Stringify:** Serialize parsed JSON, primitive values, arrays, Collections, Dictionaries, and JSON nodes.
* **Pretty output:** Use spaces or a custom indentation string such as `vbTab`.
* **Office compatibility:** Supports 32-bit and 64-bit VBA through conditional declarations.

## Installation

1. Download or clone this repository.
2. Open the VBA editor with `Alt + F11`.
3. Choose **File > Import File...**.
4. Import [package/JSON.cls](package/JSON.cls).
5. Save your Office file as a macro-enabled document such as `.xlsm`, `.pptm`, `.docm`, or `.accdb`.

No external references are required for the JSON class itself. Examples that use `Scripting.Dictionary` create it with late binding.

## Quick Start

```vb
Public Sub ReadJson()
    Dim text As String
    text = "{""name"":""Ueslei"",""age"":18,""active"":true}"

    Dim doc As JSON
    Set doc = JSON.Parse(text)

    Debug.Print doc.StringKey("name")
    Debug.Print doc.NumberKey("age")
    Debug.Print doc.BoolKey("active")
End Sub
```

## Reading Objects

Use object keys for direct field access.

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""project"":""JSON"",""language"":""VBA""}")

Debug.Print doc.StringKey("project")
Debug.Print doc.StringKey("language")
Debug.Print doc.ExistsKey("project")
```

Nested objects and arrays are exposed as lightweight `JSON` nodes.

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""user"":{""name"":""Ueslei"",""role"":""developer""}}")

Dim user As JSON
Set user = doc.NodeKey("user")

If Not user Is Nothing Then
    Debug.Print user.StringKey("name")
    Debug.Print user.StringKey("role")
End If
```

## Reading Arrays

Array access is zero-based.

```vb
Dim items As JSON
Set items = JSON.Parse("[""Excel"",""Access"",""Word""]")

Dim i As Long
For i = 0 To items.Count - 1
    Debug.Print items.StringIndex(i)
Next i
```

Nested arrays and objects can be accessed with `NodeIndex`.

```vb
Dim rows As JSON
Set rows = JSON.Parse("[{""name"":""Ana""},{""name"":""Bia""}]")

Dim first As JSON
Set first = rows.NodeIndex(0)

Debug.Print first.StringKey("name")
```

## Keys and Exists

`Keys()` gives you a simple way to inspect object fields or array positions without using `Scripting.Dictionary`.

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""name"":""Ueslei"",""age"":18,""active"":true}")

Dim keys As Variant
keys = doc.Keys

Dim i As Long
For i = LBound(keys) To UBound(keys)
    Debug.Print keys(i)
Next i
```

For objects, `Keys()` returns field names.

```vb
Debug.Print doc.ExistsKey("name")
Debug.Print doc.ExistsKey("missing")
```

For arrays, `Keys()` returns indexes.

```vb
Dim arr As JSON
Set arr = JSON.Parse("[10,20,30]")

Debug.Print arr.ExistsIndex(0)
Debug.Print arr.ExistsIndex(3)
```

The generic `Exists` helper is also available when you want a single API for both keys and indexes.

```vb
Debug.Print doc.Exists("name")
Debug.Print arr.Exists(0)
```

## Default Member Chaining

Because `Item` is the default member, object keys and array indexes can be chained naturally.

```vb
Dim myJson As JSON
Set myJson = JSON.Parse("{""names"":[""Ana"",""Bia"",""Caio""]}")

Debug.Print myJson("names")(0)
Debug.Print myJson("names")(1)
```

This style is useful for short reads from known JSON shapes. For hot loops or large payloads, prefer typed accessors such as `StringKey`, `NumberKey`, `BoolKey`, `NodeKey`, and `NodeIndex`.

## Large Arrays

For large arrays of objects, use token iteration to avoid creating a node wrapper for each row.

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

## Writing JSON

Serialize a parsed document or node with `Stringify`.

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""name"":""JSON"",""tags"":[""VBA"",""parser""]}")

Debug.Print doc.Stringify()
Debug.Print doc.Stringify(True)
Debug.Print doc.StringifyWithIndent(True, vbTab)
```

Serialize normal VBA values with `StringifyValue`.

```vb
Dim data As Object
Set data = CreateObject("Scripting.Dictionary")

data("name") = "JSON"
data("version") = "1.0.1"
data("fast") = True

Debug.Print JSON.StringifyValue(data, True)
```

## Public API Summary

Core methods and properties:

* `Parse(text)`: Parses JSON text and returns a `JSON` document.
* `Stringify(pretty, indentSize)`: Serializes the current document or node.
* `StringifyWithIndent(pretty, indentText)`: Serializes with custom indentation text.
* `StringifyValue(value, pretty, indentSize)`: Serializes an external VBA value.
* `StringifyValueWithIndent(value, pretty, indentText)`: Serializes an external VBA value with custom indentation text.
* `Item(key)`: Reads a child value by key or index through the default member.
* `Value`: Reads the current node value.
* `Count`: Returns direct child count.
* `JsonType`: Returns `object`, `array`, `string`, `number`, `boolean`, `null`, or an empty string.
* `Keys()`: Returns object keys or array indexes.
* `Exists(key)`: Checks whether an object key or array index exists.
* `ExistsKey(key)`: Checks whether an object key exists.
* `ExistsIndex(index)`: Checks whether an array index exists.
* `NodeKey(key)`: Returns an object or array child node by key.
* `NodeIndex(index)`: Returns an object or array child node by index.
* `Node(key)`: Compatibility helper for child node access by key.
* `NodeAt(index)`: Compatibility helper for child node access by index.
* `ValueAt(index)`: Reads any child value by position.
* `KeyAt(index)`: Reads an object child key by position.

Typed key accessors:

* `StringKey(key)`
* `NumberKey(key)`
* `BoolKey(key)`
* `RawStringKey(key)`

Typed index accessors:

* `StringIndex(index)`
* `NumberIndex(index)`
* `BoolIndex(index)`
* `RawStringIndex(index)`

Compatibility typed accessors:

* `StringValue(key)`, `NumberValue(key)`, `BoolValue(key)`, `RawStringValue(key)`.
* `StringAt(index)`, `NumberAt(index)`, `BoolAt(index)`, `RawStringAt(index)`.

Token helpers:

* `FirstChildToken()`, `LastChildToken()`, `NextToken(tokenId)`, `NodeFromToken(tokenId)`.
* `TokenKey(tokenId)`, `TokenValue(tokenId)`, `TokenStringValue(tokenId)`, `TokenRawStringValue(tokenId)`, `TokenNumberValue(tokenId)`, `TokenBoolValue(tokenId)`.
* `TokenString(tokenId, key)`, `TokenRawString(tokenId, key)`, `TokenRawField(tokenId, key)`, `TokenNumber(tokenId, key)`, `TokenBool(tokenId, key)`, `TokenNode(tokenId, key)`.

## Validation Tests

The project includes a validation module for checking the parser against normal JSON usage.

The safe validation suite currently covers:

* valid objects;
* valid arrays;
* primitive roots;
* escaped strings;
* unicode escapes;
* numbers;
* `Stringify` and reparse;
* `StringifyValue`;
* `Keys` and `Exists`;
* token iteration.

Current safe test result:

```text
Total:   75
Passed:  75
Failed:  0
```

Invalid JSON tests should be run separately after strict parser guards are enabled, because malformed inputs can expose parser progress issues in VBA hosts.

## Practical Guidance

* Keep the root parsed `JSON` document alive while using child nodes returned by `NodeKey`, `NodeIndex`, `Node`, `NodeAt`, `TokenNode`, or `NodeFromToken`.
* Use typed accessors when the payload schema is known.
* Use `ExistsKey` and `ExistsIndex` when a missing value must be distinguished from `""`, `0`, or `False`.
* Use generic `Exists` when writing code that may receive either object keys or array indexes.
* Use `Keys()` when you need to enumerate object fields or array positions.
* Use token iteration for large arrays and high-volume object loops.
* Use raw field access when forwarding or caching nested JSON without fully traversing it.
* Use compact `Stringify(False)` for storage and transport, and pretty `Stringify(True)` for debugging or readable output.

## Examples

The [examples](examples) directory contains importable `.bas` modules:

* [BasicRead.bas](examples/BasicRead.bas) shows parsing, object fields, nested nodes, and arrays.
* [TokenIteration.bas](examples/TokenIteration.bas) shows fast traversal over arrays of objects.
* [StringifyValues.bas](examples/StringifyValues.bas) shows writing JSON from Dictionaries, Collections, arrays, and parsed nodes.

## Documentation

* [API Reference](docs/API_REFERENCE.md) provides method-level documentation and recipes.
* [Architecture](docs/ARCHITECTURE.md) explains the token tree, parser pipeline, and writer pipeline.

## License

MIT. Designed for fast JSON parsing, clean traversal, low allocation, and practical data automation inside Microsoft Office.
