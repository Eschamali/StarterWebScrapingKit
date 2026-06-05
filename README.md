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
  <img src="https://github.com/vbacollective/json/actions/workflows/release-assets.yml/badge.svg" alt="Release Assets" />
  <img src="https://img.shields.io/badge/version-v1.0.1-blue.svg" alt="Version" />
  <img src="https://img.shields.io/badge/language-VBA-867DB1.svg" alt="Language" />
  <img src="https://img.shields.io/badge/platform-Windows-0078D6.svg" alt="Platform" />
  <img src="https://img.shields.io/badge/arch-32%20%26%2064--bit-green.svg" alt="Architecture" />
  <img src="https://img.shields.io/badge/parser-zero--copy-success.svg" alt="Zero-copy parser" />
  <img src="https://img.shields.io/badge/dependencies-none-success.svg" alt="Dependencies" />
  <img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License" />
</p>

**JSON** is a single-file JSON reader and writer for VBA Office projects. Import `package/JSON.cls` into Excel, Word, PowerPoint, Access, or another VBA host and use the predeclared `JSON` class directly.

The parser stores a compact token tree over the original JSON text. It avoids building nested Dictionaries, Collections, or wrapper objects during parse, then creates lightweight node wrappers only when your code asks for them.

## Features

* **Single class:** Import only `JSON.cls`.
* **No references:** Parsing, traversal, and writing do not require entries in **Tools > References**.
* **Zero-copy parser:** Keys and values are stored as slices of the source text until requested.
* **Typed reads:** Use `StringValue`, `NumberValue`, `BoolValue`, `StringAt`, `NumberAt`, and `BoolAt`.
* **Node traversal:** Use `Node`, `NodeAt`, `ValueAt`, `KeyAt`, `Count`, `Exists`, and `JsonType`.
* **Raw access:** Use `RawStringValue`, `RawStringAt`, `TokenRawString`, and `TokenRawField`.
* **Token iteration:** Walk large arrays with token handles instead of allocating a wrapper for every item.
* **Stringify:** Serialize parsed JSON, primitive values, arrays, Collections, Dictionaries, and JSON nodes.
* **Pretty output:** Use spaces or a custom indentation string such as `vbTab`.
* **Office compatibility:** Supports 32-bit and 64-bit VBA through conditional declarations.

## Repository Layout

* [package/JSON.cls](package/JSON.cls) contains the production class.
* [package/README.md](package/README.md) explains packaging and import steps.
* [docs/README.md](docs/README.md) links the technical documentation.
* [docs/API_REFERENCE.md](docs/API_REFERENCE.md) documents the public API.
* [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) explains the implementation model.
* [examples/README.md](examples/README.md) explains the runnable sample modules.
* [resources/README.md](resources/README.md) describes repository assets.

## Installation

1. Download or clone this repository.
2. Open the VBA editor with `Alt + F11`.
3. Choose **File > Import File...**.
4. Import [package/JSON.cls](package/JSON.cls).
5. Save your Office file as a macro-enabled document such as `.xlsm`, `.pptm`, `.docm`, or `.accdb`.

No external references are required for the JSON class itself. The examples that use `Scripting.Dictionary` create it with late binding.

## Quick Start

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

## Reading Objects

Use object keys for direct field access.

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""project"":""JSON"",""language"":""VBA""}")

Debug.Print doc.StringValue("project")
Debug.Print doc.StringValue("language")
Debug.Print doc.Exists("project")
```

Nested objects and arrays are exposed as lightweight `JSON` nodes.

```vb
Dim doc As JSON
Set doc = JSON.Parse("{""user"":{""name"":""Ueslei"",""role"":""developer""}}")

Dim user As JSON
Set user = doc.Node("user")

If Not user Is Nothing Then
    Debug.Print user.StringValue("name")
    Debug.Print user.StringValue("role")
End If
```

## Reading Arrays

Array access is zero-based.

```vb
Dim items As JSON
Set items = JSON.Parse("[""Excel"",""Access"",""Word""]")

Dim i As Long
For i = 0 To items.Count - 1
    Debug.Print items.StringAt(i)
Next i
```

## Large Arrays

For large arrays of objects, use token iteration to avoid creating a node wrapper for each row.

```vb
Public Sub ReadRows(ByVal responseText As String)
    Dim doc As JSON
    Set doc = JSON.Parse(responseText)

    Dim rows As JSON
    Set rows = doc.Node("rows")

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

* `Parse(text)`: Parses JSON text into a document.
* `Stringify(pretty, indentSize)`: Serializes the current document or node.
* `StringifyWithIndent(pretty, indentText)`: Serializes with custom indentation text.
* `StringifyValue(value, pretty, indentSize)`: Serializes an external VBA value.
* `StringifyValueWithIndent(value, pretty, indentText)`: Serializes an external VBA value with custom indentation text.
* `Item(key)`: Reads a child value by key or index.
* `Value`: Reads the current node value.
* `Count`: Returns direct child count.
* `JsonType`: Returns `object`, `array`, `string`, `number`, `boolean`, `null`, or an empty string.
* `Exists(key)`: Checks whether an object key or array index exists.
* `Node(key)`: Returns a child object or array node.
* `NodeAt(index)`: Returns an object or array child node by position.
* `ValueAt(index)`: Reads any child value by position.
* `KeyAt(index)`: Reads an object child key by position.

Typed accessors:

* `StringValue(key)`, `NumberValue(key)`, `BoolValue(key)`, `RawStringValue(key)`.
* `StringAt(index)`, `NumberAt(index)`, `BoolAt(index)`, `RawStringAt(index)`.

Token helpers:

* `FirstChildToken()`, `LastChildToken()`, `NextToken(tokenId)`, `NodeFromToken(tokenId)`.
* `TokenKey(tokenId)`, `TokenValue(tokenId)`, `TokenStringValue(tokenId)`, `TokenRawStringValue(tokenId)`, `TokenNumberValue(tokenId)`, `TokenBoolValue(tokenId)`.
* `TokenString(tokenId, key)`, `TokenRawString(tokenId, key)`, `TokenRawField(tokenId, key)`, `TokenNumber(tokenId, key)`, `TokenBool(tokenId, key)`, `TokenNode(tokenId, key)`.

## Practical Guidance

* Keep the root parsed `JSON` document alive while using child nodes returned by `Node`, `NodeAt`, `TokenNode`, or `NodeFromToken`.
* Use typed accessors when the payload schema is known.
* Use `Exists` when a missing value must be distinguished from `""`, `0`, or `False`.
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
* [Architecture](docs/ARCHITECTURE.md) explains the token tree, SAFEARRAY aliasing, parser pipeline, and writer pipeline.

## License

MIT. Designed for fast JSON parsing, clean traversal, low allocation, and practical data automation inside Microsoft Office.
