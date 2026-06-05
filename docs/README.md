# Documentation

Welcome to the JSON technical documentation. This directory contains detailed guides on how to integrate, use, understand, and maintain the JSON parser and writer.

## Resource Index

| File | Audience | Description |
|:---|:---|:---|
| [**API Reference**](API_REFERENCE.md) | Developers | Exhaustive guide to every public method, property, accessor, token helper, and serialization function. |
| [**Architecture**](ARCHITECTURE.md) | Advanced | Internal implementation details: zero-copy parsing, SAFEARRAY string aliasing, token tree storage, lazy node wrappers, and Stringify pipeline. |

## Quick Integration Pattern

JSON is designed for simple, synchronous usage in standard VBA modules.

```vb
' Basic JSON Parsing Procedure
Public Sub RunJsonTest()
    Dim text As String
    text = "{""name"":""Ryan"",""age"":18,""active"":true}"

    Dim doc As JSON
    Set doc = JSON.Parse(text)

    Debug.Print doc.StringValue("name")
    Debug.Print doc.NumberValue("age")
    Debug.Print doc.BoolValue("active")
End Sub
```

## Stringify Pattern

Use `StringifyValue` to serialize normal VBA values such as Dictionaries, Collections, arrays, strings, numbers, and booleans.

```vb
' Basic JSON Writing Procedure
Public Sub RunStringifyTest()
    Dim data As Object
    Set data = CreateObject("Scripting.Dictionary")

    data("name") = "JSON"
    data("language") = "VBA"
    data("fast") = True

    Debug.Print JSON.StringifyValue(data, True)
End Sub
```

## Large Array Pattern

For large arrays of objects, prefer token iteration instead of creating a node wrapper for every item.

```vb
' Fast Array Traversal Procedure
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

        t = rows.NextToken(t)
    Loop
End Sub
```

## Lifecycle Management

JSON does not require explicit startup or shutdown. Parsed documents are normal VBA objects and are cleaned up automatically when released.

```vb
Public Sub ParseAndRelease()
    Dim doc As JSON
    Set doc = JSON.Parse("{""ok"":true}")

    Debug.Print doc.BoolValue("ok")

    Set doc = Nothing
End Sub
```

> [!IMPORTANT]
> Child nodes returned by `Node`, `NodeAt`, `TokenNode`, or `NodeFromToken` depend on the root parsed document. Keep the root `JSON` document alive while using any child node wrappers.

```vb
Public Sub KeepRootAlive()
    Dim doc As JSON
    Set doc = JSON.Parse("{""user"":{""name"":""Ueslei""}}")

    Dim user As JSON
    Set user = doc.Node("user")

    Debug.Print user.StringValue("name")
End Sub
```
