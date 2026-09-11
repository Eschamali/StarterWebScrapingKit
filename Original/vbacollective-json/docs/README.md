# Documentation

This directory contains the technical documentation for JSON.

## Guides

* [API Reference](API_REFERENCE.md): Public methods, properties, typed accessors, token helpers, serialization methods, recipes, and troubleshooting.
* [Architecture](ARCHITECTURE.md): Internal parser and writer design, including token storage, SAFEARRAY string aliasing, lazy nodes, raw slices, and compatibility notes.

## Quick Integration Pattern

```vb
Public Sub RunJsonTest()
    Dim text As String
    text = "{""name"":""Ueslei"",""age"":18,""active"":true}"

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

## Lifecycle

Parsed documents are normal VBA objects and are cleaned up automatically when released. Keep the root parsed document alive while using child nodes returned by `Node`, `NodeAt`, `TokenNode`, or `NodeFromToken`.
