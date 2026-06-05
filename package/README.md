# Package

This directory contains the production-ready JSON class. This is the only file you need to integrate JSON parsing and writing into your own VBA project.

## Module Index

| File | Purpose |
|:---|:---|
| [**JSON.cls**](JSON.cls) | The complete, standalone JSON parser and writer. Includes zero-copy parsing, typed accessors, token iteration, raw field access, and Stringify support. |

## Installation

Integrating JSON into a new or existing Office project takes only a few seconds:

1. **Download:** Get the latest `JSON.cls` from the [releases](https://github.com/vbacollective/json/releases) page.
2. **Import:**
   - Open your Excel, PowerPoint, Word, or Access file.
   - Press `Alt + F11` to open the VBA Editor.
   - Go to **File > Import File...** (or press `Ctrl + M`).
   - Select `JSON.cls`.
3. **Save:** Save your document as a macro-enabled file (e.g., `.xlsm`, `.pptm`, `.docm`, `.accdb`).

## Configuration & Dependencies

- **No References:** JSON does not require any entries in **Tools > References** for parsing, traversal, or basic serialization.
- **No DLLs:** The class is fully self-contained and does not require external native binaries, ActiveX controls, installers, or registered components.
- **Single Class:** The entire parser and writer live inside `JSON.cls`.
- **Architecture:** JSON is designed to work in both 32-bit and 64-bit Office through conditional compilation where needed.
