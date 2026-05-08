# Excel VBA Web Automation Starter Kit

![Logo](doc/Logo.png)

> [!IMPORTANT]
> This text was translated into English by AI based on "README-jp.md".

[🇯🇵日本語のREADMEはこちら](README-jp.md)

![Intro Image](doc/Top.png)

## The World of the Internet, in Your Hands

All the essential elements for web scraping have been packed into this **single** macro workbook.  
No more tedious environment setup. From the moment you open this workbook, your journey toward business efficiency and automated internet operations begins.

This tool implements the "Three Sacred Treasures" required to conquer modern web technologies:

1. **🚀 REST WebAPI (WinHTTP 5.1)**
    * The standard for high-speed, lightweight data collection. A robust implementation that works solely with reference settings.
2. **🤖 Browser Automation (CDP via Pipe & WebDriver BiDi)**
    * Freely control Chromium-based browsers (Edge/Chrome). A modern implementation using pipe communication that doesn't require external drivers (.exe).
3. **⚡ WebSocket Communication**
    * A challenge for real-time communication. Equipped with minimal connection and send/receive functions using WinAPI. An evolving feature that pushes the boundaries of VBA.

---

## 🔥【Strengths of This Tool】🔥

* **Ultimate Portable Browser Support (Freedom from Driver Version Management!)**
  * You will never encounter the "Browser and WebDriver version mismatch error" that plagues Selenium users!
  * Whether it's a modified browser, an anti-detect browser, or a portable Chrome on a USB drive, complete automation is possible in an instant just by **"pasting the exe path into a cell on the settings sheet"** 😎

* **Infinite Extensibility: Your Own Custom Tool!**
  * Simply tell an AI about the "[Template](https://github.com/Eschamali/StarterWebScrapingKit/tree/dev/ForDevelopers/TemplateExtensions)" and the "function you want," and complex automation code will be completed in seconds!
  * No need to memorize tedious CDP specifications. Anyone can easily extend functionality as long as they have an idea.
  * Depending on how you craft your prompts, you can even generate "demo code" with detailed explanations fully automatically!

* **🚀 A "New Standard" Architecture for VBA, on par with Playwright / Puppeteer**
  * Directly access the heart of the browser without leaving the "footprints" of WebDriver. This tool occupies a **low-layer position equivalent to Playwright / Puppeteer**, even though it's VBA.
  * Its greatest strength lies in its "cleanliness." By using a "pure" operation style that doesn't inject any unique JS variables or patches (which often trigger detection), it has achieved **stealthiness that makes it easier to bypass modern defenses like Cloudflare** as if you had a VIP pass.  
  * *Note: While bypass is not guaranteed, we have confirmed cases where it is easier to bypass than SeleniumVBA.*

---

## 🌈 Two Main Routes to Choose From

This project offers two branches (implementation methods) depending on your needs.

### 1. Main Branch (Edge x Pipe x CDP)
**"Control the browser freely from the outside"** 
- Controls standard Edge/Chrome via pipe communication.
- You can use existing browser profiles (favorites, login states, etc.) as they are.
- This is the mainstream method with high stability and easy debugging.

### 2. WebView2 Branch (UserForm x Native) [Under Development]
**"Incorporate the browser into Excel"**
- Directly embeds WebView2 within an Excel UserForm.
- The ultimate UI experience for those who **"really, really want the feeling of the browser moving as one with Excel."**
- Allows you to execute scraping natively from buttons on the UserForm.

---

## ⭐️ New Feature: Full WebDriver BiDi Support! (A VBA First🦊)

In addition to traditional CDP (Chrome DevTools Protocol) operations, we have quickly implemented support for **"WebDriver BiDi"** (`WebDriverBiDiCore.cls`), the next-generation protocol currently being established as a global W3C standard.

While maintaining the project's philosophy of being **"self-contained in VBA"** without using external `chromedriver.exe` or middleware like Selenium, the following advanced operations are now possible:

*   📥 **Perfect Subscription to Asynchronous Events** (Real-time detection of page load completion or console errors)
*   ⚠️ **Fine-grained Control of JavaScript Alert Dialogs** (Implementation of fallbacks that prevent VBA from freezing)
*   🔌 **CDP Tunneling via BiDi+** (Flexibility to cover areas where standard features are insufficient)

**📖 For detailed technical documentation and usage, please visit the official documentation (GitHub Pages).**
*   ➡️ **[Official Documentation Top (Usage & Technical Architecture)](https://eschamali.github.io/StarterWebScrapingKit/)**

---

### 【Credits & Acknowledgments】

This tool is a mashup of many wonderful libraries shared by VBA artisans around the world, integrated into a form that is easy to use in practice.  
I express my heartfelt respect and gratitude for the wisdom and code of my great predecessors.

* **Core Logic for WebSocket Implementation**
  * [ChromeControler-No-Selenium-WebDriver-VBAJSON](https://github.com/24000/ChromeControler-No-Selenium-WebDriver-VBAJSON)
    * Author: [@kabkabkab](https://qiita.com/kabkabkab/items/9952a796ee9244fc98ad)
* **Foundation for CDP Control and Pipe Communication**
  * [Chromium-Automation-with-CDP-for-VBA](https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA)
    * Author: longvh211
* **WinHTTP 5.1 Wrapper**
  * [VBA-Web](https://github.com/VBA-Tools-v2/VBA-Web)
    * Original Author: Tim Hall
* **High-performance JSON Parser**
  * [WebJsonConverter.cls (from SeleniumVBA)](https://github.com/GCuser99/SeleniumVBA/blob/main/src/VBA/WebJsonConverter.cls)
    * Improved by GCuser99
    * Replaced the existing JsonConverter with this for better maintainability
* **Fast Character Code Conversion Wrapper**
  * [How to convert VBA/VB6 Unicode strings to UTF-8](https://di-mgt.com.au/howto-convert-vba-unicode-to-utf8.html)
    * David Ireland, DI Management Services Pty
  * [VBAで Windows APIを使った UTF-8 ←→ Unicode相互変換](https://qiita.com/yamashiroakihito/items/9b609653fef6fa8a5ab2)
    * Author: @yamashiroakihito
* **Log Level Basics**
  * [VBA-Log](https://github.com/VBA-tools/VBA-Log)
    * Author: timhall
* **Core Logic for BiDi in Chromium Browsers**
  * [chromium-bidi](https://github.com/GoogleChromeLabs/chromium-bidi)
    * Author: GoogleChromeLabs Team
* **The Amazing Person Who Embedded WebView2 in UserForm Without Extra Downloads**
  * [WebView2-For-Excel-VBA](https://github.com/tarboh/WebView2-For-Excel-VBA)
    * Author: [Tarboh](https://x.com/fenblen_puyo)

*Note: For detailed usage and methods of each function, please refer to the documentation of the original libraries above.*

## 💡 Introduction: About the "Protected View" Displayed When Opening Downloaded Files

![Excel Protected View](doc/FirstStep1.png)

When you open a downloaded macro workbook, a yellow bar saying **"Protected View"** may appear at the top of Excel, and you might need to click the "Enable Editing" button.  
Furthermore, a security warning might appear when you try to run the macro.  
![Security Risk](doc/FirstStep2.png)

This is a normal and very smart behavior where **your PC is trying to protect you from "unknown" files coming from the internet.**

### How to Unblock

1. Close all Excel instances.
2. Right-click the downloaded Excel file and select **Properties**.  
![Right-click Menu](doc/FirstStep4.png)
3. Check the **Unblock** checkbox and click the **OK** button.  
![Properties Window](doc/FirstStep5.png)
4. Open the tool again and click the "Enable Editing" button.

To help you use this macro workbook safely and to its full potential, let me briefly explain **"why this extra step is necessary."**

### Why is This "Extra Step" Necessary? 【A Story】

Once upon a time, the internet was a much more peaceful place.  
However, at some point, malicious **"viruses" pretending to be Excel macros** began to spread worldwide.  
People suffered tragedies where their PCs were taken over just by opening an ordinary Excel file attached to an email.

So, Microsoft made a **big decision**.

**"Let's treat all files coming from the internet as 'suspicious characters of unknown origin'!"**

#### The "Mark of the Web (MOTW)" Stamp

The moment you download a file from the internet (web browser, email client, etc.), Windows places a special **"stamp"** called **`Mark of the Web` (MOTW)** on the **"invisible" part** of the file, marking it as **"this person is a person of interest from the lawless land called the internet."**

When Excel opens a file, it first checks if this "stamp" exists.  
If it finds the stamp, it judges:

**"Wait! This guy is of unknown character!**  
**It's too dangerous to let them move around freely right away.**  
**First, let's put them in an 'isolation room' called 'Protected View.' And never run their macros!"**

### What It Means for You to Check the "Unblock" Box

![Bottom of Properties Window](doc/FirstStep3.png)  
The only **official "guarantee of identity" procedure** to safely lift this strict security system.  
That is the act of opening the file properties and checking the **"Unblock"** box.

This is the same as you **declaring** to Windows:
**"I know, I know. I know this guy came from the internet.**
**But I (you) personally take responsibility and guarantee this person's 'identity'!**
**So, stop treating them as suspicious and welcome them as a formal 'citizen' of this PC."**

Once this "guarantee of identity" is provided, Windows **permanently removes the `MOTW` "stamp"** from the file.  
As a result, Excel recognizes the file as a "trusted and safe file" and allows the macros to run normally without displaying "Protected View."

---
**This macro workbook is safe.**  
**Please, with your power as a "guarantor," give this file the "permission" to be active on your PC.**

---

##  Advanced Features and Technical Details (Migrated to GitHub Pages)

Extensive documentation such as "Unique Improvements (Japanese UTF-8 support, BrowserEvents property, etc.)," "API Specification Reference," and "Deep Mechanisms and Design Philosophy" has been moved and organized into a **beautiful static site (GitHub Pages)**.

Please visit the **[Official Documentation Site (Features / API Reference)](https://Eschamali.github.io/StarterWebScrapingKit/)** and touch the abyss of browser control that exceeds the limits of VBA!

## Worksheet: Browser Startup Settings

![Worksheet: Browser Startup Settings](doc/説明1.png)

Basic explanations are provided on the worksheet. Here, we explain the startup arguments.

### Meaning of Initial Additional Startup Arguments

To eliminate troublesome elements during automation, we provide several initial arguments along with W3C-compliant arguments.

| Argument Name | Meaning | 
| ----------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | 
| no-first-run | Starts Chromium-based browsers without the first-run setup screens.<br>Skips the "Welcome" screen or prompts to log in to Google/Microsoft accounts. | 
| disable-fre | Same as `no-first-run`. Used together because `no-first-run` alone may not fully suppress it depending on the version or environment. | 
| disable-popup-blocking | Disables popup blocking. | 
| disable-sync | Disables automatic login and synchronization with accounts. | 
| disable-background-networking | Disables several subsystems that perform network requests in the background.<br>Eliminates communications other than the intended one as much as possible. | 
| disable-default-apps | Disables the installation of default apps on first run. | 
| no-service-autorun | Suppresses the startup of extra background services. | 
| enable-automation | Enables the display indicating that the browser is controlled by automation.<br>Serves as a marker to prevent mixing with normal browsers. | 
| test-type=ExcelVBA | Specifies the type of test harness. In short, it's just for decoration. | 

### About Bot Detection Bypass Mode

Adds `disable-blink-features=AutomationControlled` to the startup arguments. This overrides `navigator.webdriver` to `false`, enabling bot detection bypass.  
Some sites check this flag to block access, so turn it ON as needed.

However, keep in mind that this argument is not officially supported and may stop working someday.  
As of the time of writing, a warning message appears, but it still works.  
![Message at the top of browser startup](doc/説明3.png)

### Startup Arguments Within VBA

Contains minimum mandatory arguments for browser automation. You can find these arguments around line 350 of the `CDPBrowser` class module.

| Argument Name | Meaning | 
| --------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | 
| remote-debugging-pipe | Makes the browser allow debugging from a "different process (Excel)" than the "main process."<br>Uses pipe communication. Although it says "remote," it is specified to be accessible only from within the same PC. | 
| user-data-dir | Specifies the full path to the browser's data directory (Cookies, extensions, password vault, etc.).<br>Normally it is `C:\Users\%USERNAME%\AppData\Local\Microsoft\Edge\User Data`, but due to [measures against Cookie theft using debugging features](https://developer.chrome.com/blog/remote-debugging-port?hl=ja) it is now mandatory to specify a folder path other than `User Data`.<br>By default, this tool creates a path in the same hierarchy as `User Data` as `Automation Data`. | 
| homepage | Specifies the initial URL when the browser starts, but it is set to `about:blank` to suppress extra communication.<br>However, if an arbitrary URL is passed to the `app` in the next item, this will not be added. | 
| app | Corresponds to the 2nd argument of the `start` method. If you want to specify the initial URL when starting the browser, you specify it here.<br>Starting with a URL here allows you to prevent user actions that interfere with automation to some extent, such as:<br>・Changing to an arbitrary URL not allowed<br>・Creating tabs not allowed<br><br>It's like a simple kiosk mode. | 
| KioskMode | Corresponds to the 6th argument of the `start` method. Please use this when embedding Edge in a UserForm. `fullscreen` recommended. See [here](https://learn.microsoft.com/en-us/deployedge/microsoft-edge-configure-kiosk-mode) for details. | 

## 🚀 No more WebDriver.exe

**"The simple invocation spell from the IE days, once again."**

Once, we controlled the world with just three lines of code.

```bas
Set ie = CreateObject("InternetExplorer.Application")
ie.Visible = True
ie.Navigate "URL"
```

To all VBAers who are being crushed by the weight of driver version management and environment setup after the disappearance of IE.  
This tool does not give up on the romance of a **"single Excel file"** and brings back the omnipotence of those days to the modern era by directly hitting the CDP.

The basic startup template is as follows.  
The browser will start with the settings defined in the **Worksheet: Browser Startup Settings**, so we recommend this template code unless you have specific requirements.  
In that case, your automation journey begins with just one line.

### In Case of CDP Control

```bas
Public Function StartCDPFromSettingsSheet(Optional StartURL As String, Optional SwitchUser As String, Optional KioskMode As edgeKioskType) As CDPBrowser
    ' Get and apply settings from each cell of the settings sheet
    With ShSetting01_StartBrowser
        ' Setting the type of browser to start
        ' Since it is operated by CDP-Json commands, I think it can be used for things other than Edge and Chrome if it is Chromium-based, but for now, only the major ones.
        Dim BrowserName As String: BrowserName = IIf(.Range(.UseRangeName(4, "Demo_CDP.StartCDPFromSettingsSheet")).Value, "chrome", "edge")

        ' If the 2nd argument is omitted, apply the settings from the sheet side
        Dim UseDataDir As String: UseDataDir = IIf(StrPtr(SwitchUser) = 0, .Range(.UseRangeName(2, "Demo_CDP.StartCDPFromSettingsSheet")).Value, SwitchUser)

        ' Start browser
        Set StartCDPFromSettingsSheet = New CDPBrowser
        StartCDPFromSettingsSheet.start BrowserName, StartURL, .Range(.UseRangeName(6, "Demo_CDP.StartCDPFromSettingsSheet")).Value, UseDataDir, .Range(.UseRangeName(3, "Demo_CDP.StartCDPFromSettingsSheet")).Value, KioskMode
    End With
End Function

Sub BeginningOfAdventureByCDP()
    ' Launch browser based on settings sheet
    Dim HelloWorldAutomationBrowser As CDPBrowser: Set HelloWorldAutomationBrowser = StartCDPFromSettingsSheet

    ' ↓ From here, turn your image into code ↓




    ' Close the browser normally
    HelloWorldAutomationBrowser.quit
End Sub
```

### In Case of BiDi Control

```bas
Public Function StartBiDiFromSettingsSheet(Optional StartURL As String, Optional SwitchUser As String, Optional KioskMode As edgeKioskType, Optional sessionCapabilitiesRequest As Dictionary) As WebDriverBiDiCore
    ' Get and apply settings from each cell of the settings sheet
    With ShSetting01_StartBrowser
        ' Setting the type of browser to start
        ' Since it is operated by BiDi-Json commands specialized for Chromium, I think it can be used for things other than Edge and Chrome if it is Chromium-based, but for now, only the major ones.
        Dim BrowserName As String: BrowserName = IIf(.Range(.UseRangeName(4, "Demo_WebDriverBiDi.StartBiDiFromSettingsSheet")).Value, "chrome", "edge")

        ' If the 2nd argument is omitted, apply the settings from the sheet side
        Dim UseDataDir As String: UseDataDir = IIf(StrPtr(SwitchUser) = 0, .Range(.UseRangeName(2, "Demo_WebDriverBiDi.StartBiDiFromSettingsSheet")).Value, SwitchUser)

        ' Start browser
        Set StartBiDiFromSettingsSheet = New WebDriverBiDiCore
        StartBiDiFromSettingsSheet.start BrowserName, StartURL, .Range(.UseRangeName(6, "Demo_WebDriverBiDi.StartBiDiFromSettingsSheet")).Value, UseDataDir, .Range(.UseRangeName(3, "Demo_WebDriverBiDi.StartBiDiFromSettingsSheet")).Value, KioskMode, sessionCapabilitiesRequest
    End With
End Function

Sub BeginningOfAdventureByBiDi()
    ' Launch browser based on settings sheet
    Dim HelloWorldAutomationBrowser As WebDriverBiDiCore: Set HelloWorldAutomationBrowser = StartBiDiFromSettingsSheet

    ' ↓ From here, turn your image into code ↓




    ' Close the browser normally
    HelloWorldAutomationBrowser.quit
End Sub
```
