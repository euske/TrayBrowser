TrayBrowser
===========

TrayBrowser is a mini Internet Explorer that can "close to System Tray".

How to Build
------------

 1. Install Windows SDK (v7.1 or newer).
 2. Launch Windows SDL 7.1 Command Prompt.
 3. Type `nmake`.
 4. Install TrayBrowser.reg.
 5. Copy TrayBrowser.exe to somewhere.

**Caution**

Make sure to install TrayBrowser.reg, the registry file,
successfully in your system. This is crucial to make the browser
compatible to the latest version. Without this step, it does not
load certain websites correctly (notably hitbox.tv).

How to Use
----------

TrayBrowser.exe appeared as a SysTray icon. Right-click it to open the menu.
It has the following commands:

 * `Open`: Opens a new URL.
 * `Pin`: Pins the window to the topmost position.
   This also makes the window not appearing on the Alt+Tab list.
 * `Recent`: Opens a recent URL.
 * `Exit`: Quits the program.

TrayBrowser.exe tries to read the TrayBrowser.ini file on the same directory.
The file has bookmark entries in the following format:

    [bookmarks]
    url0 = http://twitch.tv/ffstv/popup
    url1 = http://rtmp.example.com/

The program also saves new URLs to the .ini file when it exits.
