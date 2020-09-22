<div align="center">

## Low\-Level KeyLogger


</div>

### Description

This application uses the SetWindowsHookEx function to retrieve all keyboard events. You can configurate the keys you want to receive by event callback function. Uses no timer + GetAsyncKeyState to fetch pressed keys. The user control sends the Chr-compatible byte code for the pressed character (if any), the name for the key/special key (if known), the application the key was typed in and a flag, indicating whether the application the user was typing in was changed. You may use this code as keylogger or modify it to prevent users from using windows default key combinations like Ctrl + tab etc. Plz vote if you like my code. My first submission on psc ;)
 
### More Info
 
Some knowledge about windows hooks and subclassing.

Ensure to call <UserControl>.DisableLogging before shutting down application. Hook functions may cause application crashes if not unhook cleanly.


<span>             |<span>
---                |---
**Submitted On**   |2004-06-29 22:36:02
**By**             |[aYa\.net](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aya-net.md)
**Level**          |Advanced
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Low\-Level\_1763896292004\.zip](https://github.com/Planet-Source-Code/aya-net-low-level-keylogger__1-54663/archive/master.zip)








