<div align="center">

## Screensaver Wallpaper


</div>

### Description

These are API calls that will allow you to change the wallpaper (instant update) and turn on the screensaver. I looked for examples of this on PCS and never found any that worked. These do.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[David Trep](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-trep.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/david-trep-screensaver-wallpaper__1-32226/archive/master.zip)





### Source Code

<br><br>
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long<br><br><br>
Private Const SPI_SETSCREENSAVEACTIVE = 17<br>
Private Const SPIF_UPDATEINIFILE = &H1<br>
Private Const SPIF_SENDWININICHANGE = &H2<br>
Private Const SPI_GETSCREENSAVETIMEOUT = 14<br>
Private Const SPI_SETSCREENSAVETIMEOUT = 15<br>
Private Const SPI_SETDESKWALLPAPER = 20<br><br><br>
Private Sub ChangeWallPaper(strWP As String)<br>
 Dim ret As Long<br>
 ret = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, strWP, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)<br>
End Sub<br><br>
Private Sub ClearWallPaper()<br>
 Dim ret As Long<br>
 ret = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, "(None)", SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)<br>
End Sub<br><br>
Private Function ScreenSaverActive(Value As Boolean)<br>
 Call SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, Value, 0&, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)<br>
End Function<br><br><br>
Public Function SetScreenSaverTimeOut(ByVal NewValueInMinutes As Long) As Boolean<br>
 'Sets Screen Saver Timeout in Minutes
 <br>Dim lRet As Long<br>
 Dim lSeconds As Long<br>
 lSeconds = NewValueInMinutes * 60<br>
 lRet = SystemParametersInfo(SPI_SETSCREENSAVETIMEOUT, lSeconds, ByVal 0&, SPIF_UPDATEINIFILE + SPIF_SENDWININICHANGE)<br>
 SetScreenSaverTimeOut = lRet <> 0<br>
End Function<br>

