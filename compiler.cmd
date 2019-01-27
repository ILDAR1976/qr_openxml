@echo off
set DOT_NET=C:\Windows\Microsoft.NET\Framework64\v4.0.30319
set OLD_PATH=%PATH%
set PATH=%PATH%;%DOT_NET%

csc /r:QRCoder.dll;Pdf417.dll;DataMatrix.net.dll;System.Runtime.dll;DocumentFormat.OpenXml.dll  /out:qr.exe *.cs 

set PATH=%OLD_PATH%
