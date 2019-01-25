set DOT_NET=C:\Windows\Microsoft.NET\Framework64\v4.0.30319
set PATH=%PATH%;%DOT_NET%

csc /r:QRCoder.dll;DocumentFormat.OpenXml.dll  /out:qr.exe *.cs 
