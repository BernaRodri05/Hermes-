Set oWS = WScript.CreateObject("WScript.Shell")
sLinkFile = oWS.CurrentDirectory & "\HERMES V1.lnk"
Set oLink = oWS.CreateShortcut(sLinkFile)
oLink.TargetPath = oWS.CurrentDirectory & "\archivos\EJECUTAR.bat"
oLink.WorkingDirectory = oWS.CurrentDirectory & "\archivos"
oLink.IconLocation = oWS.CurrentDirectory & "\python_icon.ico"
oLink.Description = "HERMES V1 - Envío de mensajes WhatsApp"
oLink.Save
WScript.Echo "Acceso directo creado exitosamente"
