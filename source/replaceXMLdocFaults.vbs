' develop support script to replace all faults in XML docs (event properties)
Set fso = CreateObject("Scripting.FileSystemObject")
Set ifile = fso.OpenTextFile("bin\Release\DBaddin.xml", 1)
wscript.Echo "reading bin\Release\DBaddin.xml"
theText = ifile.readAll()
Set regEx = New RegExp
regEx.Global = True
regEx.IgnoreCase = True
regEx.Pattern = "F:DBaddin.MenuHandler._ctMenuStrip"
changedText = regEx.Replace(theText, "P:DBaddin.MenuHandler.ctMenuStrip")
regEx.Pattern = "F:DBaddin.AddInEvents._Application"
changedText = regEx.Replace(changedText, "P:DBaddin.AddInEvents.Application")
regEx.Pattern = "F:DBaddin.AddInEvents._cb"
changedText = regEx.Replace(changedText, "P:DBaddin.AddInEvents.cb")
regEx.Pattern = "F:DBaddin.AddInEvents._m"
changedText = regEx.Replace(changedText, "P:DBaddin.AddInEvents.m")


ifile.Close
Set ifile = Nothing
wscript.Echo "writing bin\Release\DBaddin.xml"
Set ofile = fso.CreateTextFile("bin\Release\DBaddin.xml", True)
ofile.write(changedText) 
ofile.Close
Set ofile = Nothing
Set fso = Nothing

