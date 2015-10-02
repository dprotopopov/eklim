Attribute VB_Name = "Module6"
Sub testAI()
Dim appRef, x
x = Shell("X:\Program Files\Adobe\Adobe Illustrator CS6 (64 Bit)\Support Files\Contents\Windows\Illustrator.exe", vbNormalFocus)
Set appRef = CreateObject("Illustrator.Application")
End Sub
