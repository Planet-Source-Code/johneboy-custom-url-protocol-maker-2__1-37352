Attribute VB_Name = "SubMain"
Option Explicit
Sub Main()
Dim EXEpath As String
Dim Path As String
'just to make sure we get the quotes and %1 in the registry entry below
EXEpath = "" + Chr(34) + "" + App.Path + "\" + App.EXEName + ".exe" + Chr(34) + " %1"

'create the needed registry entries for our custom protocol

CreateKey "HKEY_CLASSES_ROOT\Tester"
CreateKey "HKEY_CLASSES_ROOT\Tester\shell"
CreateKey "HKEY_CLASSES_ROOT\Tester\shell\open"
CreateKey "HKEY_CLASSES_ROOT\Tester\shell\open\command"
SetStringValue "HKEY_CLASSES_ROOT\Tester", "", "URL: TesterProtocol"
SetStringValue "HKEY_CLASSES_ROOT\Tester", "URL Protocol", ""
SetBinaryValue "HKEY_CLASSES_ROOT\Tester", "EditFlags", Chr$(&H2) + Chr$(&H0) + Chr$(&H0) + Chr$(&H0)
SetStringValue "HKEY_CLASSES_ROOT\Tester\shell\open\command", "", "" & EXEpath


'command is the command that open the program, if you double clicked the icon, the path would be "" nothing.
'if a file called the opening of this program, the Command would be the path of that file

Path = Command
If Path = "" Then  ' if no file requested it open
        Form1.Show
        Else       ' if it was requested by a file, then do the following
        ExtractData1
        ExtractData2
        ExtractData3
        Form1.Show
        End If

End Sub
Public Sub ExtractData1()
Dim CMDLine As String
CMDLine = Command
CMDLine = Mid(CMDLine, InStr(1, CMDLine, ":") + 1)      'tells it, we want everything 1 space in front of the :
CMDLine = Mid(CMDLine, 1, InStr(1, CMDLine, "?") - 1)   'tells it, we want everything 1 space behind the ?
Form1.Text1.Text = CMDLine                              'text1 get the cleaned up stuff
End Sub

Public Sub ExtractData2()
Dim CMDLine As String
CMDLine = Command
CMDLine = Mid(CMDLine, InStr(1, CMDLine, "?") + 1)      'tells it, we want everything 1 space in front of the ?
CMDLine = Mid(CMDLine, 1, InStr(1, CMDLine, "$") - 1)   'tells it, we want everything 1 space behind the $
Form1.Text2.Text = CMDLine                              'text2 get the cleaned up stuff
End Sub

Public Sub ExtractData3()
Dim CMDLine As String
CMDLine = Command
CMDLine = Mid(CMDLine, InStr(1, CMDLine, "$") + 1)      'tells it, we want everything 1 space in front of the $
CMDLine = Mid(CMDLine, 1, InStr(1, CMDLine, "*") - 1)   'tells it, we want everything 1 space behind the *
'the above line isnt really needed, i dont think, but i left it there incase you want to add more
Form1.Text3.Text = CMDLine                              'text3 get the cleaned up stuff
End Sub

