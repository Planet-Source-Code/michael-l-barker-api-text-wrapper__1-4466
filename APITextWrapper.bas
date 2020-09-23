Attribute VB_Name = "APITextWrapper"
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const ES_LOWERCASE = &H10&
Private Const ES_UPPERCASE = &H8&
Private Const ES_NUMBER = &H2000&

' Comments  : Allow only numbers in a textbox
' Returns   : The Style of the textbox before the change.
Public Function NumbersOnly(tBox As TextBox)
Dim DefaultStyle As Long
DefaultStyle = GetWindowLong(tBox.hwnd, GWL_STYLE)
NumbersOnly = SetWindowLong(tBox.hwnd, GWL_STYLE, DefaultStyle Or ES_NUMBER)
End Function

' Comments  : Allow only uppercase letters in a textbox
' Returns   : The Style of the textbox before the change.
Public Function UpperCaseOnly(tBox As TextBox)
Dim DefaultStyle As Long
DefaultStyle = GetWindowLong(tBox.hwnd, GWL_STYLE)
UpperCaseOnly = SetWindowLong(tBox.hwnd, GWL_STYLE, DefaultStyle Or ES_UPPERCASE)
End Function

' Comments  : Allow only lowercase letters in a textbox
' Returns   : The Style of the textbox before the change.
Public Function LowerCaseOnly(tBox As TextBox)
Dim DefaultStyle As Long
DefaultStyle = GetWindowLong(tBox.hwnd, GWL_STYLE)
LowerCaseOnly = SetWindowLong(tBox.hwnd, GWL_STYLE, DefaultStyle Or ES_LOWERCASE)
End Function


' Comments  : Sets the style of a textbox.
' Returns   : The new style.
Public Function SetStyle(tBox As TextBox, NewStyle As Long)
SetStyle = SetWindowLong(tBox.hwnd, GWL_STYLE, NewStyle)
End Function


' Comments  : Gets the current style of a textbox.
' Returns   : The Style of the textbox.
Public Function GetStyle(tBox As TextBox)
GetStyle = GetWindowLong(tBox.hwnd, GWL_STYLE)
End Function

Public Function StyleNumberToText(tBox As TextBox)
Dim StyleNum  As Long
Dim StyleText As String

StyleNum = GetStyle(tBox)

Select Case StyleNum
    Case 1409360064: StyleText = "Number"
    Case 1409351880: StyleText = "Uppercase"
    Case 1409351888: StyleText = "Lowercase"
    Case Else: StyleText = "Other"
End Select

StyleNumberToText = StyleText
End Function
