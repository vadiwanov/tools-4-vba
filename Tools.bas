Attribute VB_Name = "Tools"
Option Explicit
'Option Private Module

Sub TakeOff()

'Switch off some parameters for speeding a macro

    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayStatusBar = False
        .DisplayAlerts = False
    End With
    
End Sub

Sub Landing()

'Switch on all parameters

    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .DisplayStatusBar = True
        .DisplayAlerts = True
    End With
    
End Sub

Sub Fragmentation()

'Unloading VBA files in directory

                            Dim module As Object

If Len(Dir(ThisWorkbook.Path + "\Modules", vbDirectory)) = 0 Then MkDir (ThisWorkbook.Path + "\Modules")
    For Each module In ThisWorkbook.VBProject.VBComponents
        Select Case module.Type
            Case Is = 1
                ThisWorkbook.VBProject.VBComponents.Item(module.Name).Export ThisWorkbook.Path + "\Modules" + "\" + module.Name + ".bas"
            Case Is = 2
                ThisWorkbook.VBProject.VBComponents.Item(module.Name).Export ThisWorkbook.Path + "\Modules" + "\" + module.Name + ".frm"
            Case Is = 3
                ThisWorkbook.VBProject.VBComponents.Item(module.Name).Export ThisWorkbook.Path + "\Modules" + "\" + module.Name + ".cls"
        End Select
    Next
End Sub

Function RegExpSearch(Data As String, Pattern As String, Optional Item As Integer = 1) As String

'Function for searching through Regular Expressions

                            Dim oRegEx As Object
                            Dim Matches As Object

Set oRegEx = CreateObject("VBScript.RegExp")
    oRegEx.Pattern = Pattern
    oRegEx.Global = True
    If oRegEx.test(Data) Then
        Set Matches = oRegEx.Execute(Data)
            RegExpSearch = Matches.Item(Item - 1)
    End If
    
End Function

Function Substring(text, Delimiter, n) As String

'Function for getting a substring

                            Dim x As Variant
                            
    x = Split(text, Delimiter)
    If n > 0 And n - 1 <= UBound(x) Then
        Substring = x(n - 1)
    Else
        Substring = ""
    End If
End Function

Function CountChar(Data As String, Optional Char As String = " ") As Byte

'Function for getting a count of words or chars (words by default)

                            Dim i As Integer
                            Dim Count As Integer
                            
Count = 0

For i = 1 To Len(Data)
    If Mid(Data, i, 1) = Char Then Count = Count + 1
Next

CountChar = Count + 1

End Function

Function DoubleLike(xWord As String, yWord As String, zWord As String) As Boolean

'function for easy double 'like' compare

                            Dim isDoubleLike As Boolean
                            
isDoubleLike = False

If zWord Like "*" & xWord & "*" Or zWord Like "*" & yWord & "*" Then isDoubleLike = True

DoubleLike = isDoubleLike

End Function

Function ReverseLike(xData As String, yData As String) As Boolean

'Function for reverse 'like' checking two strings

                            Dim isReverseLike As Boolean
                            
isReverseLike = False

If Len(Trim(xData)) = 0 Or Len(Trim(yData)) = 0 Then
    isReverseLike = False
    Exit Function
End If

If xData Like "*" & yData & "*" Or yData Like "*" & xData & "*" Then isReverseLike = True

ReverseLike = isReverseLike

End Function

Function GetPercent(xValue As Long, yValue As Long) As Byte

'Function for getting a percent value from two digits

GetPercent = (yValue / xValue) * 100

End Function

Function Incremention(Value As Integer, Optional Increment As Integer = 1) As Integer

'function for digits incremention

Value = Value + Increment

Incremention = Value

End Function

Function Concatenation(xWord As String, yWord As String, Delimeter As String) As String

'Function for data concatination

If Len(Trim(xWord)) = 0 Then
    Delimeter = Chr(32)
Else
    Delimeter = Chr(10)
End If

xWord = xWord & Delimeter & yWord

Concatenation = xWord

End Function

Function ReverseInstr(xData, yData) As Boolean

'Function for double 'InStr' checking

ReverseInstr = False

If InStr(xData, yData) > 0 Then
    If InStr(yData, xData) > 0 Then
        ReverseInstr = True
    End If
End If

End Function
