remove numbers

https://www.extendoffice.com/documents/excel/3243-excel-remove-numbers-from-strings.html

    Function RemoveNumbers(Txt As String) As String
    With CreateObject("VBScript.RegExp")
    .Global = True
    .Pattern = "[0-9]"
    RemoveNumbers = .Replace(Txt, "")
    End With
    End Function
 =RemoveNumbers(A1)
