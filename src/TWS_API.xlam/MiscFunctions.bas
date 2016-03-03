Attribute VB_Name = "MiscFunctions"
Public Function transposeArray(myArr() As Variant) As Variant
    Dim temp() As Variant
    ReDim temp(UBound(myArr, 2), UBound(myArr, 1))
    
    transposeArray = Application.transpose(myArr)

End Function


Public Function regexCompare(strInput As String, strPattern As String) As Boolean

    Dim regEx As New RegExp
    
    regEx.Pattern = strPattern


    If regEx.test(strInput) Then
        regexCompare = True
    Else
        regexCompare = False
    End If

End Function
