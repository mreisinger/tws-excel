Attribute VB_Name = "MiscFunctions"
Public Function transposeArray(myArr() As Variant) As Variant
    Dim temp() As Variant
    ReDim temp(UBound(myArr, 2), UBound(myArr, 1))
    
    transposeArray = Application.transpose(myArr)

End Function
