Attribute VB_Name = "MNew"
Option Explicit

Public Function ComInterface(aIID As Guid, Optional aName As String) As ComInterface
    Set ComInterface = New ComInterface: ComInterface.New_ aIID, aName
End Function

Public Function ComIntfcFunction(aComIntfc As ComInterface, ByVal aName As String, ByVal aIndex As Long, ByVal ReturnType As EVbVarType) As ComIntfcFunction
    Set ComIntfcFunction = New ComIntfcFunction: ComIntfcFunction.New_ aComIntfc, aName, aIndex, ReturnType
End Function

Public Function VVariant(aValue) As VVariant
    'Create a VVariant-object with a Variant
    Set VVariant = New VVariant: VVariant.New_ aValue
End Function

Public Function VVariantVt(vt As EVbVarType, aValue) As VVariant
    'Create a VVariant-object with a Variant and set the vartype yourself,
    'like e.g. give a signed Long and set vt to unsigned Long
    Set VVariantVt = New VVariant: VVariantVt.NewVt vt, aValue
End Function

