Attribute VB_Name = "MNew"
Option Explicit

Public Function ComInterface(aIID As Guid, Optional aName As String) As ComInterface
    Set ComInterface = New ComInterface: ComInterface.New_ aIID, aName
End Function

Public Function ComIntfcFunction(aComIntfc As ComInterface, ByVal aName As String, ByVal aIndex As Long, ByVal ReturnType As EVbVarType) As ComIntfcFunction
    Set ComIntfcFunction = New ComIntfcFunction: ComIntfcFunction.New_ aComIntfc, aName, aIndex, ReturnType
End Function

