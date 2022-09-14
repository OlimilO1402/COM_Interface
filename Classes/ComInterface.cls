VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ppv  As LongPtr
Private m_IID  As Guid
Private m_Name As String
Private m_Funx As List 'Of ComIntfcFunction

Private Sub Class_Initialize()
    Set m_Funx = New List: m_Funx.New_ EVbVarType.vbObject, , True
    'https://docs.microsoft.com/en-us/windows/win32/api/unknwn/nn-unknwn-iunknown
    'only Release will be used here, but it's mandatory
    'Add the functions for IUnknown, not needed, but just for completeness ;-)
    Dim Func As ComIntfcFunction, i As Long
    Set Func = New ComIntfcFunction: Func.New_ Me, "QueryInterface", i, EVbVarType.vbLong: m_Funx.Add Func: i = i + 1 ' 0
    Set Func = New ComIntfcFunction: Func.New_ Me, "AddRef", i, EVbVarType.vbULong:        m_Funx.Add Func: i = i + 1 ' 1
    Set Func = New ComIntfcFunction: Func.New_ Me, "Release", i, EVbVarType.vbULong:       m_Funx.Add Func ': i = i + 1 ' 2
End Sub
'Public Function ComInterface(aIID As Guid, Optional aName As String) As ComInterface
'    Set ComInterface = New ComInterface: ComInterface.New_ aIID, aName
'End Function

Private Sub Class_Terminate()
    If m_ppv = 0 Then Exit Sub
    'Dim varRet As Variant
    Dim Release As ComIntfcFunction: Set Release = m_Funx.Item(2)
    Release.Invoke
    'If DispCallFunc(m_ppv, IID_Release, CC_STDCALL, vbLong, 0&, 0&, 0&, varRet) <> 0 Then
    '    MsgBox Err.LastDllError
    'End If
End Sub

Friend Sub New_(aIID As Guid, Optional aName As String)
    Set m_IID = aIID
    m_Name = aName
End Sub

'Public Function Prepare(Optional aPPV As LongPtr = 0) As Boolean
'    If m_ppv = 0 Then m_ppv = aPPV
'    If m_ppv = 0 Then
'        MsgBox "ppv=0; this is illegal, create ppv first! (ppv = pointer to pointer to vtable)"
'        Exit Function
'    End If
'    m_Funcs.Add New_ComIntfcFunction
'End Function

Public Property Get Ptr() As LongPtr
    'return the pointer to m_ppv for writing to the memory-location of m_ppv
    Ptr = VarPtr(m_ppv)
End Property

Public Property Get ppv() As LongPtr
    'return the pointer to the pointer to the VTable
    ppv = m_ppv
End Property

Public Property Get IID() As Guid
    Set IID = m_IID
End Property

Public Sub Add(aFunction As ComIntfcFunction)
    m_Funx.Add aFunction
End Sub

Public Property Get Func(ByVal Index As Long) As ComIntfcFunction
    Set Func = m_Funx(Index)
End Property
Public Property Get FuncByName(ByVal FunctionName As String) As ComIntfcFunction
    Set FuncByName = m_Funx.ItemByKey(FunctionName)
End Property

