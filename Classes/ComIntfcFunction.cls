VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComIntfcFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const S_OK As Long = 0

'typedef enum tagCALLCONV {
'  CC_FASTCALL = 0,
'  CC_CDECL = 1,
'  CC_MSCPASCAL,
'  CC_PASCAL,
'  CC_MACPASCAL,
'  CC_STDCALL,
'  CC_FPFASTCALL,
'  CC_SYSCALL,
'  CC_MPWCDECL,
'  CC_MPWPASCAL,
'  CC_MAX
'} CALLCONV;
Public Enum CallConv
'    CC_FASTCALL = 0
'    CC_CDECL = 1
'    CC_MSCPASCAL = 2
'    CC_PASCAL = 3
'    CC_MACPASCAL = 3
    CC_STDCALL = 4
'    CC_FPFASTCALL = 5
'    CC_SYSCALL = 6
'    CC_MPWCDECL = 7
'    CC_MPWPASCAL = 8
'    CC_MAX = 9
End Enum
' ----==== Oleaut32 Deklarationen ====----
'https://docs.microsoft.com/en-us/windows/win32/api/oleauto/nf-oleauto-dispcallfunc
Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As Long, ByVal cc As Long, ByVal vtReturn As EVbVarType, ByVal cActuals As Long, ByVal prgvt As LongPtr, ByVal prgpvarg As LongPtr, ByVal pvargResult As LongPtr) As Long

Private m_Com    As ComInterface
Private m_Name   As String
Private m_Index  As Long
Private m_VOfst  As LongPtr 'VTable-Offset = m_index * SizeOf_LongPtr
Private m_RType  As EVbVarType
Private m_VarRet As VVariant

Friend Sub New_(aComIntfc As ComInterface, ByVal aName As String, ByVal aIndex As Long, ByVal ReturnType As EVbVarType)
    Set m_Com = aComIntfc: m_Name = aName: m_Index = aIndex
    m_VOfst = m_Index * MPtr.SizeOf_LongPtr: m_RType = ReturnType
    Set m_VarRet = New VVariant: m_VarRet.New_ ReturnType
End Sub

'Public Function ComIntfcFunction(aComIntfc As ComInterface, ByVal aName As String, ByVal aIndex As Long, ByVal ReturnType As EVbVarType) As ComIntfcFunction
'    Set ComIntfcFunction = New ComIntfcFunction: ComIntfcFunction.New_ aComIntfc, aName, aIndex, ReturnType
'End Function

Public Property Get Name() As String
    Name = m_Name
End Property
Public Property Get Key() As String
    Key = m_Name
End Property

Public Property Get ReturnType() As EVbVarType
    ReturnType = m_RType
End Property

'most of the time the return type is HResult = vbLong = VT_INt=3
Public Function Invoke(ParamArray params()) As Variant
'    If m_Com Is Nothing Then Exit Function
'    Dim ppv As LongPtr: ppv = m_Com.ppv
'    If UBound(params) >= 0 Then
'        Dim varParam As Variant: varParam = params
'        'If IsArray(varParam) Then varParam = varParam(0)
'        Dim u As Long: u = UBound(varParam)
'        Dim olePtr(10) As LongPtr
'        Dim oleTyp(10) As Integer
'        Dim i As Long
'        For i = 0 To u
'            oleTyp(i) = VarType(varParam(i))
'            olePtr(i) = VarPtr(varParam(i))
'        Next
'    End If
'    If DispCallFunc(ppv, m_VOfst, CC_STDCALL, m_RType, i, VarPtr(oleTyp(0)), VarPtr(olePtr(0)), ByVal m_VarRet.Ptr) <> S_OK Then
'        Debug.Print "Fehler beim Aufrufen der Interface-Funktion: " & m_Name
'    End If
'    Set Invoke = m_VarRet

    Dim ppv As LongPtr: ppv = m_Com.ppv
    Dim v As Variant
    v = PrivateInvoke(ppv, m_Index, params)
    
    Dim vv As VVariant: Set vv = MNew.VVariant(v)
    Set Invoke = vv
End Function

Private Function PrivateInvoke(ByVal pInterface As Long, ByVal eInterfaceFunction As Long, ParamArray arrParam()) As Variant
    
    If pInterface <> 0& Then
        
        PrivateInvoke = OleInvoke(pInterface, eInterfaceFunction, arrParam)
        
    End If
    
End Function

'
'' ----==== Call OleInvoke ====----
'Private Function PrivInvoke(ByVal pInterface As Long, ByVal eInterfaceFunction As vtb_Interfaces, ParamArray params()) As Variant
'
'    If pInterface <> 0& Then
'
'        PrivInvoke = OleInvoke(pInterface, eInterfaceFunction, arrParam)
'
'    End If
'
'End Function

' ----==== Call Interface Function ====----
'Private Function OleInvoke(ByVal pInterface As LongPtr, ByVal lngCmd As Long, ParamArray arrParam()) As Variant
'
'    If pInterface <> 0& Then
'
'        If UBound(arrParam) >= 0 Then
'
'            Dim varParam As Variant: varParam = arrParam
'
'            If IsArray(varParam) Then varParam = varParam(0)
'
'            Dim u As Long: u = UBound(varParam)
'
'            Dim olePtr(10) As LongPtr
'            Dim oleTyp(10) As Integer
'            Dim i As Long
'            For i = 0 To u
'
'                oleTyp(i) = VarType(varParam(i))
'
'                olePtr(i) = VarPtr(varParam(i))
'
'            Next
'
'        End If
'
'        'Dim varRet As VVariant: Set varRet = New VVariant
'        If DispCallFunc(pInterface, lngCmd * 4, CC_STDCALL, vbLong, lngItem, VarPtr(oleTyp(0)), VarPtr(olePtr(0)), ByVal varRet.Ptr) <> S_OK Then
'
'            Debug.Print "Fehler beim Aufrufen der Interface-Funktion!"
'
'        End If
'
'    End If
'
'    OleInvoke = varRet
'
'End Function
'
' ----==== Call Interface Function ====----
Private Function OleInvoke(ByVal pInterface As Long, ByVal lngCmd As Long, ParamArray arrParam()) As Variant
    
    Dim lngItem As Long
    Dim lngCount As Long
    Dim varRet As Variant
    Dim varParam As Variant
    Dim olePtr(10) As Long
    Dim oleTyp(10) As Integer

    If pInterface <> 0& Then

        If UBound(arrParam) >= 0 Then

            varParam = arrParam
            
            If IsArray(varParam) Then varParam = varParam(0)

            lngCount = UBound(varParam)

            For lngItem = 0 To lngCount

                oleTyp(lngItem) = VarType(varParam(lngItem))

                olePtr(lngItem) = VarPtr(varParam(lngItem))

            Next

        End If

        If DispCallFunc(pInterface, lngCmd * 4, CC_STDCALL, m_RType, lngItem, VarPtr(oleTyp(0)), VarPtr(olePtr(0)), varRet) <> S_OK Then
            
            Debug.Print "Fehler beim Aufrufen der Interface-Funktion!"
            
        End If
        
    End If

    OleInvoke = varRet

End Function

