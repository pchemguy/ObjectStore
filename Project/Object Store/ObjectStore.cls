VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Manages object references using scalar handles."
'@Folder "Object Store"
'@ModuleDescription "Manages object references using scalar handles."
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit

Private Type TObjectStore
    Store As Scripting.Dictionary
End Type
Private this As TObjectStore


Private Sub Class_Initialize()
    Debug.Print GetTimeStampMs & ": ObjectStore - Class_Initialize"
    Set this.Store = New Scripting.Dictionary
    this.Store.CompareMode = TextCompare
End Sub

Private Sub Class_Terminate()
    Debug.Print GetTimeStampMs & ": ObjectStore - Class_Terminate"
    this.Store.RemoveAll
    Set this.Store = Nothing
End Sub

'''' For testing only
Friend Property Get Store() As Scripting.Dictionary
    Set Store = this.Store
End Property

#If VBA7 Then
Public Sub DelRef(ByVal Handle As LongPtr)
#Else
Public Sub DelRef(ByVal Handle As Long)
#End If
    If this.Store.Exists(Handle) Then this.Store.Remove Handle
End Sub

#If VBA7 Then
Public Function SetRef(ByVal ObjRef As Object) As LongPtr
    Dim Handle As LongPtr
#Else
Public Function SetRef(ByVal ObjRef As Object) As Long
    Dim Handle As Long
#End If
    Handle = ObjPtr(ObjRef)
    If Not this.Store.Exists(Handle) Then
        Set this.Store(Handle) = ObjRef
    End If
    SetRef = Handle
End Function

#If VBA7 Then
Public Function GetRef(ByVal Handle As LongPtr) As Object
#Else
Public Function GetRef(ByVal Handle As Long) As Object
#End If
    If this.Store.Exists(Handle) Then
        Set GetRef = this.Store(Handle)
    Else
        Set GetRef = Nothing
    End If
End Function
