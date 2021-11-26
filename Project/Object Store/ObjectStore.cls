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
''''
'''' This classes uses self-assigned declaration as an alternative to setting
'''' the predeclared attribute, making it possible to destroy it. Presently,
'''' this class is also used as singleton.
''''
Option Explicit

Private Type TObjectStore
    Store As Scripting.Dictionary
End Type
Private this As TObjectStore


Private Sub Class_Initialize()
    Set this.Store = New Scripting.Dictionary
    this.Store.CompareMode = TextCompare
End Sub


Private Sub Class_Terminate()
    this.Store.RemoveAll
    Set this.Store = Nothing
End Sub


'''' For testing only
Friend Property Get Store() As Scripting.Dictionary
    Set Store = this.Store
End Property


Public Function AllocateHandle(Optional ByVal Handle As String = vbNullString _
                              ) As Currency
    Dim IDCandidate As Currency
    If Len(Handle) > 0 Then
        If Not this.Store.Exists(Handle) Then
            this.Store(Handle) = Empty
            IDCandidate = -1
        Else
            IDCandidate = 0
        End If
    Else
        IDCandidate = CCur(GenerateSerialID() / 10000)
        Do While this.Store.Exists(IDCandidate)
            IDCandidate = IDCandidate + 1
        Loop
        this.Store(IDCandidate) = Empty
    End If
    AllocateHandle = IDCandidate
End Function


Public Sub FreeHandle(ByVal Handle As Variant)
    If this.Store.Exists(Handle) Then this.Store.Remove Handle
End Sub


Public Function RefSet(ByVal Handle As Variant, ByVal ObjRef As Object) As Variant
    If Not this.Store.Exists(Handle) Then
        RefSet = vbNullString
    ElseIf Not IsEmpty(this.Store(Handle)) Then
        RefSet = vbNullString
    Else
        Set this.Store(Handle) = ObjRef
        RefSet = Handle
    End If
End Function


Public Function RefGet(ByVal Handle As Variant) As Object
    If Not this.Store.Exists(Handle) Then
        Set RefGet = Nothing
    ElseIf IsEmpty(this.Store.Exists(Handle)) Then
        Set RefGet = Nothing
    Else
        Set RefGet = this.Store(Handle)
    End If
End Function
