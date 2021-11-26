Attribute VB_Name = "ObjectStoreTests"
Attribute VB_Description = "Tests for the Guard class."
'@Folder "Object Store"
'@TestModule
'@ModuleDescription "Tests for the Guard class."
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext, IndexedDefaultMemberAccess
Option Explicit
Option Private Module

Private Const MODULE_NAME As String = "ObjectStoreTests"
Private TestCounter As Long

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
    With Logger
        .ClearLog
        .DebugLevelDatabase = DEBUGLEVEL_MAX
        .DebugLevelImmediate = DEBUGLEVEL_NONE
        .UseIdPadding = True
        .UseTimeStamp = False
        .RecordIdDigits 3
        .TimerSet MODULE_NAME
    End With
    TestCounter = 0
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
    Logger.TimerLogClear MODULE_NAME, TestCounter
    Logger.PrintLog
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Object Store")
Private Sub ztcNew_VerifySelfAssignment()
    On Error GoTo TestFail
    
Arrange:
Act:
    Set ObjectStore = Nothing
Assert:
    Assert.IsFalse ObjectStore Is Nothing, "ObjectStore is not set."
    Assert.IsFalse ObjectStore.Store Is Nothing, "Store is not set."
    Assert.AreEqual 0, ObjectStore.Store.Count, "Store count mismatch."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Handles")
Private Sub ztcAllocateHandle_VerifyAllocateHandle()
    On Error GoTo TestFail
    
Arrange:
    Set ObjectStore = Nothing
Act:
    Dim Result As Currency
Assert:
    Result = ObjectStore.AllocateHandle("ABC")
    Assert.AreEqual 1, ObjectStore.Store.Count, "Store count mismatch."
    Assert.AreEqual -1, Result, "Result mismatch with new custom handle."
    Assert.IsTrue ObjectStore.Store.Exists("ABC"), "Handle should be in the Store."
    Assert.IsTrue IsEmpty(ObjectStore.Store("ABC")), "Element should be empty."
    
    Result = ObjectStore.AllocateHandle("ABC")
    Assert.AreEqual 1, ObjectStore.Store.Count, "Store count mismatch."
    Assert.AreEqual 0, Result, "Result mismatch with existing handle."

    Result = ObjectStore.AllocateHandle(vbNullString)
    Assert.AreEqual 2, ObjectStore.Store.Count, "Store count mismatch with new auto handle."
    Assert.IsTrue Result > 0, "Result mismatch with new auto handle."

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Handles")
Private Sub ztcFreeHandle_VerifyFreeHandle()
    On Error GoTo TestFail
    
Arrange:
    Dim RefTimeStamp As Double
    RefTimeStamp = CDbl(DateDiff("s", DateSerial(1970, 1, 1), Date)) + Timer
    Set ObjectStore = Nothing
Act:
    Dim Result As Currency
Assert:
    Result = ObjectStore.AllocateHandle(vbNullString)
    Assert.AreEqual 1, ObjectStore.Store.Count, "Store count mismatch."
    Assert.IsTrue Result >= RefTimeStamp, "Result mismatch with new auto handle."
    Result = ObjectStore.AllocateHandle(vbNullString)
    Assert.AreEqual 2, ObjectStore.Store.Count, "Store count mismatch."
    ObjectStore.FreeHandle vbNullString
    Assert.AreEqual 2, ObjectStore.Store.Count, "Store count mismatch."
    ObjectStore.FreeHandle Result
    Assert.AreEqual 1, ObjectStore.Store.Count, "Store count mismatch."
    Result = ObjectStore.AllocateHandle("ABC")
    Assert.AreEqual 2, ObjectStore.Store.Count, "Store count mismatch."
    Assert.AreEqual -1, Result, "Result mismatch with new custom handle."
    ObjectStore.FreeHandle "ABC"
    Assert.AreEqual 1, ObjectStore.Store.Count, "Store count mismatch."
    Dim Store As Scripting.Dictionary
    Set Store = ObjectStore.Store
    Set ObjectStore = Nothing
    Assert.AreEqual 0, Store.Count, "Store count mismatch."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("References")
Private Sub ztcRefSet_VerifyRefSet()
    On Error GoTo TestFail
    
Arrange:
    Set ObjectStore = Nothing
Act:
    Dim Result As Variant
Assert:
    Result = ObjectStore.RefSet("ABC", Application)
    Assert.AreEqual vbNullString, Result, "Result mismatch with non-allocated key."
    Assert.AreEqual 0, ObjectStore.Store.Count, "Store count mismatch."
    ObjectStore.Store("ABC") = "ABC"
    Result = ObjectStore.RefSet("ABC", Application)
    Assert.AreEqual vbNullString, Result, "Result mismatch with non-empty slot."
    Assert.AreEqual "ABC", ObjectStore.Store("ABC"), "Existing value should not change."
    Assert.AreEqual 1, ObjectStore.Store.Count, "Store count mismatch."
    ObjectStore.Store("ABC") = Empty
    Result = ObjectStore.RefSet("ABC", Application)
    Assert.IsTrue ObjectStore.Store("ABC") Is Application, "Reference is not saved."
    Assert.AreEqual "ABC", Result, "Result mismatch."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("References")
Private Sub ztcRefGet_VerifyRefGet()
    On Error GoTo TestFail
    
Arrange:
    Set ObjectStore = Nothing
Act:
    Dim Result As Variant
Assert:
    ObjectStore.Store("ABC") = Empty
    Result = ObjectStore.RefSet("ABC", Application)
    Assert.AreEqual "ABC", Result
    Assert.IsTrue ObjectStore.RefGet("ABC") Is Application, "Reference is not saved."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
