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
Private Sub ztcRefXeX_VerifyRefXeX()
    On Error GoTo TestFail
    
Arrange:
    Set ObjectStore = Nothing
    #If VBA7 Then
        Dim Handle As LongPtr
    #Else
        Dim Handle As Long
    #End If
Act:
Assert:
    Assert.AreEqual 0, ObjectStore.Store.Count, "Store count mismatch."
    Handle = ObjectStore.SetRef(Application)
    Assert.AreEqual ObjPtr(Application), Handle, "SetRef - Object handle mismatch"
    Handle = ObjectStore.SetRef(ThisWorkbook)
    Assert.AreEqual ObjPtr(ThisWorkbook), Handle, "SetRef - Object handle mismatch"
    Assert.AreEqual 2, ObjectStore.Store.Count, "Store count mismatch."
    Dim AppObj As Excel.Application
    Set AppObj = ObjectStore.GetRef(ObjPtr(Application))
    Assert.IsTrue AppObj Is Application, "GetRef - Object mismatch"
    Assert.IsTrue ObjectStore.GetRef(0) Is Nothing, "GetRef - Object mismatch"
    ObjectStore.DelRef ObjPtr(Application)
    Assert.AreEqual 1, ObjectStore.Store.Count, "Store count mismatch."
    Assert.IsTrue ObjectStore.GetRef(ObjPtr(ThisWorkbook)) Is ThisWorkbook, "GetRef - Object mismatch"
    ObjectStore.DelRef ObjPtr(ThisWorkbook)
    Assert.AreEqual 0, ObjectStore.Store.Count, "Store count mismatch."
    
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
