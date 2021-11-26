Attribute VB_Name = "ObjectStoreGlobals"
'@Folder "Object Store"
Option Explicit

'''' This pattern is functionally very similar to the predeclared feature.
'''' It requires explicit declaration of the variable. At the same time,
'''' this variable can be destroyed by setting it to Nothing, as opposed
'''' to predeclared class references.
'@Ignore EncapsulatePublicField: Alternative to predeclared attribute
Public ObjectStore As New ObjectStore

Public Enum ReferenceLoopManagementModeEnum
    REF_LOOP_NO_MANAGEMENT = 0&
    REF_LOOP_CLEANUP_CASCADE = 1&
    REF_LOOP_OBJECT_STORE = 2&
End Enum

Public ReferenceLoopManagementMode As ReferenceLoopManagementModeEnum

