Attribute VB_Name = "ObjectStoreGlobals"
'@Folder "Object Store"
Option Explicit

'''' This pattern is functionally very similar to the predeclared feature.
'''' It requires explicit declaration of the variable. At the same time,
'''' this variable can be destroyed by setting it to Nothing, as opposed
'''' to predeclared class references.
'@Ignore EncapsulatePublicField: Alternative to predeclared attribute
Public ObjectStore As New ObjectStore

