---
layout: default
title: Tracking class instances
nav_order: 4
permalink: /instance-tracker
---

In VBA, predeclared class instances can act as [object factories][Factories], and in general, it makes sense to keep this instance stateless. However, default instances can bear more responsibilities. For example, they can track regular class instances. The DbManager class below keeps a counter of its existing objects and runs a clean-up protocol whenever this count goes to zero.

**DbManager.cls**

```vb
'@Folder "DbManager"
'@ModuleDescription "Top database API class. Abstract factory for DbConnection."
'@PredeclaredId
Option Explicit

Private Type TDbManager
    Connections As Scripting.Dictionary
    InstanceCount As Long
End Type
Private this As TDbManager


Public Function Create() As DbManager
    Dim Instance As DbManager
    Set Instance = New DbManager
    Instance.Init
    Set Create = Instance
End Function

Friend Sub Init()
    Set this.Connections = New Scripting.Dictionary
    this.Connections.CompareMode = TextCompare
End Sub

Private Sub Class_Initialize()
    If Me Is DbManager Then
        this.InstanceCount = 0
    Else
        DbManager.InstanceAdd
    End If
End Sub

Private Sub Class_Terminate()
    DbManager.InstanceDel
End Sub

Public Property Get InstanceCount() As Long
    If Me Is DbManager Then
        InstanceCount = this.InstanceCount
    Else
        InstanceCount = DbManager.InstanceCount
    End If
End Property

Public Property Let InstanceCount(ByVal Value As Long)
    If Me Is DbManager Then
        this.InstanceCount = Value
    Else
        DbManager.InstanceCount = Value
    End If
End Property

Public Sub InstanceAdd()
    If Me Is DbManager Then
        this.InstanceCount = this.InstanceCount + 1
    Else
        DbManager.InstanceCount = DbManager.InstanceCount + 1
    End If
End Sub

Public Sub InstanceDel()
    If Me Is DbManager Then
        this.InstanceCount = this.InstanceCount - 1
        If this.InstanceCount = 0 Then
            '''' CLEANUP WHEN ALL REGULAR INSTANCES DESTROYED
        End If
    Else
        DbManager.InstanceCount = DbManager.InstanceCount - 1
    End If
End Sub
```


<!-- References -->

[Factories]: https://rubberduckvba.wordpress.com/2018/04/24/factories-parameterized-object-initialization/
