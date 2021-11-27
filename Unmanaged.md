---
layout: default
title: Unmanaged circular references
nav_order: 2
permalink: /unmanaged-refs
---


Each demo class includes Class_Initialize() and Class_Terminate() routines printing class name, instance type (default/predeclared or regular), and sub name. Setting `ReferenceLoopManagementMode = REF_LOOP_NO_MANAGEMENT` in the Main.Main() entry selects the unmanaged circular references regime. The first run of the Main.Main() sub should yield a similar output in the immediate panel:

```
2021-11-27 00:25:35.579: DbManager   /Default - Class_Initialize
2021-11-27 00:25:35.580: DbManager   /Regular - Class_Initialize
2021-11-27 00:25:35.580: DbConnection/Default - Class_Initialize
2021-11-27 00:25:35.581: DbConnection/Regular - Class_Initialize
2021-11-27 00:25:35.581: DbStatement /Default - Class_Initialize
2021-11-27 00:25:35.582: DbStatement /Regular - Class_Initialize
2021-11-27 00:25:35.582: DbManager   /Regular - Class_Terminate
```

Note how Class_Initialize() is executed twice for each class - first for the default and then for the regular instance. For example, on the first run, the following line `Set dbm = DbManager.Create()` causes initialization of the default DbManager instance first, followed by initialization of the regular instance created inside the factory. This behavior of predeclared variables is similar to that of auto-assigned variables: both are instantiated on first use. 

The other important observation is that there is only one Class_Terminate() event, which is due to the regular instance of DbManager. DbManager is not involved in circular references, and it gets disposed of by VBA. The two other classes form an unmanaged reference loop and cannot be destroyed automatically. The default instances, however, are stateless and not affected by the circular reference issue. Still, all of them remain usable past the program termination. Repeated executions of the MainMain() sub confirm the persistence of the default instances. The new output should be similar to this:

```
2021-11-27 01:15:58.912: DbManager   /Regular - Class_Initialize
2021-11-27 01:15:58.913: DbConnection/Regular - Class_Initialize
2021-11-27 01:15:58.913: DbStatement /Regular - Class_Initialize
2021-11-27 01:15:58.914: DbManager   /Regular - Class_Terminate
```

showing no events from the default instances. These events come back after a Reset.
