MSAccessHookTests
=================

EasyHook tests for hooking into MSAccess. If new instances are spawned in VBA, also try to re-hook those as well.

This project is based on the EasyHook example, FileMon.

Notes:

* This project gives an example of intercepting COM objects upon creation.
* In CoCreateInstanceEx_Hook(), the custom hook for CoCreateInstanceEx, we manually marshal the MULTI_QI object to prevent crashes.

