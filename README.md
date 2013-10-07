MSAccessHookTests
=================

EasyHook tests for hooking into MSAccess. If new instances are spawned in VBA, also try to re-hook those as well.

This project is based on the EasyHook example, FileMon.

Notes:

* This project gives an example of intercepting COM objects upon creation.
* In CoCreateInstanceEx_Hook(), the custom hook for CoCreateInstanceEx, we manually marshal the MULTI_QI object to prevent crashes.
* You have to download EasyHook separately & add references. Also need EasyHook32.dll & EasyHook64.dll in the same folder as the compiled program.

To Run:
On command line, run MSAccessHookTests.exe <pid of MS-Access instance to hook>

Any Access COM objects that are subsequently created via VBA (Set App = New Access.Application), should also get hooked.

Issues:
The static instance of the Main class is not done the way that FileMon does it (is this an issue?)
Instances of MS-Access are kept until the main instance is closed.




