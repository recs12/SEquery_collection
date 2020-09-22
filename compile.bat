
:: Set this file for compiling the executable of the macro.

ipyc.exe /main:__main__.py ^
helper.py ^
Queries.py ^
Interop.SolidEdge.dll ^
/embed ^
/out:queries_collections_64x_0-0-1 ^
/platform:x64 ^
/standalone ^
/target:exe ^
/win32icon:icon.ico
