:: system will wait for the subprocess when using call
::@ echo off
::call coilsim.exe

:: system will not wait for the subprocess when using start
@ echo off
start/min/wait coilsim.exe