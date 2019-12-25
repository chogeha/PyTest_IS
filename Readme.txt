Preliminary Version for Microsoft Windows System

Requirement:
	1. Microsoft Windows OS 7 and later (X86/X64)
	2. Python 3.5+ and later
	3. InstallShield 2016 and later of manipulate experience.
	4. "Microsoft Windows Installer" of manipulate experience.

Feature:
	1 Upgrade Test. (from OLD to NEW) 
	2 Downgrade Test (from NEW to OLD)
	3 Uninstall and Keep The Last Test (For different of IS EXE package but share the same of IS MSI package)
		- IS EXE: InstallShield *.exe installation package
		- IS MSI: InstallShield *.msi installation package
		
File Structure:
	+ EVT_Test.py : Demonstrate how to run testing.
	+ Library\ADJ_IS_Tool.py : Core library.
	+ RG_UseCase.png: Concept of user scenario diagram.
	+ Config\CfgInit.ini : Initial configuration.
	+ Config\CfgReg.ini : Windows registry property check.
