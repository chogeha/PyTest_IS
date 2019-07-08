#!/usr/bin/env python3
'''Purpose for InstallShield Installation Package of Regression Test'''

__all__ = [	'mf_InitEnv',
			'mf_DoTest',
			'_mf_ResetTempFolder',
			'_mf_UpgradeTest',
			'_mf_DowngradeTest',
			'_mf_UninstallKeepTest',
			'_mf_PreUninstall',
			'_mf_UninstallPackage',
			'_mf_ExecuteProcess',      
			'_mf_QueryRegINIList',
			'_mf_QueryRegProperty',
			'_mf_ResultCheck',
			'_mf_VersionCheck',
           ]
__author__ = "Jim.Chiu <chogeha@gmail.com>"
__date__ = "08 July 2019"

# Known bugs that can't be fixed here:
#   - TBD
import os
import ctypes
import winreg
import shutil
import subprocess
import configparser
import platform

# For Switch the content of log repository between old and new file
MODE_OLD = 0
MODE_NEW = 1
# For Version Check Mode Selection
MODE_REPORT_UG = 0
MODE_REPORT_DG = 1
MODE_REPORT_KEEP = 2

# For Uninstallation Test
MODE_KEEP_OLD = 0
MODE_KEEP_NEW = 1

STR_PATH_REG_UNINST = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\"
STR_PATH_REG_UNINST_SYSWOW = "SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\"

class ADJ_IS_TOOL:
    def __init__(self):
        '''Initial parameters.			
        '''
        print("========== SYSTEM INFO ==========")
        self.dictSystemEnv = platform.uname()._asdict()
        for strName in self.dictSystemEnv:
            print(strName.upper() + " = " + self.dictSystemEnv[strName])

        if platform.machine() == 'AMD64':
            self.Is64bitOS = True
            print("System is 64bits.")
        else:
            self.Is64bitOS = False
            print("System is 32bits.")

        self.strCurrentPath = os.path.dirname(os.path.abspath(__file__))
        self.strCurrentPath = self.strCurrentPath.replace('\Library','')
        print("Current Folder Path = " + self.strCurrentPath)
           
        strTemp = self.strCurrentPath + "\Result"
        if os.path.isdir(strTemp) == True:
            shutil.rmtree(strTemp)
            os.mkdir(strTemp)
        else:
            os.mkdir(strTemp)

        self.strPathTempFileOld = self.strCurrentPath + "\Temp\TempOld.ini" 
        self.strPathTempFileNew = self.strCurrentPath + "\Temp\TempNew.ini"
        self.strPathReportUG = self.strCurrentPath + "\Result\ReportUG.log"
        self.strPathReportDG = self.strCurrentPath + "\Result\ReportDG.log"
        self.strPathReportKeep = self.strCurrentPath + "\Result\ReportKEEP.log"
        self.strPathCfgInit = self.strCurrentPath + "\Config\CfgInit.ini" # Path of Initialization of Configuration
        self.strPathCfgReg = self.strCurrentPath + "\Config\CfgReg.ini"  # Path of Test Item's Registry of Information
		
        self.nLogMode = -1 # 0: OLD , 1: NEW
        self.hContentOLD = "" # OLD of IS MSI Infomation
        self.bSilentInstallOLD = False
        self.hContentNEW = "" # NEW of IS MSI Infomation
        self.bSilentInstallNEW = False
        self.strPkgOld = ""
        self.strPkgOldPDID = ""
        self.strPkgNew = ""
        self.strPkgNewPDID = ""
        self.bTestUpgrade = False
        self.bTestDowngrade = False
        self.bTestVersionCheck = False
        self.strVerACE = ""
        self.strVerPXI = ""
        self.strVerSFP = ""
        self.strVerPCIS = ""
        self.strVerD2K = ""
        self.strVerDSA = ""
        self.strVerUD = ""
        self.strVerWD = ""
        self.dictVer = dict()
        self.dictPreUninstall = dict()

    def mf_InitEnv(self, strPathCfgInit = "", strPathCfgReg = ""):
        '''Initial EVT Test Environment.
        '''	    

        if strPathCfgInit != "":
            self.strPathCfgInit = strPathCfgInit
        
        if strPathCfgReg != "":            
            self.strPathCfgReg = strPathCfgReg

        cfgINI = configparser.ConfigParser()
        fileOP = open(self.strPathCfgInit, 'r', encoding='utf-8')
        cfgINI.read_file(fileOP)

        # Environment Configuration Setting
        if cfgINI.has_section('Environment') == True:
            print("\n========== Environment Section ==========\n")
            cfgSection = "Environment"           
            self.strPkgOld = self.strCurrentPath + '\\' + str(cfgINI[cfgSection]['BEF_PKG']).replace('"','')
            self.strPkgOldPDID = cfgINI[cfgSection]['BEF_PD_ID']
            self.bSilentInstallOLD = int(str(cfgINI[cfgSection]['BEF_SILENT_ON']).replace('"',''))
            self.strPkgNew = self.strCurrentPath + '\\' + str(cfgINI[cfgSection]['AFT_PKG']).replace('"','')
            self.strPkgNewPDID = cfgINI[cfgSection]['AFT_PD_ID']
            self.bSilentInstallNEW = int(str(cfgINI[cfgSection]['AFT_SILENT_ON']).replace('"',''))
            self.bTestUpgrade = int(str(cfgINI[cfgSection]['UPGRADE_CHK']).replace('"',''))
            self.bTestDowngrade = int(str(cfgINI[cfgSection]['DOWNGRADE_CHK']).replace('"',''))
            self.bTestUninstallKeep = int(str(cfgINI[cfgSection]['UNINSTALL_KEEP_CHK']).replace('"',''))
            self.bTestVersionCheck = int(str(cfgINI[cfgSection]['VERSION_CHK']).replace('"',''))
            print("Before Package    = " + self.strPkgOld)
            print("Before Product ID = " + self.strPkgOldPDID.replace('"',''))
            print("Before Silent Flag= " + str(self.bSilentInstallOLD))
            print("After Package     = " + self.strPkgNew)
            print("After Product  ID = " + self.strPkgNewPDID.replace('"',''))
            print("After Silent Flag = " + str(self.bSilentInstallNEW))
            print("Upgrade Test Flag         = " + str(self.bTestUpgrade))
            print("Downgrade Test Flag       = " + str(self.bTestDowngrade))
            print("Uninstall Keep Test Flag  = " + str(self.bTestDowngrade))
            print("Fixed Version Check Flag  = " + str(self.bTestVersionCheck))
            print("")
        else:
            raise ValueError("File " + self.strPathCfgInit + " does not contains Environment section.")  

        if self.bTestVersionCheck == True:          
            if cfgINI.has_section('VersionCheckList') == True:
                print("\n========== Version Check Section ==========\n")
                cfgSection = "VersionCheckList"           
                for nIndex, strItemName in enumerate(cfgINI[cfgSection]):
                    strItemName = strItemName.upper()
                    self.dictVer[strItemName] = cfgINI[cfgSection][strItemName]
                    print('Name = ' + strItemName + " Value = " + cfgINI[cfgSection][strItemName])
            else:
                raise ValueError("File " + self.strPathCfgInit + " does not contains section of 'VersionCheckList'.")     

        if cfgINI.has_section('PreUninstall') == True:
            print("\n========== PreUninstall Section ==========\n")
            cfgSection = "PreUninstall" 
            for nIndex, strItemName in enumerate(cfgINI[cfgSection]):
                    strItemName = strItemName.upper()
                    self.dictPreUninstall[strItemName] = cfgINI[cfgSection][strItemName]
                    print('Name = ' + strItemName + " Value = " + cfgINI[cfgSection][strItemName])

    def mf_DoTest(self):
        '''Do EVT Testing
        '''	
        if self.strPkgOld == "" or self.strPkgNew == "":
            raise ValueError("Please do the function of mf_InitiEnv firstly.")

        if self.bTestUpgrade == True:
            self._mf_UpgradeTest()

        if self.bTestDowngrade == True:
            self._mf_DowngradeTest()       

        if self.bTestUninstallKeep == True:
            self._mf_UninstallKeepTest()       

        print("Test Finished.")

    def _mf_ResetTempFolder(self):
	    '''(Internal Purpose) Reset Temp Folder before Testing
        '''	
        strTemp = self.strCurrentPath + "\Temp"
        if os.path.isdir(strTemp) == True:
            shutil.rmtree(strTemp)
            os.mkdir(strTemp)
        else:
            os.mkdir(strTemp)

    def _mf_UpgradeTest(self):
		''' (Internal Purpose) Upgrade Test Procedure
        '''	
        print("\n\n========== UPGRADE TESTING ==========\n")
        self._mf_ResetTempFolder()
        self._mf_PreUninstall()
        if self.strPkgOldPDID == self.strPkgNewPDID:
            print("IS old package same as new.")
            print("Uninstall " + self.strPkgOldPDID)
            self._mf_UninstallPackage(self.strPkgOldPDID)
        else:
            print("IS old package is different to new.")
            print("Uninstall " + self.strPkgOldPDID)
            self._mf_UninstallPackage(self.strPkgOldPDID)
            print("Uninstall " + self.strPkgNewPDID)
            self._mf_UninstallPackage(self.strPkgNewPDID)

        print("\nInstall Before Stage")
        self.hContentOLD = open(self.strPathTempFileOld,'w')
        self.nLogMode = MODE_OLD
        self._mf_ExecuteProcess(self.strPkgOld, self.bSilentInstallOLD)
        self._mf_QueryRegINIList(self.strPathCfgReg)
        self.hContentOLD.close()

        print("\nInstall After Stage")
        self.hContentNEW = open(self.strPathTempFileNew, 'w')
        self.nLogMode = MODE_NEW
        self._mf_ExecuteProcess(self.strPkgNew, self.bSilentInstallNEW)
        self._mf_QueryRegINIList(self.strPathCfgReg)          
        self.hContentNEW.close()

        self._mf_ResultCheck(MODE_REPORT_UG)      

    def _mf_DowngradeTest(self):
		''' (Internal Purpose) Downgrade Test Procedure
        '''	
        print("\n\n========== DOWNGRADE TESTING ==========\n")
        self._mf_ResetTempFolder()
        self._mf_PreUninstall()
        if self.strPkgOldPDID == self.strPkgNewPDID:
            print("IS old package same as new.")
            print("Uninstall " + self.strPkgOldPDID)
            self._mf_UninstallPackage(self.strPkgOldPDID)
        else:
            print("IS old package is different to new.")
            print("Uninstall " + self.strPkgOldPDID)
            self._mf_UninstallPackage(self.strPkgOldPDID)
            print("Uninstall " + self.strPkgNewPDID)
            self._mf_UninstallPackage(self.strPkgNewPDID)

        print("\nInstall Before Stage")
        self.hContentOLD = open(self.strPathTempFileNew,'w')
        self.nLogMode = MODE_OLD
        self._mf_ExecuteProcess(self.strPkgNew, self.bSilentInstallNEW)
        self._mf_QueryRegINIList(self.strPathCfgReg)
        self.hContentOLD.close()

        print("\nInstall After Stage")
        self.hContentNEW = open(self.strPathTempFileOld, 'w')
        self.nLogMode = MODE_NEW
        self._mf_ExecuteProcess(self.strPkgOld, self.bSilentInstallOLD)
        self._mf_QueryRegINIList(self.strPathCfgReg)
        self.hContentNEW.close()            

        self._mf_ResultCheck(MODE_REPORT_DG)

    def _mf_UninstallKeepTest(self):
		''' (Internal Purpose) Uninstall Keep Test Procedure
        '''	
        print("\n\n========== UNINSTALL KEEP TESTING ==========\n")
        self._mf_ResetTempFolder()
        self._mf_PreUninstall()
        if self.strPkgOldPDID == self.strPkgNewPDID:
            print("IS old package same as new.")
            print("Uninstall " + self.strPkgOldPDID)
            self._mf_UninstallPackage(self.strPkgOldPDID)
        else:
            print("IS old package is different to new.")
            print("Uninstall " + self.strPkgOldPDID)
            self._mf_UninstallPackage(self.strPkgOldPDID)
            print("Uninstall " + self.strPkgNewPDID)
            self._mf_UninstallPackage(self.strPkgNewPDID)

        print("\nInstall Before Stage")
        self.hContentOLD = open(self.strPathTempFileOld,'w')
        self.nLogMode = MODE_OLD
        self._mf_ExecuteProcess(self.strPkgOld, self.bSilentInstallOLD)
        self._mf_QueryRegINIList(self.strPathCfgReg)
        self.hContentOLD.close()

        print("\nInstall After Stage")
        self.hContentNEW = open(self.strPathTempFileNew, 'w')
        self.nLogMode = MODE_NEW
        self._mf_ExecuteProcess(self.strPkgNew, self.bSilentInstallNEW)
        print("\nUninstall After Stage")    
        self._mf_UninstallPackage(self.strPkgNewPDID) 
        self._mf_QueryRegINIList(self.strPathCfgReg)
        self.hContentNEW.close()

        self._mf_ResultCheck(MODE_REPORT_KEEP)

    def _mf_PreUninstall(self):
		''' (Internal Purpose) According CfgInit.ini of configuration to uninstall whole of package before testing.
        '''	
        if len(self.dictPreUninstall) > 0:
            for strName in self.dictPreUninstall:
                print("Pre Uninstall " + strName)
                self._mf_UninstallPackage(self.dictPreUninstall[strName])

    def _mf_UninstallPackage(self, strPDID, bSilent = True):
        '''(Internal Purpose) Uninstall InstallShield Package.
        '''	
        print ("Uninstall ProductID = " + strPDID)

        strPDID = strPDID.replace('"','')
        if bSilent == True:
            strExecute = "msiexec /x " + strPDID + " /qn REBOOT=ReallySupress"
        else:
            strExecute = "msiexec /x " + strPDID + " /qr REBOOT=ReallySupress"

        print("Uninstall command = " + strExecute)
        pTemp = subprocess.Popen(strExecute, stdout=subprocess.PIPE)	  
        readContent = str(pTemp.communicate()[0])					
        #print(readContent)
        print("Uninstall Finished.")

    def _mf_ExecuteProcess(self, strPath, bSilent = False):      
        '''(Internal Purpose) Install InstallShield Package.
        '''	  
        print ("Execute File = " + strPath)

        if bSilent == True:
            strExecute = strPath + " /s /v\"/qn /norestart\""
        else:
            strExecute = strPath + " /s /v\"/qr /norestart\""

        pTemp = subprocess.Popen(strExecute, stdout=subprocess.PIPE)	  
        readContent = str(pTemp.communicate()[0])					
        print(readContent)

    def _mf_QueryRegINIList(self, strFilePath):
        '''(Internal Purpose) Query IS Uninstall Key Property
			
		strFilePath: Reg list file of path
		'''    
        if os.path.isfile(strFilePath):
            print("file " + strFilePath + " exist")
        else:
            raise ValueError("file " + strFilePath + " does not exist")

        cfgINI = configparser.ConfigParser()
        fileOP = open(strFilePath, 'r', encoding='utf-8')
        cfgINI.read_file(fileOP)

        for secIndex, cfgSection in enumerate(cfgINI):
            print('section = ' + str(cfgSection))
            if cfgSection != 'DEFAULT':                
                strPDKey = cfgINI[cfgSection]['ProductID']
                strHaveMSIX86 = int(str(cfgINI[cfgSection]['HaveMSIx86']).replace('"',''))
                strHaveMSIX64 = int(str(cfgINI[cfgSection]['HaveMSIx64']).replace('"',''))
                if self.nLogMode == MODE_OLD:
                    self._mf_QueryRegProperty(cfgSection, strPDKey, strHaveMSIX64, self.hContentOLD)
                else:
                    self._mf_QueryRegProperty(cfgSection, strPDKey, strHaveMSIX64, self.hContentNEW)
                print("")

    def _mf_QueryRegProperty(self, strSectionName, strPDKey, bHaveMSIx64, hContent):
        '''(Internal Purpose) Query IS Uninstall Key Property
			
		strPDKey: IS Product Key
        bHaveMSIx64: IS MSI has x64 bits packages
		'''        
        strRegKeyPath = STR_PATH_REG_UNINST + strPDKey.replace('"','')
        strRegKeyPathSysWow = STR_PATH_REG_UNINST_SYSWOW + strPDKey.replace('"','')
        
        try:
            if self.Is64bitOS:
                if bHaveMSIx64 == True:
                    print("OS: X64, IS_MSI: X64")
                    print(strRegKeyPath)
                    regKey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, strRegKeyPath, access=winreg.KEY_READ | winreg.KEY_WOW64_64KEY)
                else:
                    print("OS: X64, IS_MSI: X86")
                    print(strRegKeyPathSysWow)
                    regKey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, strRegKeyPathSysWow)            
            else:
                print("OS: X86, IS_MSI: X86")
                print(strRegKeyPath)
                regKey = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, strRegKeyPath)
                
            strDisplayName      = winreg.QueryValueEx(regKey, "DisplayName")
            strDisplayVersion   = winreg.QueryValueEx(regKey, "DisplayVersion")
            strInstallDate      = winreg.QueryValueEx(regKey, "InstallDate")
            strPublisher        = winreg.QueryValueEx(regKey, "Publisher")
            strUninstallString  = winreg.QueryValueEx(regKey, "UninstallString")

            print(strDisplayName)
            print(strDisplayVersion)
            print(strInstallDate)
            print(strPublisher)
            print(strUninstallString)

            hContent.write("[" + strSectionName + "]\n");
            hContent.write("DisplayName = \"" + strDisplayName[0] + "\"\n"); 
            hContent.write("DisplayVersion = \"" + strDisplayVersion[0] + "\"\n"); 
            hContent.write("InstallDate = \"" + strInstallDate[0] + "\"\n");
            hContent.write("Publisher = \"" + strPublisher[0] + "\"\n");
            hContent.write("UninstallString = \"" + strUninstallString[0] + "\"\n");
            hContent.write("\n")

        except FileNotFoundError:
            print(strPDKey + " does not exist in system.")
            raise ValueError(strPDKey + " does not exist in system.")

    def _mf_ResultCheck(self, nMode):
        '''(Internal Purpose) Version Check.
        '''	
        print("Doing version check between old and new installation content.")
        
        cfgOldIni = configparser.ConfigParser()
        cfgNewIni = configparser.ConfigParser()

        fileOld = open(self.strPathTempFileOld, 'r', encoding='utf-8')
        fileNew = open(self.strPathTempFileNew, 'r', encoding='utf-8')

        cfgOldIni.read_file(fileOld)
        cfgNewIni.read_file(fileNew)

        hTemp = ""
        if nMode == MODE_REPORT_UG:
            hTemp = open(self.strPathReportUG,'w')
        elif nMode == MODE_REPORT_DG:
            hTemp = open(self.strPathReportDG,'w')
        elif nMode == MODE_REPORT_KEEP:
            hTemp = open(self.strPathReportKeep,'w')
        else:
            raise ValueError("Test Mode " + str(nMode) + " Unsupported")

        # Record System Environment
        hTemp.write("========== System Environment ==========\n")
        for strItemName in self.dictSystemEnv:
            hTemp.write(strItemName.upper() + "\t= " + self.dictSystemEnv[strItemName] + "\n")

        hTemp.write("\n========== Test Condition ==============\n")
        # Record Test Condition
        hTemp.write("Before Installer = " + self.strPkgOld + "\n")
        hTemp.write("Product Code     = " + self.strPkgOldPDID + "\n")
        hTemp.write("After  Installer = " + self.strPkgNew + "\n")
        hTemp.write("Product Code     = " + self.strPkgNewPDID + "\n")

        hTemp.write("\n========== Test Result =================\n")
        bResult = True
        for secIndexOld, cfgSectionOld in enumerate(cfgOldIni):
            if cfgSectionOld == 'DEFAULT':
                continue

            for secIndex, cfgSectionNew in enumerate(cfgNewIni):
                if cfgSectionOld == 'DEFAULT':
                    continue    

                if cfgSectionOld == cfgSectionNew:
                    print('Compare ' + cfgSectionOld + " and " + cfgSectionNew)
                    strTempOld = str(cfgOldIni[cfgSectionOld]['DisplayVersion']).replace('"','')
                    nTempOld = int(strTempOld.replace('.',''))

                    strTempNew = str(cfgNewIni[cfgSectionNew]['DisplayVersion']).replace('"','')
                    nTempNew = int(strTempNew.replace('.',''))

                    strTemp = "Test Case: [" + cfgSectionNew + "]\n"
                    if nMode == MODE_REPORT_UG:
                        strTemp += "\tUpgrade Test   = "
                        if nTempNew > nTempOld:
                            strTemp += "PASS"
                        else:
                            strTemp += "FAIL"
                            bResult = False
                        strTemp += " (Before Version = " + strTempOld + ", After Version = " + strTempNew + ")\n"
                    elif nMode == MODE_REPORT_DG:
                        strTemp += "\tDowngrade Test = "
                        if nTempNew == nTempOld:
                            strTemp += "PASS"
                        else:
                            strTemp += "FAIL"
                            bResult = False
                        strTemp += " (Before Version = " + strTempOld + ", After Version = " + strTempNew + ")\n"
                    elif nMode == MODE_REPORT_KEEP:
                        strTemp += "\tKeep Test      = "
                        if nTempNew >= nTempOld or nTempNew == nTempOld:
                            strTemp += "PASS"
                        else:
                            strTemp += "FAIL"
                            bResult = False
                        strTemp += " (Before Version = " + strTempOld + ", After Version = " + strTempNew + ")\n"

                    print(strTemp)
                    hTemp.write(strTemp)
                    if self.bTestVersionCheck == True:
                        bTemp = self._mf_VersionCheck(cfgSectionNew, strTempNew, hTemp)
                        if bTemp == False:
                            bResult = False

        fileOld = open(self.strPathTempFileOld, 'r', encoding='utf-8')
        fileNew = open(self.strPathTempFileNew, 'r', encoding='utf-8')
        
        hTemp.write("\n========== Before Record Reference =====\n")
        hTemp.write(fileOld.read())
        hTemp.write("=========== After Record Reference ======\n")
        hTemp.write(fileNew.read())
        hTemp.write("\n")
        hTemp.close()

        fileOld.close()
        fileNew.close()  
        if bResult == False:
            raise ValueError("Test Result = FAIL. Please refer to the above logs for verification.")

    def _mf_VersionCheck(self, strName, strVersion, hFile):
        '''(Internal Purpose) Partial IS MSI Version Check.
        '''	
        bResult = True
        strLog = ""
        if strName in self.dictVer:
            strTemp = self.dictVer[strName]
            strLog = "\tVersion Check  = "
            strTemp = strTemp.replace('"','')
            strVersion = strVersion.replace('"','')

            if strVersion == strTemp:
                strLog += "PASS"
            else:
                strLog += "FAIL"
                bResult = False
            strLog += " (Read Version = " + strVersion + ", Expected Version = " +  strTemp + ")"
        else:
            strLog += "Version Check  = Fail. (" + strName + " does not supported)"
            bResult = False

        print(strLog)
        strLog += "\n"
        hFile.write(strLog)
        return bResult