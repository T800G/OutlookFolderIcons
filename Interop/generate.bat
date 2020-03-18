@ECHO OFF
tlbimp.exe ^
stdole2.tlb ^
/out:Microsoft.Interop.Stdole.dll ^
/keyfile:t800.snk ^
/strictref:nopia /nologo /asmversion:1.0.0.0 /sysarray

PAUSE


tlbimp.exe ^
msaddndr_14.tlb ^
/out:Microsoft.Office.Interop.Extensibility14.dll ^
/keyfile:t800.snk ^
/strictref:nopia /nologo /asmversion:1.0.0.0 /sysarray

PAUSE


tlbimp.exe ^
mso_14.tlb ^
/out:Microsoft.Office.Interop.Office14.dll ^
/keyfile:t800.snk ^
/strictref:nopia /nologo /asmversion:1.0.0.0 ^
/reference:Microsoft.Interop.Stdole.dll

PAUSE


tlbimp.exe ^
vbe6ext_14.olb ^
/out:Microsoft.Office.Interop.VBAExtensibility14.dll ^
/keyfile:t800.snk ^
/strictref:nopia /nologo /asmversion:1.0.0.0 ^
/reference:Microsoft.Office.Interop.Office14.dll

PAUSE

tlbimp.exe ^
msoutl_14.olb ^
/out:Microsoft.Office.Interop.Outlook14.dll ^
/keyfile:t800.snk ^
/strictref:nopia /nologo /asmversion:1.0.0.0 ^
/reference:Microsoft.Interop.Stdole.dll ^
/reference:Microsoft.Office.Interop.Office14.dll
 
PAUSE