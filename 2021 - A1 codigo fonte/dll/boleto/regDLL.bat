rem regsvr32 impactro.cobranca.dll
REM REMOVENDO
REM regsvr32 /u impactro.cobranca.dll

REM Registrar DLL
"C:\windows\microsoft.net\framework\v2.0.50727\regasm.exe" /TLB impactro.cobranca.dll
REM Remover registro
rem "C:\windows\microsoft.net\framework\v2.0.50727\regasm.exe" /unregister impactro.cobranca.dll
 pause