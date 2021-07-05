echo =============================================================================== >> c:\microsoft_keys.txt
echo Hostname >> c:\microsoft_keys.txt
hostname >> c:\microsoft_keys.txt
echo =============================================================================== >> c:\microsoft_keys.txt
echo Serial Number Bios >> c:\microsoft_keys.txt
WMIC BIOS GET SERIALNUMBER >> c:\microsoft_keys.txt
echo =============================================================================== >> c:\microsoft_keys.txt
echo Windows Serial Key >> c:\microsoft_keys.txt
wmic path softwarelicensingservice get OA3xOriginalProductKey >> c:\microsoft_keys.txt
echo =============================================================================== >> c:\microsoft_keys.txt
echo Office 2016/2019 (32-bit) on a 32-bit version of Windows
cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dstatus >> c:\microsoft_keys.txt
echo =============================================================================== >> c:\microsoft_keys.txt
echo Office 2016/2019 (32-bit) on a 64-bit version of Windows >> c:\microsoft_keys.txt
cscript "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS" /dstatus >> c:\microsoft_keys.txt
echo =============================================================================== >> c:\microsoft_keys.txt
echo Office 2016/2019 (64-bit) on a 64-bit version of Windows >> c:\microsoft_keys.txt
cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dstatus >> c:\microsoft_keys.txt
echo =============================================================================== >> c:\microsoft_keys.txt
echo Office 2013 (32-bit) on a 32-bit version of Windows >> c:\microsoft_keys.txt
cscript "C:\Program Files\Microsoft Office\Office15\OSPP.VBS" /dstatus >> c:\microsoft_keys.txt
echo =============================================================================== >> c:\microsoft_keys.txt
echo Office 2013 (32-bit) on a 64-bit version of Windows >> c:\microsoft_keys.txt
cscript "C:\Program Files (x86)\Microsoft Office\Office15\OSPP.VBS" /dstatus >> c:\microsoft_keys.txt
echo =============================================================================== >> c:\microsoft_keys.txt
echo Office 2013 (64-bit) on a 64-bit version of Windows >> c:\microsoft_keys.txt
cscript "C:\Program Files\Microsoft Office\Office15\OSPP.VBS" /dstatus >> c:\microsoft_keys.txt
echo =============================================================================== >> c:\microsoft_keys.txt
echo Office 2010 (32-bit) on a 32-bit version of Windows >> c:\microsoft_keys.txt
cscript "C:\Program Files\Microsoft Office\Office14\OSPP.VBS" /dstatus >> c:\microsoft_keys.txt
echo =============================================================================== >> c:\microsoft_keys.txt
echo Office 2010 (32-bit) on a 64-bit version of Windows >> c:\microsoft_keys.txt
cscript "C:\Program Files (x86)\Microsoft Office\Office14\OSPP.VBS" /dstatus >> c:\microsoft_keys.txt
echo =============================================================================== >> c:\microsoft_keys.txt
echo Office 2010 (64-bit) on a 64-bit version of Windows >> c:\microsoft_keys.txt
cscript "C:\Program Files\Microsoft Office\Office14\OSPP.VBS" /dstatus >> c:\microsoft_keys.txt
echo =============================================================================== >> c:\microsoft_keys.txt
