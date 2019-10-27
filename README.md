### Description
---
Just a simple script to generate JScript code for calling Win32 API functions using XLM/Excel 4.0 macros via Excel.Application COM object and "ExecuteExcel4Macro" method.

### Usage
---
  ``-o string
    	output payload filename``<br>
  ``-sh string
    	Shellcode file path, ex: go run genXLM.go -sh shellcode.bin``<br>
  ``-wsh string
    	payload template js/hta, ex: go run genXLM.go -sh shellcode.bin -wsh js``<br>

### Detection:
---
Currently not detected on VT;
  * (0/56) VT https://www.virustotal.com/gui/file/f5a67b22f0362403b851664b6edd25928383d7f68099b61612e580b94734fe7a/detection

  * XLM macros are not being covered by AMSI scans
  * Instantiating Excel.Application COM objects from JS/VBS and calling ExecuteExcel4Macro is not flagged by WinDefender/AMSI 

### Details:
---
Generate a simple JS using ``go run genXLM.go -sh shellcode.bin -wsh js`` and have a look at the generated js code "self-descriptive". 

Check calc.hta, calc.js for examples. shellcode was generated using msvenom.

### References:
---
* https://outflank.nl/blog/2018/10/06/old-school-evil-excel-4-0-macros-xlm/

### Disclaimer:
---
Use it for authorized red teaming and/or nonprofit educational purposes only. Any misuse of this script will not be the responsibility of the author. Use it at your own networks and/or with the network owner's permission.

