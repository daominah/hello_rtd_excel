﻿# Excel RTD demo
Create a RTD DLL for receiving real time data in Excel.  
Relative project [hello_electron](https://bitbucket.tradex.vn/scm/~tungdt/hello_electron.git)

### Create a RTD DLL
* Create new project C# Class Library .Net Framework in 
[VS2019](https://visualstudio.microsoft.com/thank-you-downloading-visual-studio/?sku=Community&rel=16)
* Install NuGet package Microsoft.Office.Interop.Excel
* Implements interface Microsoft.Office.Interop.Excel.IRtdServer (
    [Create a RealTimeData server for Excel](https://docs.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/create-realtimedata-server-in-excel))
* Project > Properties > Signing: sign the assembly
* Build the solution, output is a DLL in `bin\Debug`
* Run admin cmd to 
    [register the built C# RTD DLL ](https://stackoverflow.com/questions/58018613/compiling-an-irtdserver-interface-for-excel-rtd-function-in-net-core) ([RegAsm path](https://stackoverflow.com/a/48393717/4097963)):  
    * .NET 4.x: `C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe /codebase PathToBuiltDLL`  
    * .NET 2.0, 3.0, and 3.5: `C:\Windows\Microsoft.NET\Framework64\v2.0.50727\RegAsm.exe /codebase PathToBuiltDLL`
* Write VBA functions to receive real time data: 
`WorksheetFunction.RTD(dLLName, vbNullString, topic, 1)`.  
Example is file `TestRTD0.xlsm` in this dir.

### Add VBA functions description
[Stackoverflow question](https://stackoverflow.com/questions/4262421/how-to-put-a-tooltip-on-a-user-defined-function)
* [Method 1](https://docs.microsoft.com/en-us/office/vba/api/excel.application.macrooptions) 
(built-in by Microsoft, need to press `Ctrl A` to show the description): 
implement `Workbook_Open()` in the `VBAProject/ThisWorkBook`. Example:  
````
Private Sub Workbook_Open()
    funcName = "MyFunc0"
    funcDesc = "MyFunc0 returns sum of the inputs"
    Dim args(1 To 2)
    args(1) = "the first argument description"
    args(2) = "hihi pussy, fuck Microsoft codes anyway"
    Application.MacroOptions Macro:=funcName, Description:=funcDesc, ArgumentDescriptions:=args, Category:=14
End Sub
````
* [Method 2](https://github.com/Excel-DNA/IntelliSense/wiki/Getting-Started)
(user have to install a XLL add-in):  
    * Create a new sheet with the name "\_IntelliSense\_" that save function's 
    descriptions. Example in the `TestRTD0.xlsm`.
    * In Excel, press `Alt T, I` to register file `ExcelDna.IntelliSense64.xll`.
    * The add-in can be register by `regsvr32.exe` (
        [register DLL on windows 64-bit](https://stackoverflow.com/questions/4897685/how-do-i-register-a-dll-file-on-windows-7-64-bit))
