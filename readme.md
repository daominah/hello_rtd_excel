# Excel RTD demo
Create a DLL for receiving real time data in Excel.  


### Steps
* Create new project C# Class Library .Net Framework in [VS2019](https://visualstudio.microsoft.com/thank-you-downloading-visual-studio/?sku=Community&rel=16)
* Install NuGet package Microsoft.Office.Interop.Excel
* Implements interface Microsoft.Office.Interop.Excel.IRtdServer
* Project > Properties > Signing: sign the assembly
* Build the solution, output is a DLL in `bin\Debug`
* Run admin cmd:  
`C:\Windows\Microsoft.NET\Framework64\v4.0.30319>RegAsm.exe PathToBuiltDLL /codebase`  
* Write VBA functions to receive real time data: `WorksheetFunction.RTD(dLLName, vbNullString, topic, 1)`. Example is file `TestRTD0.xlsm` in this dir.

### References
* [Create a RealTimeData server for Excel](https://docs.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/create-realtimedata-server-in-excel)  
* [Register the built DLL](https://stackoverflow.com/questions/58018613/compiling-an-irtdserver-interface-for-excel-rtd-function-in-net-core)
