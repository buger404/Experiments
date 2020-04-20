Attribute VB_Name = "MachineCode"
'====================================================================================
'   Machine Code
'   Made: Error 404(1361778219)
'   Version: 1.0(20.01.27)
'   Describe: 用于获取设备特征码（MAC地址）
'   Note: 需要依赖MD5.cls
'====================================================================================
    '返回设备特征码
    Function GetMachineCode() As String
        Dim o As Object, o2 As Object, r As String
        
        Set o = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=TRUE")
        For Each o2 In o
            r = r & o2.macaddress & vbCrLf
        Next
        
        Set o = Nothing
        Set o = New MD5
        GetMachineCode = o.Md5_String_Calc(r)
    End Function
'====================================================================================
