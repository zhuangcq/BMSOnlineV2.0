Module Module1
    Public TheFileName As String
    Public TheNewPath As String
    Public FileInfoTitle As String
    Public Z, Z1, Z0, Z2, Z3, Z4, Z5, Z6, Z7, Z8, Z9 As Single
    Public Z11, Z10, Z12, Z13, Z14, Z15, Z16, Z17, Z18, Z19 As Single
    Public Z21, Z20, Z22, Z23, Z24, Z25, Z26, Z27, Z28, Z29 As Single
    Public Z31, Z30, Z32, Z33, Z34, Z35, Z36, Z37, Z38, Z39 As Single
    Public Z41, Z40, Z42, Z43, Z44, Z45, Z46, Z47, Z48, Z49 As Single
    Public Z51, Z50, Z52, Z53, Z54, Z55, Z56, Z57, Z58, Z59 As Single
    Public Z100, Rm As Single

    Public Vmax, Vmin, TK As Single
    Public Vmax1, Vmin1, Vmax2, Vmin2 As Single

    Public Counter, Countera As Integer
    Public Counter1, Counterb As Integer
    Public e, e1 As Integer
    Public R, Ra, Ra1, Ro1, Rb, W, R1, M, n, m1, m2, m3, m4, m5 As Integer
    Public Ho, V11o, Hn, V11n, Ver1, Ver1o, Ver2, Ver2o As Single
    Public Ver3, Ver3o, Ver4, Ver4o, Ver5, Ver5o As Single
    Public Ver1b, Ver1bo, Ver2b, Ver2bo, Ver3b, Ver3bo As Single
    Public Ver4b, Ver4bo, Ver5b, Ver5bo As Single
    Public Ver1c, Ver1oc, Ver2c, Ver2oc, Ver3c, Ver3oc As Single
    Public Ver5c, Ver5oc, Ver4c, Ver4oc, Ver, Vero As Single
    Public V1n, Ho1, Hn1 As Single
    Public l(59) As Single
    Public Se, Index1, GA As Single
    Public Flag(5), Flag1(5) As Integer
    Public n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11 As Object
    Public period As Single
    Public Index As Integer
    Public Mb(2), N17, L7 As Integer
    Public Fcurve As Integer

    Declare Function CreateFileNS Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
    Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
    Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

    Public iTask As Long

    Public Const GW_HWNDFIRST = 0
    Public Const GW_HWNDLAST = 1
    Public Const GW_HWNDNEXT = 2
    Public Const GW_HWNDPREN = 3
    Public Const GW_OWNER = 4
    Public Const GW_CHILD = 5
    Public hact, Now As Long ' handle to the active window
    Public retval As Long  ' return value
    'Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    Declare Function GetActiveWindow Lib "user32" () As Long
    Declare Function FlashWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
    Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
    'Declare Function ExitWindowsEx Lib "user32.dll" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
    Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function GetFocus Lib "user32" () As Long
    Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
    Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
    Declare Function EndDeferWindowPos Lib "user32" (ByVal hWinPosInfo As Long) As Long
    'Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
    Declare Function CloseWindowStation Lib "user32" (ByVal hWinSta As Long) As Boolean
    Declare Function GetProcessShutdownParameters Lib "kernel32" (ByVal lpdwLevel As Long, ByVal lpdwFlags As Long) As Long

    Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Public Const INFINITE = -1&
    Public Const SYNCHRONIZE = &H100000

    Public Sub order()
        Flag(1) = 0
        Flag(2) = 0
        Flag(3) = 0
        Flag(4) = 0
        Flag(5) = 0
        ' Dim E As String
        '  Dim E1 As String
        Countera = Counter
        If Z11 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 11
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Aver.P"
            curves.Panel2.Controls.Item(Index).Text = "Average Pollutant Concentration of Zones(ppm)"
        Else
        End If
        If Z12 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 12
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "VAVF.Z1"
            curves.Panel2.Controls.Item(Index).Text = "VAV Flow rate of Zone1(kg/s)"
        Else
        End If
        If Z13 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 13
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "VAVF.Z2"
            curves.Panel2.Controls.Item(Index).Text = "VAV Flow rate of Zone2(kg/s)"
        Else
        End If
        If Z14 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 14
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "FlowZ1"
            curves.Panel2.Controls.Item(Index).Text = "Air Flow setpoint of Zone1 "
        Else
        End If
        If Z15 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 15
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "FlowZ2"
            curves.Panel2.Controls.Item(Index).Text = "Air Flow setpoint of Zone2 "
        Else
        End If
        If Z16 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 16
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "FlowZ3"
            curves.Panel2.Controls.Item(Index).Text = "Air Flow setpoint of Zone3 "
        Else
        End If
        If Z17 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 17
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "FlowZ4"
            curves.Panel2.Controls.Item(Index).Text = "Air Flow setpoint of Zone4 "
        Else
        End If
        If Z18 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 18
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "FlowZ5"
            curves.Panel2.Controls.Item(Index).Text = "Air Flow setpoint of Zone5 "
        Else
        End If
        If Z19 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 19
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "FlowZ6"
            curves.Panel2.Controls.Item(Index).Text = "Air Flow setpoint of Zone6 "
        Else
        End If
        If Z20 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 20
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "FlowZ7"
            curves.Panel2.Controls.Item(Index).Text = "Air Flow setpoint of Zone7 "
        Else
        End If
        If Z21 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 21
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "FlowZ8"
            curves.Panel2.Controls.Item(Index).Text = "Air Flow setpoint of Zone8 "
        Else
        End If
        If Z22 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 22
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Dam.Z1"
            curves.Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone1 "
        Else
        End If
        If Z23 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 23
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Dam.Z2"
            curves.Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone2 "
        Else
        End If
        If Z24 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 24
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Dam.Z3"
            curves.Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone3 "
        Else
        End If
        If Z25 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 25
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Dam.Z4"
            curves.Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone4 "
        Else
        End If
        If Z26 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 26
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Dam.Z5"
            curves.Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone5 "
        Else
        End If
        If Z27 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 27
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Dam.Z6"
            curves.Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone6 "
        Else
        End If
        If Z28 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 28
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Dam.Z7"
            curves.Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone7 "
        Else
        End If
        If Z29 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 29
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Dam.Z8"
            curves.Panel2.Controls.Item(Index).Text = "VAV damper position demand of Zone8 "
        Else
        End If
        If Z32 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 32
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "CAVcoil"
            curves.Panel2.Controls.Item(Index).Text = "CAV cooling coil demand "
        Else
        End If
        If Z34 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 34
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "VAVcoil"
            curves.Panel2.Controls.Item(Index).Text = "VAV cooling coil demand "
        Else
        End If
        If Z37 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 37
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Sup.Fan"
            curves.Panel2.Controls.Item(Index).Text = "VAV supply fan control"
        Else
        End If
        If Z38 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 38
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Sup.F"
            curves.Panel2.Controls.Item(Index).Text = "Total supply air flow rate "
        Else
        End If
        If Z39 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 39
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Retu.F"
            curves.Panel2.Controls.Item(Index).Text = "Total return air flow rate(kg/s)"
        Else
        End If
        If Z40 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 40
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Re.Fan"
            curves.Panel2.Controls.Item(Index).Text = "Return fan control demand "
        Else
        End If
        If Z41 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 41
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Fres.F"
            curves.Panel2.Controls.Item(Index).Text = "Flow rate of fresh air(m3/s) "
        Else
        End If
        If Z42 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 42
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Fre.Fset"
            curves.Panel2.Controls.Item(Index).Text = "Fresh air flow setpoint (m3/s) "
        Else
        End If
        If Z43 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 43
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Out.Dam"
            curves.Panel2.Controls.Item(Index).Text = "Outdoor air damper position demand "
        Else
        End If
        If Z51 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 51
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Wat.V"
            curves.Panel2.Controls.Item(Index).Text = "Water valve control of VAV coil"
        Else
        End If
        If Z52 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 52
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Wat.C"
            curves.Panel2.Controls.Item(Index).Text = "Water valve of CAV coil"
        Else
        End If
        If Z9 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 9
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Aver.H"
            curves.Panel2.Controls.Item(Index).Text = "Average humidity of Zones (kg/kg)"
        Else
        End If
        If Z46 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 46
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Retu.H"
            curves.Panel2.Controls.Item(Index).Text = "Humidity of return air(kg/kg)"
        Else
        End If
        If Z47 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 47
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Fres.H"
            curves.Panel2.Controls.Item(Index).Text = "Humidity of fresh air(kg/kg)"
        Else
        End If
        If Z35 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 35
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Sup.P"
            curves.Panel2.Controls.Item(Index).Text = "VAV supply pressure(pa) "
        Else
        End If
        If Z36 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 36
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "P.set"
            curves.Panel2.Controls.Item(Index).Text = "Static pressure setpoint(pa)"
        Else
        End If
        If Z10 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 10
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Ave.CO2"
            curves.Panel2.Controls.Item(Index).Text = "Average CO2 concentration of Zones(ppm)"
        Else
        End If
        If Z44 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 44
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Re.CO2"
            curves.Panel2.Controls.Item(Index).Text = "CO2 concentration of return air(ppm)"
        Else
        End If
        If Z49 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 49
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Fre.CO2"
            curves.Panel2.Controls.Item(Index).Text = "CO2 concentration of fresh air((ppm)"
        Else
        End If
        If Z50 = 1 Then
            Countera = Countera - 1
            E = Counter - Countera
            Flag(E) = 50
            Index = E - 1
            curves.Panel1.Controls.Item(Index).Text = "Number"
            curves.Panel2.Controls.Item(Index).Text = "Number of people "
        Else
        End If

        Flag1(1) = 0
        Flag1(2) = 0
        Flag1(3) = 0
        Flag1(4) = 0
        Flag1(5) = 0

        Counterb = Counter1

        If Z1 = 1 Then
            Counterb = Counterb - 1
            E1 = Counter1 - Counterb
            Flag1(E1) = 1
            Index = E1 - 1
            curves.Panel3.Controls.Item(Index).Text = "Tz1"
            curves.Panel4.Controls.Item(Index).Text = "Temperature of Zone 1 (°C)"
        Else
        End If
        If Z2 = 1 Then
            Counterb = Counterb - 1
            E1 = Counter1 - Counterb
            Flag1(E1) = 2
            Index = E1 - 1
            curves.Panel3.Controls.Item(Index).Text = "Tz2"
            curves.Panel4.Controls.Item(Index).Text = "Temperature of Zone 2 (°C)"
        Else
        End If
        If Z3 = 1 Then
            Counterb = Counterb - 1
            E1 = Counter1 - Counterb
            Flag1(E1) = 3
            Index = E1 - 1
            curves.Panel3.Controls.Item(Index).Text = "Tz3"
            curves.Panel4.Controls.Item(Index).Text = "Temperature of Zone 3 (°C)"
        Else
        End If
        If Z4 = 1 Then
            Counterb = Counterb - 1
            E1 = Counter1 - Counterb
            Flag1(E1) = 4
            Index = E1 - 1
            curves.Panel3.Controls.Item(Index).Text = "Tz4"
            curves.Panel4.Controls.Item(Index).Text = "Temperature of Zone 4 (°C)"
        Else
        End If
        If Z5 = 1 Then
            Counterb = Counterb - 1
            E1 = Counter1 - Counterb
            Flag1(E1) = 5
            Index = E1 - 1
            curves.Panel3.Controls.Item(Index).Text = "Tz5"
            curves.Panel4.Controls.Item(Index).Text = "Temperature of Zone 5 (°C)"
        Else
        End If
        If Z6 = 1 Then
            Counterb = Counterb - 1
            E1 = Counter1 - Counterb
            Flag1(E1) = 6
            Index = E1 - 1
            curves.Panel3.Controls.Item(Index).Text = "Tz6"
            curves.Panel4.Controls.Item(Index).Text = "Temperature of Zone 6 (°C)"
        Else
        End If
        If Z7 = 1 Then
            Counterb = Counterb - 1
            E1 = Counter1 - Counterb
            Flag1(E1) = 7
            Index = E1 - 1
            curves.Panel3.Controls.Item(Index).Text = "Tz7"
            curves.Panel4.Controls.Item(Index).Text = "Temperature of Zone 7 (°C)"
        Else
        End If
        If Z8 = 1 Then
            Counterb = Counterb - 1
            E1 = Counter1 - Counterb
            Flag1(E1) = 8
            Index = E1 - 1
            curves.Panel3.Controls.Item(Index).Text = "Tz8"
            curves.Panel4.Controls.Item(Index).Text = "Temperature of Zone 8 (°C)"
        Else
        End If
        If Z30 = 1 Then
            Counterb = Counterb - 1
            E1 = Counter1 - Counterb
            Flag1(E1) = 30
            Index = E1 - 1
            curves.Panel3.Controls.Item(Index).Text = "Tout.C"
            curves.Panel4.Controls.Item(Index).Text = "CAV cooling coil outlet temperature (°C)"
        Else
        End If
        If Z31 = 1 Then
            Counterb = Counterb - 1
            E1 = Counter1 - Counterb
            Flag1(E1) = 31
            Index = E1 - 1
            curves.Panel3.Controls.Item(Index).Text = "Tset.Sup"
            curves.Panel4.Controls.Item(Index).Text = "Supply air temperature setpoint(°C)"
        Else
        End If
        If Z33 = 1 Then
            Counterb = Counterb - 1
            E1 = Counter1 - Counterb
            Flag1(E1) = 33
            Index = E1 - 1
            curves.Panel3.Controls.Item(Index).Text = "Tout.V"
            curves.Panel4.Controls.Item(Index).Text = "VAV cooling coil outlet temperature(°C)"
        Else
        End If
        If Z45 = 1 Then
            Counterb = Counterb - 1
            E1 = Counter1 - Counterb
            Flag1(E1) = 45
            Index = E1 - 1
            curves.Panel3.Controls.Item(Index).Text = "T.retu"
            curves.Panel4.Controls.Item(Index).Text = "Temperature of return air (°C)"
        Else
        End If
        If Z48 = 1 Then
            Counterb = Counterb - 1
            E1 = Counter1 - Counterb
            Flag1(E1) = 48
            Index = E1 - 1
            curves.Panel3.Controls.Item(Index).Text = "T.fres"
            curves.Panel4.Controls.Item(Index).Text = "Temperature of fresh air (°C)"
        Else
        End If

    End Sub
End Module

