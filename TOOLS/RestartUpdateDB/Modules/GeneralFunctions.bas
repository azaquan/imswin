Attribute VB_Name = "GeneralFunctions"
 Public Const PROCESS_QUERY_INFORMATION = 1024
    Public Const PROCESS_VM_READ = 16
      Public Const MAX_PATH = 260
      Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
      Public Const SYNCHRONIZE = &H100000
      
      Public Const PROCESS_ALL_ACCESS = &H1F0FFF
      Public Const TH32CS_SNAPPROCESS = &H2&
      Public Const hNull = 0
 Public Declare Function Process32First Lib "kernel32" ( _
         ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

      Public Declare Function Process32Next Lib "kernel32" ( _
         ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long

      Public Declare Function CloseHandle Lib "Kernel32.dll" _
         (ByVal Handle As Long) As Long

      Public Declare Function OpenProcess Lib "Kernel32.dll" _
        (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, _
            ByVal dwProcId As Long) As Long

      Public Declare Function EnumProcesses Lib "psapi.dll" _
         (ByRef lpidProcess As Long, ByVal cb As Long, _
            ByRef cbNeeded As Long) As Long

      Public Declare Function GetModuleFileNameExA Lib "psapi.dll" _
         (ByVal hProcess As Long, ByVal hModule As Long, _
            ByVal ModuleName As String, ByVal nSize As Long) As Long

      Public Declare Function EnumProcessModules Lib "psapi.dll" _
         (ByVal hProcess As Long, ByRef lphModule As Long, _
            ByVal cb As Long, ByRef cbNeeded As Long) As Long

      Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" ( _
         ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long

      Public Declare Function GetVersionExA Lib "kernel32" _
         (lpVersionInformation As OSVERSIONINFO) As Integer

      Public Type PROCESSENTRY32
         dwSize As Long
         cntUsage As Long
         th32ProcessID As Long
         th32DefaultHeapID As Long
         th32ModuleID As Long
         cntThreads As Long
         th32ParentProcessID As Long
         pcPriClassBase As Long
         dwFlags As Long
         szExeFile As String * 260
      End Type

      Public Type OSVERSIONINFO
         dwOSVersionInfoSize As Long
         dwMajorVersion As Long
         dwMinorVersion As Long
         dwBuildNumber As Long
         dwPlatformId As Long
                                        

         szCSDVersion As String * 128
      End Type

Public Type fileinfo

    WaitTime As Integer
    ExeFullPath As String
    Email As String
    
End Type


Public Function GetIniInformation() As fileinfo

Dim sa As Scripting.FileSystemObject

Dim t As Scripting.TextStream

Dim line As String

Dim location As Integer


Dim mfileinfo As fileinfo

Dim Filepath As String

Set sa = New Scripting.FileSystemObject

Filepath = App.Path & "\UpdateDbRestart.txt"

If sa.FileExists(Filepath) = False Then

    MsgBox "There is no INI file associated with this program. Please create one."

    Exit Function
    
End If
 
Set t = sa.OpenTextFile(Filepath)

   line = t.ReadLine
    
   location = InStr(line, "=")
    
   mfileinfo.WaitTime = Mid(line, location + 1, Len(line))
   
    line = t.ReadLine
    
   location = InStr(line, "=")
    
   mfileinfo.ExeFullPath = Mid(line, location + 1, Len(line))
       
   line = t.ReadLine
    
   location = InStr(line, "=")
    
   mfileinfo.Email = Mid(line, location + 1, Len(line))
    

   t.Close
   
   Set t = Nothing

GetIniInformation = mfileinfo

End Function

      Function StrZToStr(s As String) As String
         StrZToStr = Left$(s, Len(s) - 1)
      End Function

      Public Function getVersion() As Long
         Dim osinfo As OSVERSIONINFO
         Dim retvalue As Integer
         osinfo.dwOSVersionInfoSize = 148
         osinfo.szCSDVersion = Space$(128)
         retvalue = GetVersionExA(osinfo)
         getVersion = osinfo.dwPlatformId
      End Function

Public Function isActive(Proceso As String) As Boolean
Select Case getVersion()
      Case 1

         Dim f As Long
         Dim sname As String
         Dim hSnap As Long
         Dim proc As PROCESSENTRY32
         hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
         If hSnap = hNull Then Exit Function
         proc.dwSize = Len(proc)
         
         f = Process32First(hSnap, proc)
         Do While f
           sname = StrZToStr(proc.szExeFile)
           If Proceso = sname Then
                isActive = True
                Exit Function
            End If
           f = Process32Next(hSnap, proc)
         Loop
      Case 2
         Dim cb As Long
         Dim cbNeeded As Long
         Dim NumElements As Long
         Dim ProcessIDs() As Long
         Dim cbNeeded2 As Long
         Dim NumElements2 As Long
         Dim Modules(1 To 200) As Long
         Dim lRet As Long
         Dim ModuleName As String
         Dim nSize As Long
         Dim hProcess As Long
         Dim i As Long
         
         cb = 8
         cbNeeded = 96
         Do While cb <= cbNeeded
            cb = cb * 2
            ReDim ProcessIDs(cb / 4) As Long
            lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
         Loop
         NumElements = cbNeeded / 4

         For i = 1 To NumElements
            
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
               Or PROCESS_VM_READ, 0, ProcessIDs(i))
            
            If hProcess <> 0 Then
                
                lRet = EnumProcessModules(hProcess, Modules(1), 200, _
                                             cbNeeded2)
                
                If lRet <> 0 Then
                   ModuleName = Space(MAX_PATH)
                   nSize = 500
                   lRet = GetModuleFileNameExA(hProcess, Modules(1), _
                                   ModuleName, nSize)
                   If Proceso = Left(ModuleName, lRet) Then
                        isActive = True
                        Exit Function
                    End If
                End If
            End If
          
         lRet = CloseHandle(hProcess)
         Next

      End Select
      isActive = False
End Function



