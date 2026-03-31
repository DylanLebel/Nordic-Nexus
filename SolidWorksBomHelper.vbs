Option Explicit

Dim gAssemblyPath, gOutPath
Dim gFSO, gParts, gSeen

Set gFSO = CreateObject("Scripting.FileSystemObject")
Set gParts = CreateObject("Scripting.Dictionary")
Set gSeen = CreateObject("Scripting.Dictionary")

gAssemblyPath = ""
gOutPath = ""

ParseArgs
InitOutput

If Trim(gAssemblyPath) = "" Then
    LogMsg "ERROR: Assembly path is empty."
    WScript.Quit 2
End If

LogMsg "VBS helper starting for: " & gAssemblyPath

Dim swApp, createdSw, model, openTitle
createdSw = False
openTitle = ""

On Error Resume Next
Set swApp = GetObject(, "SldWorks.Application")
If Err.Number <> 0 Then
    Err.Clear
    Set swApp = Nothing
End If
On Error GoTo 0

If swApp Is Nothing Then
    On Error Resume Next
    Set swApp = CreateObject("SldWorks.Application")
    If Err.Number <> 0 Then
        LogMsg "ERROR: Could not create SolidWorks COM object: " & Err.Description
        WScript.Quit 3
    End If
    On Error GoTo 0
    createdSw = True
    LogMsg "Created SolidWorks instance in VBS helper."
Else
    LogMsg "Attached to existing SolidWorks instance in VBS helper."
End If

On Error Resume Next
swApp.Visible = True
swApp.UserControl = True
On Error GoTo 0

Set model = OpenAssembly(swApp, gAssemblyPath)
If model Is Nothing Then
    LogMsg "ERROR: Failed to open assembly in VBS helper."
    Cleanup swApp, createdSw, openTitle
    WScript.Quit 4
End If

' Activate the target document so it is the active doc — ensures
' subsequent operations work on the correct assembly, not a parent.
On Error Resume Next
Dim activateErrV2: activateErrV2 = 0
swApp.ActivateDoc2 gAssemblyPath, False, activateErrV2
If Err.Number <> 0 Then Err.Clear
openTitle = model.GetTitle
model.ResolveAllLightWeightComponents True
On Error GoTo 0

Dim cfg, rootComp, asmBase, k
Set cfg = Nothing
Set rootComp = Nothing

On Error Resume Next
Set cfg = model.GetActiveConfiguration
If Not cfg Is Nothing Then Set rootComp = cfg.GetRootComponent3(True)
If rootComp Is Nothing And Not cfg Is Nothing Then Set rootComp = cfg.GetRootComponent
On Error GoTo 0

If rootComp Is Nothing Then
    LogMsg "ERROR: Root component unavailable."
    Cleanup swApp, createdSw, openTitle
    WScript.Quit 5
End If

Function FindMacroUpTree(ByVal startDir, ByVal macroName)
    Dim curDir, candidate, parentDir, hopCount
    FindMacroUpTree = ""
    curDir = Trim(startDir)
    hopCount = 0

    Do While curDir <> ""
        candidate = gFSO.BuildPath(curDir, macroName)
        If gFSO.FileExists(candidate) Then
            FindMacroUpTree = candidate
            Exit Function
        End If

        On Error Resume Next
        parentDir = gFSO.GetParentFolderName(curDir)
        If Err.Number <> 0 Then
            Err.Clear
            parentDir = ""
        End If
        On Error GoTo 0

        If parentDir = "" Then Exit Do
        If StrComp(parentDir, curDir, vbTextCompare) = 0 Then Exit Do

        curDir = parentDir
        hopCount = hopCount + 1
        If hopCount > 20 Then Exit Do
    Loop
End Function

' ---- Try RunMacro2 via NmtBomExtract.swp first (runs inside SW with early binding) ----
Dim macroPartCount
macroPartCount = 0
On Error Resume Next
Dim vbsScriptDir, swpPath, macroShell, macroArgsPath, macroOutPath
vbsScriptDir = gFSO.GetParentFolderName(WScript.ScriptFullName)
swpPath      = FindMacroUpTree(vbsScriptDir, "NmtBomExtract.swp")
Set macroShell  = CreateObject("WScript.Shell")
macroArgsPath   = macroShell.ExpandEnvironmentStrings("%TEMP%") & "\nmt_bom_args.txt"
macroOutPath    = macroShell.ExpandEnvironmentStrings("%TEMP%") & "\nmt_bom_result_v2_" & Replace(gFSO.GetTempName, ".", "_") & ".txt"
On Error GoTo 0

If gFSO.FileExists(swpPath) Then
    LogMsg "Using NmtBomExtract.swp from: " & swpPath
    On Error Resume Next
    Dim mafTs
    Set mafTs = gFSO.OpenTextFile(macroArgsPath, 2, True, 0)
    mafTs.Write "ASSEMBLY=" & gAssemblyPath & vbLf
    mafTs.Write "OUT=" & macroOutPath & vbLf
    mafTs.Close
    Err.Clear

    ' Try RunMacro2 first (5-arg form)
    LogMsg "Calling RunMacro2 -> NmtBomExtract1.Main"
    Dim macroErrCode, macroOk
    macroErrCode = CLng(0)
    Err.Clear
    macroOk = swApp.RunMacro2(swpPath, "NmtBomExtract1", "Main", CLng(0), macroErrCode)
    If Err.Number <> 0 Then
        LogMsg "RunMacro2 COM error: " & CStr(Err.Number) & " - " & Err.Description
        Err.Clear
    Else
        LogMsg "RunMacro2 returned: ok=" & CStr(macroOk) & " errCode=" & CStr(macroErrCode)
    End If

    ' If RunMacro2 failed or produced no file, try RunMacro (3-arg, no ByRef — avoids type mismatch)
    If Not gFSO.FileExists(macroOutPath) Then
        Dim rmCandidates, rmi2, rmMod2, rmOk2
        rmCandidates = Array("NmtBomExtract1", "NmtBomExtract", "Module1", "")
        For rmi2 = 0 To UBound(rmCandidates)
            rmMod2 = CStr(rmCandidates(rmi2))
            Err.Clear
            rmOk2 = swApp.RunMacro(swpPath, rmMod2, "Main")
            If Err.Number <> 0 Then
                LogMsg "RunMacro(" & rmMod2 & ") error: " & CStr(Err.Number) & " - " & Err.Description
                Err.Clear
            Else
                LogMsg "RunMacro(" & rmMod2 & ") returned: " & CStr(rmOk2)
                If CStr(rmOk2) = "True" Or CStr(rmOk2) = "-1" Then Exit For
            End If
        Next
    End If

    ' Wait up to 30s for the macro to write its output file
    Dim macroDeadline
    macroDeadline = DateAdd("s", 30, Now)
    Do While Not gFSO.FileExists(macroOutPath) And Now < macroDeadline
        WScript.Sleep 200
    Loop

    LogMsg "Macro wait done. Output file exists: " & CStr(gFSO.FileExists(macroOutPath))
    If gFSO.FileExists(macroOutPath) Then
        Dim mofTs, mLine, mpn
        Set mofTs = gFSO.OpenTextFile(macroOutPath, 1, False, 0)
        Do Until mofTs.AtEndOfStream
            mLine = Trim(mofTs.ReadLine)
            If UCase(Left(mLine, 5)) = "PART|" Then
                mpn = Trim(Mid(mLine, 6))
                If mpn <> "" Then
                    If Not gParts.Exists(mpn) Then gParts.Add mpn, True
                    macroPartCount = macroPartCount + 1
                End If
            End If
        Loop
        mofTs.Close
        LogMsg "RunMacro2 produced " & CStr(macroPartCount) & " PART lines."
    Else
        LogMsg "RunMacro2 output file not found; falling back to VBScript traversal."
    End If
    On Error GoTo 0
Else
    LogMsg "NmtBomExtract.swp not found at: " & swpPath & " — falling back to VBScript traversal."
End If

If macroPartCount = 0 Then
    CollectPartsRecursive rootComp
End If

asmBase = FileBaseNameUpper(gAssemblyPath)
If asmBase <> "" Then
    If Not gParts.Exists(asmBase) Then gParts.Add asmBase, True
End If

For Each k In gParts.Keys
    EmitPart CStr(k)
Next

LogMsg "VBS helper finished. Parts: " & CStr(gParts.Count)
Cleanup swApp, createdSw, openTitle
WScript.Quit 0

Sub ParseArgs()
    Dim i, arg, argsFile
    argsFile = ""

    i = 0
    Do While i < WScript.Arguments.Count
        arg = CStr(WScript.Arguments(i))
        If LCase(arg) = "/argsfile" Then
            If i + 1 < WScript.Arguments.Count Then
                argsFile = CStr(WScript.Arguments(i + 1))
                i = i + 2
            Else
                i = i + 1
            End If
        Else
            ParseKeyValue arg
            i = i + 1
        End If
    Loop

    If argsFile <> "" Then
        ReadArgsFile argsFile
    End If
End Sub

Sub ParseKeyValue(ByVal line)
    Dim eqPos, key, valueText
    line = Trim(line)
    If line = "" Then Exit Sub

    eqPos = InStr(1, line, "=", vbTextCompare)
    If eqPos > 1 Then
        key = UCase(Trim(Left(line, eqPos - 1)))
        valueText = Mid(line, eqPos + 1)
        If key = "ASSEMBLY" Then
            gAssemblyPath = valueText
            Exit Sub
        End If
        If key = "OUT" Then
            gOutPath = valueText
            Exit Sub
        End If
    End If

    If gAssemblyPath = "" Then
        gAssemblyPath = line
    ElseIf gOutPath = "" Then
        gOutPath = line
    End If
End Sub

Sub ReadArgsFile(ByVal path)
    On Error Resume Next
    Dim ts, line
    Set ts = gFSO.OpenTextFile(path, 1, False)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        ParseKeyValue line
    Loop
    ts.Close
End Sub

Sub InitOutput()
    If Trim(gOutPath) = "" Then Exit Sub
    On Error Resume Next
    Dim parentDir
    parentDir = ""
    parentDir = gFSO.GetParentFolderName(gOutPath)
    If parentDir <> "" And Not gFSO.FolderExists(parentDir) Then
        gFSO.CreateFolder parentDir
    End If
    Dim ts
    Set ts = gFSO.CreateTextFile(gOutPath, True)
    ts.Close
    On Error GoTo 0
End Sub

Sub WriteLine(ByVal textLine)
    If Trim(gOutPath) = "" Then
        WScript.Echo textLine
        Exit Sub
    End If

    On Error Resume Next
    Dim ts
    Set ts = gFSO.OpenTextFile(gOutPath, 8, True)
    ts.WriteLine textLine
    ts.Close
    On Error GoTo 0
End Sub

Sub LogMsg(ByVal msg)
    WriteLine "LOG|" & msg
End Sub

Sub EmitPart(ByVal pn)
    pn = UCase(Trim(pn))
    If pn = "" Then Exit Sub
    WriteLine "PART|" & pn
End Sub

Function FileBaseNameUpper(ByVal pathOrName)
    On Error Resume Next
    Dim nm, dotPos
    nm = gFSO.GetFileName(pathOrName)
    dotPos = InStrRev(nm, ".")
    If dotPos > 1 Then nm = Left(nm, dotPos - 1)
    FileBaseNameUpper = UCase(Trim(nm))
    On Error GoTo 0
End Function

Function GetCompPath(ByVal comp)
    On Error Resume Next
    Dim p, m
    p = ""
    p = CStr(comp.GetPathName)
    If Trim(p) = "" Then
        Set m = Nothing
        Set m = comp.GetModelDoc2
        If Not m Is Nothing Then p = CStr(m.GetPathName)
    End If
    GetCompPath = p
    On Error GoTo 0
End Function

Function GetModelProp(ByVal modelObj, ByVal propName)
    On Error Resume Next
    Dim ext, cpm, v, rv, ok
    GetModelProp = ""
    Set ext = Nothing
    Set cpm = Nothing
    Set ext = modelObj.Extension
    If ext Is Nothing Then Exit Function
    Set cpm = ext.CustomPropertyManager("")
    If cpm Is Nothing Then Exit Function

    v = ""
    rv = ""
    ok = cpm.Get4(propName, False, v, rv)
    If Trim(CStr(rv)) <> "" Then
        GetModelProp = CStr(rv)
    ElseIf Trim(CStr(v)) <> "" Then
        GetModelProp = CStr(v)
    End If
    On Error GoTo 0
End Function

Function GetCompPartNumber(ByVal comp)
    On Error Resume Next
    Dim modelObj, pn, nm, path
    pn = ""

    ' Optimization: If the filename itself looks like a standard part number (e.g. 12345-A01),
    ' use it directly and skip the expensive COM property probe.
    path = GetCompPath(comp)
    nm = FileBaseNameUpper(path)
    If nm <> "" Then
        ' Check for standard part number pattern: 4-6 digits, dash, letter, 1-8 chars
        Dim regex, matches
        Set regex = New RegExp
        regex.Pattern = "^\d{4,6}-[A-Z][A-Z0-9]{1,8}(?:-[LR])?$"
        regex.IgnoreCase = True
        If regex.Test(nm) Then
            GetCompPartNumber = UCase(Trim(nm))
            Exit Function
        End If
    End If

    Set modelObj = Nothing
    Set modelObj = comp.GetModelDoc2
    If Not modelObj Is Nothing Then
        pn = GetModelProp(modelObj, "Part Number")
        If Trim(pn) = "" Then pn = GetModelProp(modelObj, "PartNumber")
        If Trim(pn) = "" Then pn = GetModelProp(modelObj, "PartNo")
        If Trim(pn) = "" Then pn = GetModelProp(modelObj, "Part No")
        If Trim(pn) = "" Then pn = GetModelProp(modelObj, "DrawingNo")
        If Trim(pn) = "" Then pn = GetModelProp(modelObj, "Drawing No")
        If Trim(pn) = "" Then pn = GetModelProp(modelObj, "Number")
    End If

    If Trim(pn) = "" Then
        If nm <> "" Then pn = nm
    End If

    GetCompPartNumber = UCase(Trim(pn))
    On Error GoTo 0
End Function

Function IsIgnoredPart(ByVal pn)
    Dim u
    u = UCase(Trim(pn))
    If u = "" Then
        IsIgnoredPart = True
        Exit Function
    End If
    If InStr(u, "LOAD CERT") > 0 Then
        IsIgnoredPart = True
        Exit Function
    End If
    If InStr(u, "LOAD_CERT") > 0 Then
        IsIgnoredPart = True
        Exit Function
    End If
    If InStr(u, "SCOPE") > 0 Then
        IsIgnoredPart = True
        Exit Function
    End If
    If InStr(u, "MANUAL") > 0 Then
        IsIgnoredPart = True
        Exit Function
    End If
    IsIgnoredPart = False
End Function

Sub CollectPartsRecursive(ByVal comp)
    If comp Is Nothing Then Exit Sub

    On Error Resume Next
    Dim suppressed
    suppressed = False
    suppressed = CBool(comp.IsSuppressed)
    If Err.Number <> 0 Then
        Err.Clear
        suppressed = False
    End If
    On Error GoTo 0
    If suppressed Then Exit Sub

    Dim modelPath, cfg, seenKey
    modelPath = GetCompPath(comp)
    cfg = ""
    On Error Resume Next
    cfg = CStr(comp.ReferencedConfiguration)
    On Error GoTo 0

    If Trim(modelPath) <> "" Then
        seenKey = LCase(modelPath & "|" & cfg)
        If gSeen.Exists(seenKey) Then Exit Sub
        gSeen.Add seenKey, True
    End If

    Dim pn
    pn = GetCompPartNumber(comp)
    If Not IsIgnoredPart(pn) Then
        If Not gParts.Exists(pn) Then gParts.Add pn, True
    End If

    ' Resolve this component if it's a sub-assembly so GetChildren works at depth
    On Error Resume Next
    Dim compResolveDoc
    Set compResolveDoc = Nothing
    Set compResolveDoc = comp.GetModelDoc2
    If Not compResolveDoc Is Nothing Then
        If compResolveDoc.GetType() = 2 Then compResolveDoc.ResolveAllLightWeightComponents True
    End If
    On Error GoTo 0

    On Error Resume Next
    Dim children, child
    children = comp.GetChildren
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    If IsArray(children) Then
        For Each child In children
            If Not child Is Nothing Then CollectPartsRecursive child
        Next
    End If
End Sub

Function OpenAssembly(ByVal swAppObj, ByVal path)
    On Error Resume Next
    Dim m, errs, warns, opt
    Set m = Nothing

    Set m = swAppObj.GetOpenDocumentByName(path)
    If Not m Is Nothing Then
        Set OpenAssembly = m
        Exit Function
    End If

    swAppObj.SetCurrentWorkingDirectory gFSO.GetParentFolderName(path)
    Err.Clear

    For Each opt In Array(3, 1, 0, 2)
        errs = 0
        warns = 0
        Set m = swAppObj.OpenDoc6(path, 2, CLng(opt), "", errs, warns)
        If Not m Is Nothing Then
            Set OpenAssembly = m
            Exit Function
        End If
    Next

    Set m = swAppObj.OpenDoc(path, 2)
    If Not m Is Nothing Then
        Set OpenAssembly = m
        Exit Function
    End If

    Set OpenAssembly = Nothing
End Function

Sub Cleanup(ByVal swAppObj, ByVal createdObj, ByVal docTitle)
    On Error Resume Next
    If Not swAppObj Is Nothing Then
        If createdObj Then
            swAppObj.ExitApp
        ElseIf Trim(docTitle) <> "" Then
            swAppObj.CloseDoc docTitle
        End If
    End If
    On Error GoTo 0
End Sub
