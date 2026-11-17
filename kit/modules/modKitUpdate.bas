Option Compare Database
Option Explicit

' Access UI kit - "a newer version is available" notice.
'
' A release folder holds the app's current copy and a one-line text file,
' version.txt, naming that copy's version ("1.2.0"). At startup the app calls
' CheckForUpdate with its own version and the folder. When the file names a
' newer version, the user gets one KitMsgBox saying which version and where to
' get it, once per session (reopening the main window doesn't repeat it). A
' blank folder, a missing folder or file, or a file that can't be read never
' stops the app opening: the check is skipped, and a failed read is logged
' (LogError) without a message. An unreachable network folder can delay
' opening until Windows gives up on it (about 20 seconds).
'
' A compiled front end (.accde) opens only on the bitness of the Office that
' compiled it, so a release folder holds one per bitness: <app>_FE_x64.accde
' and <app>_FE_x86.accde. Pass the front end's base name (frontEnd, e.g.
' "SABER_FE") and the notice names the file for this Office
' (ReleasedFrontEnd): the per-bitness copy, or <app>_FE.accde from a release
' made before per-bitness builds. When the folder has only the other
' bitness's copy, the notice says so instead of sending a 32-bit PC to the
' 64-bit file.
'
' The app provides APP_NAME (named in the notice) and LogError (modErrors).

Public Const KIT_VERSION_FILE As String = "version.txt"

Private m_checked As Boolean      ' CheckForUpdate already ran this session

' Show the notice when the release folder's version.txt names a version newer
' than current. True when the notice was shown. Only the first call in a
' session checks; later calls return False. advice replaces the default next
' step ("Close <app> and copy the new version from: <folder>"), for an app
' where copying the new file over the old one would lose data. frontEnd is the
' front end's name without extension ("SABER_FE"): given, the default next
' step names the file for this Office's bitness (FrontEndAdvice).
Public Function CheckForUpdate(ByVal current As String, ByVal folder As String, _
        Optional ByVal advice As String = "", Optional ByVal frontEnd As String = "") As Boolean
    Dim latest As String
    On Error GoTo Fail
    If m_checked Then Exit Function
    m_checked = True
    latest = ReleasedVersion(folder)
    If Len(latest) = 0 Then Exit Function
    If CompareVersions(latest, current) <= 0 Then Exit Function
    If Len(advice) = 0 Then advice = FrontEndAdvice(folder, frontEnd)
    KitMsgBox "Version " & latest & " of " & APP_NAME & " is available. You are using version " & current & "." & _
        vbCrLf & vbCrLf & advice, vbInformation, "New version available"
    CheckForUpdate = True
    Exit Function
Fail:
    LogError "modKitUpdate.CheckForUpdate", "folder " & folder & ", this version " & current
End Function

' The next step the notice gives: copy this Office's front end from folder.
' Without frontEnd, or when the folder has no front end the kit recognises,
' "Close <app> and copy the new version from: <folder>". When the folder has
' only the other bitness's copy, tells the user to ask their administrator
' for one built for their Office.
Public Function FrontEndAdvice(ByVal folder As String, Optional ByVal frontEnd As String = "", _
        Optional ByVal bitness As String = "") As String
    Dim fileName As String
    If Len(bitness) = 0 Then bitness = KitOfficeBitness()
    If Len(frontEnd) > 0 Then
        fileName = ReleasedFrontEnd(folder, frontEnd, bitness)
        If Len(fileName) > 0 Then
            FrontEndAdvice = "Close " & APP_NAME & " and copy " & fileName & " from:" & vbCrLf & FolderPath(folder)
            Exit Function
        End If
        If FileExists(FolderPath(folder) & "\" & FrontEndFileName(frontEnd, OtherBitness(bitness))) Then
            FrontEndAdvice = "This release has no copy for " & BitnessLabel(bitness) & " Office yet. Ask your " & _
                "administrator for " & FrontEndFileName(frontEnd, bitness) & " (it goes in:" & vbCrLf & _
                FolderPath(folder) & ")."
            Exit Function
        End If
    End If
    FrontEndAdvice = "Close " & APP_NAME & " and copy the new version from:" & vbCrLf & FolderPath(folder)
End Function

' The name of the front end in folder that this Office can open: frontEnd &
' "_" & bitness & ".accde" (bitness defaults to KitOfficeBitness), or, when the
' folder has neither bitness's copy, frontEnd & ".accde" from a release made
' before per-bitness builds. "" when the folder is blank or holds neither, or
' holds only the other bitness's copy (an older single .accde beside it is
' stale and not offered).
Public Function ReleasedFrontEnd(ByVal folder As String, ByVal frontEnd As String, _
        Optional ByVal bitness As String = "") As String
    Dim base As String
    If Len(Trim$(folder)) = 0 Or Len(frontEnd) = 0 Then Exit Function
    If Len(bitness) = 0 Then bitness = KitOfficeBitness()
    base = FolderPath(folder) & "\"
    If FileExists(base & FrontEndFileName(frontEnd, bitness)) Then
        ReleasedFrontEnd = FrontEndFileName(frontEnd, bitness)
    ElseIf FileExists(base & FrontEndFileName(frontEnd, OtherBitness(bitness))) Then
        Exit Function
    ElseIf FileExists(base & frontEnd & ".accde") Then
        ReleasedFrontEnd = frontEnd & ".accde"
    End If
End Function

' The per-bitness release file name: "SABER_FE", "x64" -> "SABER_FE_x64.accde".
Public Function FrontEndFileName(ByVal frontEnd As String, Optional ByVal bitness As String = "") As String
    If Len(bitness) = 0 Then bitness = KitOfficeBitness()
    FrontEndFileName = frontEnd & "_" & bitness & ".accde"
End Function

Private Function OtherBitness(ByVal bitness As String) As String
    OtherBitness = IIf(bitness = "x64", "x86", "x64")
End Function

' "64-bit" or "32-bit", as Office's About page words it.
Private Function BitnessLabel(ByVal bitness As String) As String
    BitnessLabel = IIf(bitness = "x64", "64-bit", "32-bit")
End Function

' The version named by folder\version.txt (its first non-blank line, trimmed),
' or "" when the folder is blank or the file doesn't exist. Errors reading the
' file go to the caller.
Public Function ReleasedVersion(ByVal folder As String) As String
    Dim path As String, line As Variant
    If Len(Trim$(folder)) = 0 Then Exit Function
    path = FolderPath(folder) & "\" & KIT_VERSION_FILE
    If Not FileExists(path) Then Exit Function
    For Each line In Split(Replace(ReadTextFile(path), vbCr, ""), vbLf)
        If Len(Trim$(line)) > 0 Then
            ReleasedVersion = Trim$(line)
            Exit Function
        End If
    Next
End Function

' Compare dotted versions part by part as numbers: -1 when a is older than b,
' 0 when they match, 1 when a is newer. A leading "v" is ignored, missing parts
' count as 0 ("1.2" = "1.2.0"), and each part counts only its leading digits
' ("3-beta" = 3; "x" = 0).
Public Function CompareVersions(ByVal a As String, ByVal b As String) As Integer
    Dim pa() As String, pb() As String, i As Long, n As Long, va As Double, vb As Double
    pa = Split(StripV(a), ".")
    pb = Split(StripV(b), ".")
    n = UBound(pa)
    If UBound(pb) > n Then n = UBound(pb)
    For i = 0 To n
        va = VersionPart(pa, i)
        vb = VersionPart(pb, i)
        If va < vb Then CompareVersions = -1: Exit Function
        If va > vb Then CompareVersions = 1: Exit Function
    Next
End Function

' A version without spaces or its leading "v".
Private Function StripV(ByVal s As String) As String
    s = Trim$(s)
    If LCase$(Left$(s, 1)) = "v" Then s = Mid$(s, 2)
    StripV = s
End Function

' Part i of a split version as a number: its leading digits, or 0 when the
' part is missing or doesn't start with a digit.
Private Function VersionPart(parts() As String, ByVal i As Long) As Double
    Dim s As String, k As Long
    If i > UBound(parts) Then Exit Function
    s = parts(i)
    For k = 1 To Len(s)
        If Not (Mid$(s, k, 1) Like "#") Then Exit For
    Next
    s = Left$(s, k - 1)
    If Len(s) > 0 Then VersionPart = Val(s)
End Function

' A folder path without spaces or a trailing backslash.
Private Function FolderPath(ByVal folder As String) As String
    folder = Trim$(folder)
    Do While Right$(folder, 1) = "\" And Len(folder) > 1
        folder = Left$(folder, Len(folder) - 1)
    Loop
    FolderPath = folder
End Function
