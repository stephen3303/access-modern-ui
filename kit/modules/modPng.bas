Option Compare Database
Option Explicit

' Access UI kit - rounded pill images. Access controls only draw
' square corners, so pills and badges sit on a bound/unbound image control
' showing an anti-aliased capsule PNG drawn here. Images are written once per
' color and size to %LOCALAPPDATA%\<APP_NAME>\pills and reused.

Private Const SCALE_UP As Long = 2          ' draw at 2x for crisp edges when zoomed
Private m_crc(255) As Long
Private m_crcReady As Boolean
Private m_cache As Collection
Private m_dir As String

' Path of a capsule image of the given Access color, sized for a control of
' wPx x hPx screen pixels. Null color (no matching condition) -> no image.
Public Function PillPng(ByVal clr As Variant, ByVal wPx As Long, ByVal hPx As Long) As Variant
    Dim key As String, path As String
    PillPng = Null
    If IsNull(clr) Or IsEmpty(clr) Then Exit Function
    If clr < 0 Or wPx < 2 Or hPx < 2 Then Exit Function
    key = "p_" & Right$("00000" & Hex$(clr), 6) & "_" & wPx & "x" & hPx
    If m_cache Is Nothing Then Set m_cache = New Collection
    On Error Resume Next
    path = m_cache(key)
    On Error GoTo Fail
    If Len(path) = 0 Then
        path = PillDir() & "\" & key & ".png"
        If Len(Dir$(path)) = 0 Then WriteCapsule path, CLng(clr), wPx * SCALE_UP, hPx * SCALE_UP
        m_cache.Add path, key
    End If
    PillPng = path
    Exit Function
Fail:
    PillPng = Null
End Function

' %LOCALAPPDATA%\<APP_NAME>\pills (TEMP if LOCALAPPDATA is missing), created on first use.
Private Function PillDir() As String
    Dim base As String
    If Len(m_dir) = 0 Then
        base = Environ$("LOCALAPPDATA")
        If Len(base) = 0 Then base = Environ$("TEMP")
        m_dir = base & "\" & APP_NAME
        If Len(Dir$(m_dir, vbDirectory)) = 0 Then MkDir m_dir
        m_dir = m_dir & "\pills"
        If Len(Dir$(m_dir, vbDirectory)) = 0 Then MkDir m_dir
    End If
    PillDir = m_dir
End Function

' Capsule (fully rounded ends) in color clr on a transparent background.
Private Sub WriteCapsule(ByVal path As String, ByVal clr As Long, ByVal w As Long, ByVal h As Long)
    Dim raw() As Byte, x As Long, y As Long, i As Long
    Dim r As Double, cx As Double, cy As Double, px As Double, py As Double, d As Double, a As Double
    Dim cr As Byte, cg As Byte, cb As Byte
    cr = clr And &HFF: cg = (clr \ &H100) And &HFF: cb = (clr \ &H10000) And &HFF
    r = h / 2
    If r > w / 2 Then r = w / 2
    cy = h / 2
    ReDim raw(0 To h * (w * 4 + 1) - 1)
    i = 0
    For y = 0 To h - 1
        raw(i) = 0: i = i + 1                         ' filter: none
        py = y + 0.5
        For x = 0 To w - 1
            px = x + 0.5
            ' distance from the pixel center to the capsule's center segment
            If px < r Then
                cx = r
            ElseIf px > w - r Then
                cx = w - r
            Else
                cx = px
            End If
            d = Sqr((px - cx) ^ 2 + (py - cy) ^ 2)
            a = r - d + 0.5
            If a < 0 Then a = 0
            If a > 1 Then a = 1
            raw(i) = cr: raw(i + 1) = cg: raw(i + 2) = cb: raw(i + 3) = CByte(a * 255)
            i = i + 4
        Next
    Next
    WritePng path, w, h, raw
End Sub

' ---------------------------------------------------------------------------
' Minimal PNG writer: RGBA8, one IDAT holding a zlib stream of stored blocks.
' ---------------------------------------------------------------------------

' Write raw (filtered RGBA scanlines) as a PNG file.
Private Sub WritePng(ByVal path As String, ByVal w As Long, ByVal h As Long, raw() As Byte)
    Dim out() As Byte, n As Long, hdr(12) As Byte, z() As Byte, fh As Integer
    ReDim out(0 To 7)
    out(0) = 137: out(1) = 80: out(2) = 78: out(3) = 71: out(4) = 13: out(5) = 10: out(6) = 26: out(7) = 10
    n = 8
    PutBE hdr, 0, w: PutBE hdr, 4, h
    hdr(8) = 8: hdr(9) = 6: hdr(10) = 0: hdr(11) = 0: hdr(12) = 0
    AddChunk out, n, "IHDR", hdr, 13
    z = Zlib(raw)
    AddChunk out, n, "IDAT", z, UBound(z) + 1
    AddChunk out, n, "IEND", hdr, 0
    ReDim Preserve out(0 To n - 1)
    If Len(Dir$(path)) > 0 Then Kill path
    fh = FreeFile
    Open path For Binary Access Write As #fh
    Put #fh, , out
    Close #fh
End Sub

' zlib stream of uncompressed (stored) blocks with an Adler-32 checksum. The
' images are tiny, so compression isn't worth the code.
Private Function Zlib(raw() As Byte) As Byte()
    Dim total As Long, blocks As Long, z() As Byte, i As Long, pos As Long, ln As Long, b As Long
    Dim s1 As Long, s2 As Long
    total = UBound(raw) + 1
    blocks = (total + 65534) \ 65535
    If blocks = 0 Then blocks = 1
    ReDim z(0 To 2 + blocks * 5 + total + 4 - 1)
    z(0) = &H78: z(1) = &H1
    pos = 2
    For b = 0 To blocks - 1
        ln = total - b * 65535
        If ln > 65535 Then ln = 65535
        z(pos) = IIf(b = blocks - 1, 1, 0)
        z(pos + 1) = ln And &HFF: z(pos + 2) = (ln \ &H100) And &HFF
        z(pos + 3) = (65535 - ln) And &HFF: z(pos + 4) = ((65535 - ln) \ &H100) And &HFF
        pos = pos + 5
        For i = 0 To ln - 1
            z(pos + i) = raw(b * 65535& + i)
        Next
        pos = pos + ln
    Next
    s1 = 1: s2 = 0
    For i = 0 To total - 1
        s1 = (s1 + raw(i)) Mod 65521
        s2 = (s2 + s1) Mod 65521
    Next
    z(pos) = (s2 \ &H100) And &HFF: z(pos + 1) = s2 And &HFF
    z(pos + 2) = (s1 \ &H100) And &HFF: z(pos + 3) = s1 And &HFF
    Zlib = z
End Function

' Append a PNG chunk (length, type, data, CRC) to out at position n.
Private Sub AddChunk(out() As Byte, n As Long, ByVal typ As String, data() As Byte, ByVal ln As Long)
    Dim i As Long, crc As Long, start As Long
    ReDim Preserve out(0 To n + 12 + ln - 1)
    PutBE out, n, ln
    start = n + 4
    For i = 1 To 4
        out(start + i - 1) = Asc(Mid$(typ, i, 1))
    Next
    For i = 0 To ln - 1
        out(start + 4 + i) = data(i)
    Next
    crc = Crc32(out, start, 4 + ln)
    PutBE out, start + 4 + ln, crc
    n = start + 4 + ln + 4
End Sub

' Write a 32-bit big-endian integer.
Private Sub PutBE(b() As Byte, ByVal at As Long, ByVal v As Long)
    b(at) = ((v And &HFF000000) \ &H1000000) And &HFF
    b(at + 1) = (v And &HFF0000) \ &H10000
    b(at + 2) = (v And &HFF00&) \ &H100
    b(at + 3) = v And &HFF
End Sub

' CRC-32 as PNG requires (table built on first use).
Private Function Crc32(b() As Byte, ByVal start As Long, ByVal ln As Long) As Long
    Dim c As Long, i As Long, k As Long
    If Not m_crcReady Then
        For i = 0 To 255
            c = i
            For k = 1 To 8
                If c And 1 Then
                    c = (((c And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor &HEDB88320
                Else
                    c = ((c And &HFFFFFFFE) \ 2) And &H7FFFFFFF
                End If
            Next
            m_crc(i) = c
        Next
        m_crcReady = True
    End If
    c = &HFFFFFFFF
    For i = start To start + ln - 1
        c = (((c And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor m_crc((c Xor b(i)) And &HFF)
    Next
    Crc32 = c Xor &HFFFFFFFF
End Function
