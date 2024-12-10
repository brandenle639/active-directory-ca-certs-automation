set wso = CreateObject("Wscript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

'####Variables####
'Enter the site uri
sitename = InputBox("Enter Site Uri")
folderloca = "<Folder URI>" & sitename & "\"
fleloc = folderloca & sitename
EASYRSA_REQ_COUNTRY = "<Country Code: 2 char>"
EASYRSA_REQ_PROVINCE = "<STATE>"
EASYRSA_REQ_CITY = "<City>"
EASYRSA_REQ_OU = "<COMPANY>"
zpfolderloca = "<Folder URI>"
zp = zpfolderloca & sitename & ".zip"
outFile = fleloc & "-command.log"
'#################

'###Script Quit on Nulls or Empties###
'Quits if input is empty
If IsEmpty(sitename)  Or sitename = "" Then
    WScript.Quit
End If
'Quits if there is variable issue
If InSTR(EASYRSA_REQ_COUNTRY,"<") Or InSTR(EASYRSA_REQ_COUNTRY,">") Or InSTR(EASYRSA_REQ_PROVINCE,"<") Or InSTR(EASYRSA_REQ_PROVINCE,">") Or InSTR(EASYRSA_REQ_CITY,"<") Or InSTR(EASYRSA_REQ_CITY,">") Or InSTR(EASYRSA_REQ_OU,"<") Or InSTR(EASYRSA_REQ_OU,">") Or InSTR(zpfolderloca,"<") Or InSTR(zpfolderloca,">") Then
    erro=MsgBox("Variable Maybe Set Wrong",16,"Cert Output Error")
    WScript.Quit
End If
'Quits if import contains http or https
If InSTR(sitename,"http://") Or InSTR(sitename,"https://") Then
    erro=MsgBox("Site Name Should Not Contain http:// or https://",16,"Cert Output Error")
    WScript.Quit
End If
'#####################################

'####Create Certs####
'Makes the output directory
set mkdr = wso.Exec("cmd /c mkdir " & folderloca)
WScript.Sleep(5000)
'Create Output file to write to
Set objFile = test.CreateTextFile(outFile,True)
'Makes the requested key
set keymkr = wso.Exec("cmd /c openssl genrsa -out " & fleloc & ".key")
'Writes command to command.log
objFile.Writeline "openssl genrsa -out " & fleloc & ".key"
WScript.Sleep(5000)
'Makes the requested request file
set nwreq = wso.Exec("cmd /c openssl req -new -key " & fleloc & ".key -subj ""/CN=" & sitename & "/C=" & EASYRSA_REQ_COUNTRY & "/ST=" & EASYRSA_REQ_PROVINCE & "/L=" & EASYRSA_REQ_CITY & "/O=" & EASYRSA_REQ_OU & """ -out " & fleloc & ".req")
'Writes command to command.log
objFile.Writeline "openssl req -new -key " & fleloc & ".key -subj ""/CN=" & sitename & "/C=" & EASYRSA_REQ_COUNTRY & "/ST=" & EASYRSA_REQ_PROVINCE & "/L=" & EASYRSA_REQ_CITY & "/O=" & EASYRSA_REQ_OU & """ -out " & fleloc & ".req"
WScript.Sleep(5000)
'Imports the request file
set certsign = wso.Exec("certreq -q -f -Kerberos -attrib ""CertificateTemplate:WebServer"" -submit " & fleloc & ".req")
'Writes command to command.log
objFile.Writeline "certreq -q -f -Kerberos -attrib ""CertificateTemplate:WebServer"" -submit " & fleloc & ".req"
WScript.Sleep(5000)
'Gets the requested id
sout = certsign.StdOut.ReadAll
lns = split(sout,vbCrLf)
nid = replace(lns(3),"RequestId: ", "")
'Resubmits the cert based on id
set certresub = wso.Exec("certutil -Resubmit " & nid)
'Writes command to command.log
objFile.Writeline "certutil -Resubmit " & nid
WScript.Sleep(5000)
'Retrieves the cert based on the id and exports to the crt file
set getcert = wso.Exec("certreq -q -f -Kerberos -Retrieve " & nid & " " & fleloc & ".crt")
'Writes command to command.log
objFile.Writeline "certreq -q -f -Kerberos -Retrieve " & nid & " " & fleloc & ".crt"
WScript.Sleep(5000)
'Creates the pfx file
set certsign = wso.Exec("openssl.exe pkcs12 -export -out " & fleloc & ".pfx -inkey " & fleloc & ".key -in " & fleloc & ".crt -in " & fleloc & ".crt -passout pass:")
'Writes command to command.log
objFile.Writeline "openssl.exe pkcs12 -export -out " & fleloc & ".pfx -inkey " & fleloc & ".key -in " & fleloc & ".crt -in " & fleloc & ".crt -passout pass:"
WScript.Sleep(5000)
objFile.Close
'####################

'####Zip Creation###
'Creates a text file marked as a zip'Imports the request file
Set ts = fso.CreateTextFile(zp, True)
ts.Write "PK" & Chr(5) & Chr(6) & String( 18, Chr(0) )
ts.Close
Set ts = Nothing
Set fso = Nothing
WScript.Sleep(5000)
'###################

'####Zip Related Files####
'Compresses the crt, key, pfx to the output zip
Set zip = CreateObject("Shell.Application").NameSpace(zp)
' Add all files/directories to the .zip file
zip.CopyHere(fleloc & ".crt")
WScript.Sleep(5000)
zip.CopyHere(fleloc & ".key")
WScript.Sleep(5000)
zip.CopyHere(fleloc & ".pfx")
WScript.Sleep(5000)
'#########################

msgbox("Done")
