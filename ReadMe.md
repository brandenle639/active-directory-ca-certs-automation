This script must be run on the Windows Active Directory CA Authority Server

Requirements

    1. Make sure the openssl bin folder is set in the PATH variable

    2. The Windows Active Directory CA Authority is installed and configured


To Configure Script

    1. Edit the following replacing the <> with the related information in the script file:

        folderloca = "<Folder URI>" & sitename & "\"
        EASYRSA_REQ_COUNTRY = "<Country Code: 2 char>"
        EASYRSA_REQ_PROVINCE = "<STATE>"
        EASYRSA_REQ_CITY = "<City>"
        EASYRSA_REQ_OU = "<COMPANY>"
        zpfolderloca = "<Folder URI>"

To Run Script

    Warning: May needed to be run as Administrator

    1. Run the Vbscript

    2. Enter your site uri

    3. Click Ok

    4. Let the process run

If Possible Errors:

    Open the related command.log and run each or specific line(s) in command prompt

        May have to run commands as Administrator
