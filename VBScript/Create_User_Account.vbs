'<<<< Prompt for username >>>>
username = inputbox("Please Enter a username:")
password = inputbox("Please Enter a password:")
if username = "" or password = "" then wscript.quit

'<<<< RootDSE is used to obtain the default naming context (saves hard-coding the domain) >>>>
set objRoot = getobject("LDAP://RootDSE")

'**** Bind to the default users container ****
set objContainer = getobject("LDAP://cn=users," & objRoot.get("defaultnamingcontext"))

'<<<< Create the user object >>>>
set objUser = objContainer.Create("user","cn=" & username)

'<<<< The sAMAccountName is the username the user will use to logon >>>>
objUser.sAMAccountName = username

'<<<< Save the changes >>>>
objUser.Setinfo

'<<<< Set a password >>>>
objUser.setpassword password

'<<<< Enable the account >>>>
objUser.AccountDisabled = False

'<<<< Save the changes >>>>
objUser.Setinfo