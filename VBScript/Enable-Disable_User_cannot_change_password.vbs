'<<<< Force Variable decleration >>>>
Option Explicit

Const CHANGE_PASSWORD_GUID = "{AB721A53-1E2F-11D0-9819-00AA0040529B}"
Const ADS_RIGHT_DS_CONTROL_ACCESS = &H100
Const ADS_ACETYPE_ACCESS_ALLOWED = &H0
Const ADS_ACETYPE_ACCESS_DENIED = &H1
Const ADS_ACETYPE_ACCESS_ALLOWED_OBJECT = &H5
Const ADS_ACETYPE_ACCESS_DENIED_OBJECT = &H6
Const ADS_ACEFLAG_INHERITED_ACE = &H10
Const ADS_ACEFLAG_OBJECT_TYPE_PRESENT = &H1


Dim objUser
'<<<< Bind to the user object using the distinguished name >>>>
set objUser = getobject("LDAP://cn=test.3,cn=users,dc=wisesoft,dc=co,dc=uk")

'<<<< Enable User Cannot change password >>>>
setUserCannotChangePassword objuser, True

'<<<< Disable User Cannot change password >>>>
setUserCannotChangePassword objuser, False


'<<<< Function takes a user object and sets the user cannot change password option >>>>
function setUserCannotChangePassword(byval objUser, Value)
	
	Dim objACESelf, objACEEveryone, objSecDescriptor, objDACL
	Dim strDN, objACE, blnSelf, blnEveryone, blnModified

	' Bind to the user security objects.
	Set objSecDescriptor = objUser.Get("ntSecurityDescriptor")
	Set objDACL = objSecDescriptor.discretionaryAcl

	' Search for ACE's for Change Password and modify.
	blnSelf = False
	blnEveryone = False
	blnModified = False
	For Each objACE In objDACL
  		If UCase(objACE.objectType) = UCase(CHANGE_PASSWORD_GUID) Then
    			If UCase(objACE.Trustee) = "NT AUTHORITY\SELF" Then
				If Value then
      					If objACE.AceType = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT Then
        					objACE.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT
        					blnModified = True
      					End If
				else
					If objACE.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT Then
        					objACE.AceType = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
        					blnModified = True
      					End If
				end if
      				blnSelf = True
    			End If
    			If UCase(objACE.Trustee) = "EVERYONE" Then
				If Value then
      					If objACE.AceType = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT Then
        					objACE.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT
        					blnModified = True
      					End If
				else
					If objACE.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT Then
        					objACE.AceType = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
        					blnModified = True
      					End If
				end if
      				blnEveryone = True
    			End If
  		End If
	Next

	' If ACE's found and modified, save changes and exit.
	If (blnSelf = True) And (blnEveryone = True) Then
  		If blnModified Then
    			objSecDescriptor.discretionaryACL = Reorder(objDACL)
    			objUser.Put "ntSecurityDescriptor", objSecDescriptor
    			objUser.SetInfo
  		End If
	else
		' If ACE's not found, add to DACL.
		If blnSelf = False Then
			' Create the ACE for Self.
  			Set objACESelf = CreateObject("AccessControlEntry")
  			objACESelf.Trustee = "NT AUTHORITY\SELF"
  			objACESelf.AceFlags = 0
			if Value then
  				objACESelf.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT
			else
				objACESelf.AceType = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
			end if
  			objACESelf.Flags = ADS_ACEFLAG_OBJECT_TYPE_PRESENT
  			objACESelf.objectType = CHANGE_PASSWORD_GUID
  			objACESelf.AccessMask = ADS_RIGHT_DS_CONTROL_ACCESS
  			objDACL.AddAce objACESelf
		End If

		If blnEveryone = False Then
			' Create the ACE for Everyone.
  			Set objACEEveryone = CreateObject("AccessControlEntry")
  			objACEEveryone.Trustee = "Everyone"
  			objACEEveryone.AceFlags = 0
			If Value then
  				objACEEveryone.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT
			else
				objACEEveryone.AceType = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT
			end if
  			objACEEveryone.Flags = ADS_ACEFLAG_OBJECT_TYPE_PRESENT
  			objACEEveryone.objectType = CHANGE_PASSWORD_GUID
  			objACEEveryone.AccessMask = ADS_RIGHT_DS_CONTROL_ACCESS
  			objDACL.AddAce objACEEveryone
		End If

		objSecDescriptor.discretionaryACL = Reorder(objDACL)
 