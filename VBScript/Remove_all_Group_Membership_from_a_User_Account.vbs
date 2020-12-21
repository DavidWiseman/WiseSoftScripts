On Error Resume Next

Const ADS_PROPERTY_DELETE = 4
Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D
 
Set objUser = GetObject _
    ("LDAP://cn=david.wiseman,cn=users,dc=wisesoft,dc=co,dc=uk") 
arrMemberOf = objUser.GetEx("memberOf")
 
If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
    WScript.Echo "This account is not a member of any security groups."
    WScript.Quit
End If
 
For Each Group in arrMemberOf
    Set objGroup = GetObject("LDAP://" & Group) 
    objGroup.PutEx ADS_PROPERTY_DELETE, _
        "member", Array("cn=david.wiseman,cn=users,dc=wisesoft,dc=co,dc=uk")
    objGroup.SetInfo
Next
