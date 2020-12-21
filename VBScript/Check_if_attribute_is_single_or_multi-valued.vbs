ldapName = inputbox("Please enter the name of the attribute")
if ldapName = "" then wscript.quit

set objAttribute = GetObject("LDAP://" & ldapName & ",schema") 

if objAttribute.MultiValued then 
    wscript.echo ldapName & " is a multivalued attribute" 
else 
    wscript.echo ldapName & " is a singlevalued attribute" 
end if 