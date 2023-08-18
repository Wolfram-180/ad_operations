# PURPOSE: change home folder location

# APPROACH: script runs through list of emails
# find user by email
# change home folder location

import secure_data
import ldap3

server = ldap3.Server(secure_data.server)
conn = ldap3.Connection(server, user=secure_data.admin_login, password=secure_data.admin_pass, auto_bind=True, authentication=ldap3.NTLM)

# users_list = ['firstname.lastname@domain.com',
#'firstname.lastname@domain.com',]

# search_base = 'OU=EN,OU=Locations,DC=companydomain,DC=local'

# script runs through list of emails
for user_email in secure_data.users_list_change_homefolder:
    search_filter = '(&(mail='+user_email+'))'
    attrs = ["*"]

    # find user by email
    conn.search(secure_data.search_base, search_filter, attributes=attrs)

    # change home folder location
    if len(conn.entries) > 0:
        user_cn = conn.entries[0].distinguishedName.value

        try:
            conn.modify(user_cn, {'homeDrive': (ldap3.MODIFY_REPLACE, 'c:')})
            conn.modify(user_cn, {'homeDirectory': (ldap3.MODIFY_REPLACE, 'temp')})
            print(user_cn, ' changed')
        except:
            print(user_cn, ' error on change')

    else:
        print(user_email, ' not found')

conn.unbind()

print('DONE')