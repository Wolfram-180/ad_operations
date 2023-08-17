# script runs through list of emails
# find user by email
# moving users from All to Separated

import secure_data
import ldap3

server = ldap3.Server(secure_data.server)
conn = ldap3.Connection(server, user=secure_data.admin_login, password=secure_data.admin_pass, auto_bind=True, authentication=ldap3.NTLM)

# users_list = [
#'firstname.lastname@domain.com',
#'firstname.lastname@domain.com',
#]
# search_base = 'OU=EN,OU=Locations,DC=companydomain,DC=local'
# add_target_group = 'CN=Separated,OU=Groups,OU=London,OU=EN,OU=Locations,DC=companydomain,DC=local'
# delete_target_group = 'CN=All,OU=Groups,OU=London,OU=EN,OU=Locations,DC=companydomain,DC=local'

# script runs through list of emails
for user_email in secure_data.users_list:
    search_filter = '(&(mail='+user_email+'))'
    attrs = ["*"]

    # find user by email
    conn.search(secure_data.search_base, search_filter, attributes=attrs)

    # moving users from All to Separated
    if len(conn.entries) > 0:
        user_cn = conn.entries[0].distinguishedName.value

        # add to Separated
        conn.extend.microsoft.add_members_to_groups(user_cn, secure_data.add_target_group)

        # delete from All
        conn.extend.microsoft.remove_members_from_groups(user_cn, secure_data.delete_target_group, fix=True)

        print(user_cn, ' moved')
    else:
        print(user_email, ' not found')

conn.unbind()

print('DONE')


