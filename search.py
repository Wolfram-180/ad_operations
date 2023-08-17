# PURPOSE: report by DFS space occupied by active/disabled users profiles

# APPROACH: 2 locations, each has active and disabled groups, so 4 containers
# 2 locations has DFS paths, 1 path per location
# script takes users from 4 containers
# get user`s home directory and calculate space occupied

import secure_data
import ldap3
import os
import csv
from datetime import datetime

# pip install pywin32
import win32com.client as com

is_header_added = False

MB = 1048576 # 1024.0 * 1024.0
fso = com.Dispatch("Scripting.FileSystemObject")

server = ldap3.Server(secure_data.server)
conn = ldap3.Connection(server, user=secure_data.admin_login, password=secure_data.admin_pass, auto_bind=True, authentication=ldap3.NTLM)

loc_mo_act = 'loc_mo_act'
loc_mo_dsbld = 'loc_mo_dsbld'
loc_zh_act = 'loc_zh_act'
loc_zh_dsbld = 'loc_zh_dsbld'

loc_list = [loc_mo_act, loc_mo_dsbld, loc_zh_act, loc_zh_dsbld]

# script takes users from 4 containers
for location in loc_list :
    if location == loc_mo_act:
        base_path_ = secure_data.base_path_msk
        search_base_ = secure_data.search_base_msk_active
    if location == loc_mo_dsbld:
        base_path_ = secure_data.base_path_msk
        search_base_ = secure_data.search_base_msk_disabled    
    if location == loc_zh_act:
        base_path_ = secure_data.base_path_zh
        search_base_ = secure_data.search_base_zh_active
    if location == loc_zh_dsbld:
        base_path_ = secure_data.base_path_zh
        search_base_ = secure_data.search_base_zh_disabled        

    conn.search(
        search_base = search_base_,
        search_filter = '(objectClass=user)',
        search_scope = 'LEVEL',
        attributes = ['member']
    )

    # run by users found
    for entry in conn.entries:
        # get by-user entity
        member = entry.entry_dn
        conn.search(
            search_base = search_base_,
            search_filter=f'(distinguishedName={member})',
            attributes=[
                'displayName', 'mail', 'userAccountControl', 'sAMAccountName', 'homeDirectory'
            ]
        )

        # read properties
        try:
            user_sAMAccountName = conn.entries[0].sAMAccountName.values[0]
        except:
            user_sAMAccountName = ''
        
        try:
            email = conn.entries[0].mail.values[0]
        except:
            email = ''
        
        try:
            displayName = conn.entries[0].displayName.values[0]
        except:
            displayName = ''
        
        try:
            userAccountControl = conn.entries[0].userAccountControl.values[0]
        except:
            userAccountControl = ''        

        try:
            homeDirectory = conn.entries[0].homeDirectory.values[0]
        except:
            homeDirectory = ''

        try:
            is_path_exists = os.path.isdir(homeDirectory)
        except:
            is_path_exists = False

        print(user_sAMAccountName)
        print(homeDirectory)
        
        # if profile path defined - calculate space occupied
        folder_size = -1
        if is_path_exists:
            try:
                folder = fso.GetFolder(homeDirectory)
                folder_size = folder.Size / MB
            except:
                folder_size = 'Error on GetFolder'

        print(folder_size)
        

        now = datetime.now()
        dt_string = now.strftime("%d/%m/%Y %H:%M:%S")

        ### status definition by code
        # https://learn.microsoft.com/en-US/troubleshoot/windows-server/identity/useraccountcontrol-manipulate-account-properties
        userStatus = 'undefined'
        if userAccountControl == 512 :
            userStatus = 'Active with password expiration'
        if userAccountControl == 514 :
            userStatus = 'Disabled'
        if userAccountControl == 66048 :
            userStatus = 'Active, no password expiration'
        ###

        # prepare and add line-per-user in report
        header = ['Location', 'User', 'Folder path', 'Folder found', 'Folder size (MB)', 'Datetime read', 'Email', 'Display name', 'Status code', 'Is active (status)']
        data = [location, user_sAMAccountName, homeDirectory, is_path_exists, folder_size, dt_string, email, displayName, userAccountControl, userStatus]

        f = open('user_data.csv', 'a', encoding='UTF8', newline='')
        writer = csv.writer(f)
        if (is_header_added == False) :
            writer.writerow(header)    
            is_header_added = True
        writer.writerow(data)
        f.close()