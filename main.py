import smartcap.SmartCap as sc

for project in sc.SmartCap_CONFIG.redcap_projects:
    update = sc.UPDATE(project)

    print(project + ' - ' + 'New Projects' + ' - ' + 'Update')
    update.new_records()
    
    # print(project + ' - ' + 'Closed Projects' + ' - ' + 'Update')
    # update.closed_records()
