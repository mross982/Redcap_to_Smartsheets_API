#******************************************************************************************
#                      RUN SETTINGS
#******************************************************************************************
testing = False # When true, selects the test_dict_enterprise_id test environment in Smartsheets


#******************************************************************************************
#                      REDCAP DEPARTMENT GLOBAL VARIABLES
#******************************************************************************************
redcap_projects = ['LDFR Inquiries']
# These items correlate to two different Redcap projects as identified in the api_tokens.txt
# file and in UPDATE.__init__()


#******************************************************************************************
#                      REDCAP DEPARTMENT GLOBAL VARIABLES
#******************************************************************************************

dict_pi_dept_convert = {

}

dict_ahe_dept_convert = {

}

dict_ci_dept_convert = {

}


# Fields in REDCAP with drop down options that need to be converted back to text data
dict_rc_dropdown_convert = {
'request_type': {'1':'Invite New Leader to LDFR', '2': 'Nominate an associate to attend one LDFR who does not meet criteria routinely', '3': 'General question about LDFR'},
'criteria_met': {'1': 'Yes', '2': 'No'}
}

# Converts REDCAP field names to SmartSheet field name description
dict_field_convert = {
'record_id': 'record_id',
'ldfr_inquiries_timestamp': 'ldfr_inquiries_timestamp',
'email': 'email',
'request_type': 'request_type',
'invitee_email': 'invitee_email',
'criteria_met': 'criteria_met',
'ldfr_inquiries_complete': 'ldfr_inquiries_complete',
'associate_role': 'associate_role',
'explain_why': 'explain_why',
'comments': 'comments'
}

ls_required_fields=[]


#**************************************************************************************************
#                     SMARTSHEET SHEET IDs GLOBAL VARIABLES
#**************************************************************************************************

dict_enterprise_id = {'workspace_id': '2658814102136708'}
# dev_dict_enterprise_id = {'workspace_id': '6722481303119748'} # DEVELOPMENT Texas PMO
test_dict_enterprise_id = {'workspace_id': '7259807481653124'} # TEST Texas PMO
project_plans_folder_pathway = ['Project Plans']
portfolio_folder_pathway = ['Portfolio Dashboards']
department_folder_pathway = ['Department Dashboards']
project_details_sheetname = 'LDFR Inquiries'

#**************************************************************************************************
#                     Initial Update
#**************************************************************************************************

smartsheet_record_number = {'project_details': 'LDFR Inquiries'}

starting_record_number = 0


#***************************************************************************
#           Project Details Update
#***************************************************************************
# Converts the field name on all updates
project_details_field_convert = {}
# Converts values
project_details_value_convert = {
    }

# Default values added to the project details sheet when there is no project plan
project_details_no_project_plan_defaults = {
}

# These cells are links to an object given 'SS Field Name': {'type of object (Sheet or Dashboard)': [path, to, object]}
project_details_obj_links = {
}

# Fields in Projects Details with cell links to a project plan
project_details_cell_links = []


#****************************************************************************
#         ALL PROJECT PLANS
#****************************************************************************

primary_row_color = ',,1,,,,,,,23,,,,,,'



#***************************************************************************
#           ALL COMMON SET UP
#***************************************************************************
# Default number of child rows
all_common_phases = 5

# Defaults for all lines in the all common record before filtering
all_common_defaults = {'Report Project Name': 'Project Name', '% Complete': 0, 'Today': '=Today()'}
all_common_timeline_defaults_from_record = ['Report Project Name']

# Fields filter
all_common_fields = ['Dashboard Display','Project Name', 'Project ID', 'Project Health', 'Department', 'Risks', 'Report Project Name', 'Today']

all_common_primary_row_defaults = {'Risks': '=IF(COUNTIF(CHILDREN(), 1) >= 1, 1, 0)', 'Dashboard Display': 'Portfolio', 'Project Health': 'Green'}

# child row fields are defined on SmartCap row 686

#*****************************************************************************
#            PROJECT PLAN SETUP
#****************************************************************************
# Name of the template project plan
project_plan_sheet = 'Project Plan'

# Project row is the top row of a project plan
ls_project_plan_primary_row = ['Project Name', 'Department', 'Report Project Name', 'Project ID']
dict_project_plan_primary_row_defaults = {'Today': '=Today()'}

# Phase rows is the second level indention rows
ls_project_plan_secondary_row = ['Report Project Name', 'Department']

# Details rows are the third level indention rows
ls_project_plan_tertiary_row = ['Report Project Name', 'Department']
dict_project_plan_tertiary_row_defaults = {'% Complete': 0}
# Given a row number, it will update the list of fields provided
project_plan_dates = {'3': ['Start Date', 'Duration']} # {'row number': [list, of, fields]}


#******************************************************************************
#            CELL UPDATES
#******************************************************************************
formula_cells = ['Today', 'Risks', 'Updates']
