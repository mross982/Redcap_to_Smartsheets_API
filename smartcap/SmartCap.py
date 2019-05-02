#************************************************************************************
#  - Process:
#       1. UPDATE.new_records(): Updates SmartSheet with REDCAP data
#
#       2. UPDATE.closed_projects(): In Progress
#
#   - Author: Michael WIlliams
#   - Date 07/6/2018
#   - Last Modified: 12/14/2018
#


import smartsheet
from requests import post
import sys
from pathlib import Path
import copy
from pprint import pprint
import numpy as np
import datetime as dt
import time
import os
import inspect
import smartcap.SmartCap_CONFIG as SmartCap_CONFIG

if SmartCap_CONFIG.testing:
        print('-- TEST MODE --')

def object_troubleshoot(obj):

    for attr in dir(obj):
        try:
            print("obj.%s = %r" % (attr, getattr(obj, attr)))
        except:
            pass

def importable_modules(search_path):

    import pkgutil
    sys.path.insert(0,search_path) # add the module's directory to the sys.path()
    all_modules = [x[1] for x in pkgutil.iter_modules(path=search_path)]
    pprint(all_modules)


class REDCAP(object):

    '''
    The REDCAP object serves as a parent class for methods that apply only to the REDCAP program
    '''

    def dataPull(self):
        '''
        API that pulls all data from Redcap
        '''

        API_KEY = self.redcap_department_token

        try:
            payload = {'token': API_KEY, 'format':'json', 'content': 'record', 'type': 'flat'}
            response = post(self.redcap_url, data=payload)
            json_data = response.json()
        except:
            print('REDCAP API Failed')
            sys.exit(0)

        if len(json_data) >= 1:
            if (isinstance(json_data, dict)) and ('You do not have permissions to use the API' in json_data.values()):
                print('Check permissions to use Redcap API')
                sys.exit(0)


        return json_data


    def get_all_records(self, raw = False):
        '''
        Gets all records and all fields THAT ARE NOT BLANK, then corrects for any drop down options
        '''
        intake_dept = str()
        ls_records = []

        json_data = REDCAP.dataPull(self)

        prinrec = False

        if raw == False:
            for record in json_data:
                intake_dept = 'N/A'
                # FIRST DETERMINE WHICH INTAKE FORM IS BEING PULLED FOR TRANSLATION
                # if record['performance_improvement_intake_form_complete'] == '2':
                #     intake_dept = 'PI'
                # elif record['analytics_intake_form_complete'] == '2':
                #     intake_dept = 'Analytics'
                # else:
                #     try:
                #         if record['test_intake_form_complete'] == '2':
                #             intake_dept = 'CI'
                #     except:
                #         continue
                # if intake_dept == '':
                #     continue

                # NEXT CONVERT DEPT INTAKE VARIABLES TO STANDARD SMARTSHEET VARIABLES
                intermediate_data = {}
                for key, value in record.items():
                    if value == '': # If value is blank, skip to the next field
                        continue
                    if (intake_dept == 'PI') and (key in SmartCap_CONFIG.dict_pi_dept_convert):
                        intermediate_data[SmartCap_CONFIG.dict_pi_dept_convert[key]] = value
                    elif (intake_dept == 'Analytics') and (key in SmartCap_CONFIG.dict_ahe_dept_convert):
                        intermediate_data[SmartCap_CONFIG.dict_ahe_dept_convert[key]] = value
                    elif (intake_dept == 'CI') and (key in SmartCap_CONFIG.dict_ci_dept_convert):
                        intermediate_data[SmartCap_CONFIG.dict_ci_dept_convert[key]] = value
                    else:
                        intermediate_data[key] = value

                # THEN CONVERT ANY INDEXED RESPONSE OPTIONS INTO THE TEXT VALUE
                converted_data = {}
                for key, value in intermediate_data.items():
                    if key in list(SmartCap_CONFIG.dict_rc_dropdown_convert): # IF the field is a drop down menu
                        dict_options = SmartCap_CONFIG.dict_rc_dropdown_convert[key]

                        for index, option in dict_options.items():
                            if value == index:
                                converted_data[key] = option
                                break
                    else:
                        converted_data[key] = value

                ls_records.append(converted_data)

        elif raw == True:
            ls_records = json_data

        return ls_records


    def get_select_records(self, ls_project_ids, raw = False):
        '''
        receives a list of project ids and returns the full record for each project ID. If raw is true, the data returned as unprocessed from Redcap. If raw if false (default), the fields that have drop down values are converted from the index number to the actual values and field names are converted to their Smartsheet equivalents. See dict_rc_dropdown_convert and dict_field_convert in the settings file for these conversions.
        '''
        if raw == True:
            ls_select_data = []
            raw_data = REDCAP.get_all_records(self, raw = True)

            if self.update == 'new':
                for record in raw_data:
                    if record['record_id'] in ls_project_ids: # Add data with correct ID to new list
                        ls_select_data.append(record)
            else:
                for record in raw_data:
                    if record['record_id'] in ls_project_ids: # Add data with correct ID to new list
                        ls_select_data.append(record)

        elif raw == False:

            raw_data = REDCAP.get_all_records(self)

            ls_raw_data = []
            ls_select_data = []

            for record in raw_data:
                if record['record_id'] in ls_project_ids: # Add data with correct ID to new list
                    ls_raw_data.append(record)

            for record in ls_raw_data:
                refined_record = {}
                # if self.update == 'new':
                #     try:
                #         refined_record = REDCAP.convert_to_duration(record)
                #     except KeyError:
                #         print('Could not convert date on record', str(record['record_id']))
                #         continue
                for k, v in record.items():
                    if k in list(SmartCap_CONFIG.dict_field_convert): # convert Redcap field names to Smartsheet field names
                        refined_record[SmartCap_CONFIG.dict_field_convert[k]] = v


                # if record['performance_improvement_intake_form_complete'] == 'Complete':
                #     refined_record['Intake Form'] = 'PI'
                # elif record['analytics_intake_form_complete'] == 'Complete':
                #     refined_record['Intake Form'] = 'Analytics'
                # elif record['test_intake_form_complete'] == 'Complete':
                #     refined_record['Intake Form'] = 'CI'
                # else:
                #     continue

                ls_select_data.append(refined_record)

        return ls_select_data


    def get_project_ids(self):
        '''
        returns any project ID's that have been identified as needing to update in smartsheets
        '''

        # get all records then select just open projects
        ls_project_ids = []
        json_data = REDCAP.get_all_records(self)

        for record in json_data:

            if self.update == 'new':
                try:
                    if record['close'] == 'Yes':
                        continue
                except:
                    pass

                if (self.redcap_form == 'long_form') and  (int(record['record_id']) > SmartCap_CONFIG.starting_record_number):
                    project_id = record.get('record_id', 'None')
                    ls_project_ids.append(project_id)
                elif (self.redcap_form == 'shortcut'):
                    project_id = record.get('record_id', 'None')
                    ls_project_ids.append(project_id)

            elif self.update == 'closed':
                if self.close_program_search == 'Redcap':
                    try:
                        if record['close'] == 'Yes':
                            # if intake is complete, project is not closed and project was accepted
                            project_id = record.get('record_id', 'None')
                            ls_project_ids.append(project_id)
                    except:
                        pass
                elif self.close_program_search == 'Smartsheet':
                    try:
                        if record['close'] != 'Yes':
                            # if intake is complete, project is not closed and project was accepted
                            project_id = record.get('record_id', 'None')
                            ls_project_ids.append(project_id)
                    except:
                        pass

        return ls_project_ids


    def convert_to_duration(record):
        '''
        Because Smartsheet does not allow the user to enter start and finished dates and only allows one date and a duration value, a new Duration field must be created and added to the record.
        '''
        refined_record = {}
        start = record['start_date']
        end = record['end_date']
        # add RAISEERROR here in the event the start date does not occur before the end date.
        days = np.busday_count( start, end )
        days = days + 1
        refined_record['Duration'] = str(days)

        return refined_record


class SMARTSHEET(object):
    '''
    The SMARTSHEET class is a parent class for all the methods that apply only to the SmartSheets program.
    '''

    def get_client(self):
        '''
        given the token object self, returns the smartsheet client object
        '''
        ss = smartsheet.Smartsheet(self.smartsheet_token)
        ss.errors_as_exceptions(False)

        return ss


    def get_cell_display_value_by_column_name(row, column_map, str_column_name):
        '''
        given a row object, column map, and the name of a column, returns cell display value
        '''
        column_id = column_map[str_column_name]
        obj_cell = row.get_column(column_id)
        cell_display_value = obj_cell.display_value

        return cell_display_value


    def get_cell_value_by_column_name(row, column_map, str_column_name):
        '''
        given a row object, column map, and the name of a column, returns cell value
        '''
        column_id = column_map[str_column_name]
        obj_cell = row.get_column(column_id)
        cell_value = obj_cell.value

        return cell_value


    def get_column_map(sheet):
        column_map = {}
        for column in sheet.columns:
            column_map[column.title] = column.id

        return column_map


    def get_row_list(sheet):
        row_list = []
        for row in sheet.rows:
            row_list.append(row.id)
        return row_list


    def get_project_ids(self):
        '''
        returns a list of all Project IDs (Record_id) from the project details sheet.
        '''

        for k, v in SmartCap_CONFIG.smartsheet_record_number.items():
            sheet_name = k
            str_filename = v

        ls_existing_SS_ids = []
        dict_ids = SMARTSHEET.get_sheet_id(self, str_filename) # GET SHEET ID FOR PROJECT DETAILS SHEET

        project_details_sheet_id = dict_ids[str_filename]

        ss = SMARTSHEET.get_client(self)
        sheet = ss.Sheets.get_sheet(project_details_sheet_id)
        column_map = SMARTSHEET.get_column_map(sheet)

        for row in sheet.rows:
            project_id = SMARTSHEET.get_cell_display_value_by_column_name(row, column_map, 'record_id')
            # closed = SMARTSHEET.get_cell_value_by_column_name(row, column_map, 'Project Closed')
            if self.update == 'new':
                if project_id is not None:
                    ls_existing_SS_ids.append(project_id)
            else:
                if self.close_program_search == 'Redcap':
                    if (project_id is not None):
                        ls_existing_SS_ids.append(project_id)
                elif self.close_program_search == 'Smartsheet':
                    if (project_id is not None) and (closed == True):
                        if any(char.isdigit() for char in project_id):
                            ls_existing_SS_ids.append(project_id)

        return ls_existing_SS_ids


    def get_select_records(self, ls_record_ids):

        for k, v in SmartCap_CONFIG.smartsheet_record_number.items():
            sheet_name = k
            str_filename = v

        ls_records = []
        dict_ids = SMARTSHEET.get_sheet_id(self, str_filename) # GET SHEET ID FOR PROJECT DETAILS SHEET
        project_details_sheet_id = dict_ids[str_filename]

        ss = SMARTSHEET.get_client(self)
        sheet = ss.Sheets.get_sheet(project_details_sheet_id)
        column_map = SMARTSHEET.get_column_map(sheet)

        for row in sheet.rows:
            project_id = SMARTSHEET.get_cell_display_value_by_column_name(row, column_map, 'Project ID')
            if project_id in ls_record_ids:
                new_record = {}
                new_record['row_id'] = row.id
                for cell in row.cells:
                    if cell.value == None:
                        continue
                    else:
                        for k, v in column_map.items():
                            if cell.column_id == v:
                                new_record[k] = cell.value

                ls_records.append(new_record)

        return ls_records


    def get_project_names(self):
        '''
        returns a list of all Project Names from the project details sheet.
        '''
        ls_existing_SS_names = []
        sheet_name = SmartCap_CONFIG.project_details_sheetname
        sheet_id = dict_enterprise_sheet_ids[sheet_name]
        ss = SMARTSHEET.get_client(self)
        sheet = ss.Sheets.get_sheet(sheet_id)
        column_map = SMARTSHEET.get_column_map(sheet)

        for row in sheet.rows:
            project_name = SMARTSHEET.get_cell_display_value_by_column_name(row, column_map, 'Project Name')

            if project_name is not None:
                ls_existing_SS_names.append(project_name)

        return ls_existing_SS_names


    def get_project_plan(self, obj_update):
        '''
        Given update object, returns the project plan sheet object
        '''

        if obj_update.sheet_type == 'All Common':
            sheet_name = obj_update.portfolio + ' All Comm'
        else:
            sheet_name = obj_update.project_name

        folder_pathway = SmartCap_CONFIG.project_plans_folder_pathway
        folder_pathway.append(obj_update.portfolio) # the portfolio is appended to the file pathway because all project plans are in their respecive portfolio's folder

        dict_folders = SMARTSHEET.get_path_ids(self, folder_pathway)

        project_plans_f_id = dict_folders[obj_update.portfolio]

        ss = SMARTSHEET.get_client(self)
        folder = ss.Folders.get_folder(project_plans_f_id)

        for sheet in folder.sheets:
            if sheet_name in sheet.name:
                obj_sheet = sheet

        return obj_sheet


    def get_path_ids(self, ls_path_terms):
        '''
        given a list of path terms (i.e. names of folders), returns a dictionary of all folders and their ID numbers.
        Note: Path terms can be sub-strings (i.e. do not have to be identical matches, just part the folder name)
        Note: If other objects (e.g. sheets/sights) are in the path terms, they will be omitted from the returned dictionary
        '''

        dict_path = {}

        ss = SMARTSHEET.get_client(self)
        workspace = ss.Workspaces.get_workspace(self.workspace_id)
        for folder in workspace.folders:
            for file_path in ls_path_terms:
                if file_path in folder.name:
                    dict_path[file_path] = folder.id
                    f_id = folder.id
                    break

        if len(ls_path_terms) > 1:
            for _ in range(len(ls_path_terms)-1): # find the terminal folder
                target_folder = ss.Folders.get_folder(f_id)
                for folder in target_folder.folders:
                    for term in ls_path_terms:
                        if term in folder.name:
                            dict_path[term] = folder.id
                            f_id = folder.id
                            break

        return dict_path


    def get_sheet_id(self, str_filename, int_folder_id=None):
        '''
        Given the sheet name and folder ID where it is located, returns a dictionary with full sheet name and sheet ID
        Note: given sheet name is matched as a substring of the actual sheet.name (i.e. it does not have to be an exact match)
        '''

        dict_path = {}
        ss = SMARTSHEET.get_client(self)
        if int_folder_id == None:
            folder = ss.Workspaces.get_workspace(self.workspace_id)
        else:
            folder = ss.Folders.get_folder(int_folder_id)

        for sheet in folder.sheets:
            if str_filename in sheet.name:
                dict_path[sheet.name] = sheet.id

        return dict_path


    def get_dashboard_id(self, str_dashboard_name, int_folder_id):
        '''
        Given the dashboard name and folder ID where it is located, returns a dictionary with full dashboard name and object ID
        Note: given dashboard name is matched as a substring of the actual dashboard.name (i.e. it does not have to be an exact match)
        '''

        dashboard_id = str()
        ss = SMARTSHEET.get_client(self)
        folder = ss.Folders.get_folder(int_folder_id)

        for db in folder.sights:
            if str_dashboard_name in db.name:
                dashboard_id = db.id

        return dashboard_id


    def name_correction(self, str_bad_name, sheet_id):
        import difflib

        ss = SMARTSHEET.get_client(self)
        sheet = ss.Sheets.get_sheet(sheet_id)
        column_map = SMARTSHEET.get_column_map(sheet)

        good_name = ''
        match = 0
        for row in sheet.rows:
            str_actual_project_name = SMARTSHEET.get_cell_display_value_by_column_name(row, column_map, 'Project Name')
            seq=difflib.SequenceMatcher(None, str_bad_name, str_actual_project_name)
            d=seq.ratio()
            if d > match:
                match = d
                good_name = str_actual_project_name

        return good_name


    def project_details_update(self, obj_update, record):
        '''

        '''
        obj_update.sheet_type = 'Project Details'
        obj_update.project_plan = 'No Project Plan'
        obj_update.intake_dept = None
        obj_update.portfolio = None
        obj_update.record_id = record['record_id']

        print('...project details update', obj_update.record_id)

        dict_ids = SMARTSHEET.get_sheet_id(self, SmartCap_CONFIG.project_details_sheetname) # GET SHEET ID FOR PROJECT DETAILS SHEET
        obj_update.sheet_id = dict_ids[SmartCap_CONFIG.project_details_sheetname]

        obj_update.parent_row_id = OBJ_UPDATE.new_row(self, obj_update)

        result = OBJ_UPDATE.cell(self, obj_update, record)

        if obj_update.project_plan != 'No Project Plan':
            print('...linking Project Plan to projects details')
            result = OBJ_UPDATE.cell_link(self, obj_update)


    def project_plan_update(self, obj_update, record):
        '''

        '''
        if record['Project Plan'] == 'All Common': # FOR PROJECTS GOING INTO AN ALL COMMON TIMELINE SHEET

            obj_update.sheet_type = record['Project Plan'] # All Common Sheet Type
            obj_update.project_plan = record['Project Plan'] # carried over to project details update
            obj_update.portfolio = record['Portfolio']
            obj_update.department = record['Department']
            obj_update.record_id = record['Project ID']

            project_plan = SMARTSHEET.get_project_plan(self, obj_update)
            obj_update.project_plan_id = project_plan.id
            print('...all common project plan update', obj_update.record_id)

            result = SMARTSHEET.common_project_plan(self, obj_update, record)

            return obj_update


        elif record['Project Plan'] == 'Project Plan': # FOR PROJECTS THAT CREATE THEIR OWN PROJECT PLAN

            obj_update.sheet_type = record['Project Plan'] # Project Plan Sheet Type
            obj_update.project_plan = record['Project Plan'] # carried over to project details update
            obj_update.portfolio = record['Portfolio']
            obj_update.department = record['Department']
            obj_update.record_id = record['Project ID']
            obj_update.project_name = record['Project Name']

            print('...new project plan update', obj_update.record_id)

            obj_project_plan = SMARTSHEET.create_project_plan(self, obj_update)
            obj_update.sheet_id = obj_project_plan.id
            obj_update.project_plan_id = obj_project_plan.id # carried over to project details update

            result = SMARTSHEET.new_project_plan(self, obj_update, record)

            return obj_update

        elif record['Project Plan'] == 'No Project Plan': # FOR PROJECTS WITH NO PROJECT PLAN

            obj_update.sheet_type = record['Project Plan']

            return obj_update


    def create_project_plan(self, obj_update):
        '''
        Copies the template project plan and saves it to the appropriate folder before updating information
        '''

        project_id = obj_update.record_id
        if (len(obj_update.project_name) + len(project_id)) > 50: # B/c Smartsheet has a 50 character limit on names, reduce name conv here.
            x = len(project_id)
            y = 45 - x
            project_name = project_id + ' ' + obj_update.project_name[:y]
        else:
            project_name = project_id + ' ' + obj_update.project_name

        folder_pathway = SmartCap_CONFIG.project_plans_folder_pathway
        folder_pathway.append(obj_update.portfolio)

        dict_folders = SMARTSHEET.get_path_ids(self, folder_pathway)

        project_plans_f_id = dict_folders[obj_update.portfolio] # GET PORTFOLIO PROJECT PLAN/TIMELINE FOLDER

        ss = SMARTSHEET.get_client(self)
        dest_folder = ss.Folders.get_folder(project_plans_f_id)

        folder_pathway = ['Templates']
        dict_folders = SMARTSHEET.get_path_ids(self, folder_pathway)

        template_folder_id = dict_folders['Templates'] # GET TEMPLATES FOLDER ID

        folder = ss.Folders.get_folder(template_folder_id)
        for sheet in folder.sheets:
            if SmartCap_CONFIG.project_plan_sheet in sheet.name: # GET THE ALL COMMON TIMELINE SHEET ID
                obj_sheet = sheet

        response = ss.Sheets.copy_sheet(obj_sheet.id, ss.models.ContainerDestination({'destination_type': 'folder','destination_id': dest_folder.id, 'new_name': project_name }), include='all')

        obj_sheet = response.result

        return obj_sheet


    def common_project_plan(self, obj_update, record):

        obj_update.row_type = 'primary'
        obj_update.primary_row_id = OBJ_UPDATE.new_row(self, obj_update)

        result = OBJ_UPDATE.cell(self, obj_update, record)

        obj_update.row_type = 'secondary'
        for phase_number in range(SmartCap_CONFIG.all_common_phases):
            record['phase_number'] = phase_number
            obj_update.secondary_row_id = OBJ_UPDATE.new_row(self, obj_update)
            result = OBJ_UPDATE.cell(self, obj_update, record)

        return result


    def new_project_plan(self, obj_update, record):

        ss = SMARTSHEET.get_client(self)
        sheet = ss.Sheets.get_sheet(obj_update.sheet_id)
        column_map = SMARTSHEET.get_column_map(sheet)

        for row in sheet.rows:
            obj_update.row_number = str(row.row_number)
            if row.row_number == 1:
                obj_update.row_type = 'primary'
                obj_update.primary_row_id = row.id

            elif row.parent_id == obj_update.primary_row_id:
                obj_update.row_type = 'secondary'
                obj_update.secondary_row_id = row.id

            elif row.parent_id == obj_update.secondary_row_id:
                obj_update.row_type = 'tertiary'
                obj_update.tertiary_row_id = row.id

            result = OBJ_UPDATE.cell(self, obj_update, record)

        return result


    # def no_project_plan_defaults(record):

    #     new_record = {}
    #     new_record = dict(record, **SmartCap_CONFIG.project_details_no_project_plan_defaults)

    #     return new_record


    def all_common_parent_row(record):

        new_record = {}
        for k, v in record.items():
            if k in SmartCap_CONFIG.all_common_fields:
                new_record[k] = v
            if k in SmartCap_CONFIG.all_common_timeline_defaults_from_record:
                new_record[k] = record[v]

        return new_record


    def all_common_child_row(record):

        child_update_record = {}
        child_update_record['Dashboard Display'] = 'No Display'
        child_update_record['Project ID'] = record['Project ID']
        child_update_record['Project Name'] = record['Project Name'] + ' example phase/milestone' + str(record['phase_number']+1)
        child_update_record['Report Project Name'] = record['Project Name']
        child_update_record['Portfolio'] = record['Portfolio']
        child_update_record['Start Date'] = record['Start Date']
        child_update_record['Duration'] = record['Duration']
        child_update_record['% Complete'] = record['% Complete']
        child_update_record['Department'] = record['Department']
        child_update_record['Risks'] = '=IF(OR([Issues and Risks]'+str(record['phase_number']+2)+' = "", [Remove Issue/Update from Dashboard]'+str(record['phase_number']+2)+' = true), 0, 1)'
        child_update_record['Updates'] = '=IF(OR([Update]'+str(record['phase_number']+2)+' = "", [Remove Issue/Update from Dashboard]'+str(record['phase_number']+2)+' = true), 0, 1)'

        return child_update_record


    # def timeline_only_row(record):

    #     only_row_record = {}
    #     for k, v in record.items():
    #         if k != 'Finished Date':
    #             only_row_record[k] = v

    #     return only_row_record


    def primary_defaults(record):

        new_record = {}
        for k, v in record.items():
            if k in SmartCap_CONFIG.ls_project_plan_primary_row:
                new_record[k] = v
        new_record['Report Project Name'] = record['Project Name']
        new_record = dict(new_record, **SmartCap_CONFIG.dict_project_plan_primary_row_defaults)

        return new_record


    def secondary_defaults(record):

        new_record = {}
        for k, v in record.items():
            if k in SmartCap_CONFIG.ls_project_plan_secondary_row:
                new_record[k] = v
        new_record['Report Project Name'] = record['Project Name']

        return new_record


    def tertiary_defaults(record, date_update=False):

        new_record = {}
        for k, v in record.items():
            if k in SmartCap_CONFIG.ls_project_plan_tertiary_row:
                new_record[k] = v
        new_record['Report Project Name'] = record['Project Name']
        new_record = dict(new_record, **SmartCap_CONFIG.dict_project_plan_tertiary_row_defaults)

        if date_update != False:
            for k, v in SmartCap_CONFIG.project_plan_dates.items():
                if k == date_update:
                    ls_fields = v
            for field in ls_fields:
                new_record[field] = record[field]

        return new_record


    def new_project_row(self, obj_update, record):

        revised_record = {}

        for k, v in record.items():
            value_converted = False
            if obj_update.project_plan == 'No Project Plan':
                revised_record = dict(revised_record, **SmartCap_CONFIG.project_details_no_project_plan_defaults)

            if k in SmartCap_CONFIG.project_details_field_convert: # Change finshed date to due date to preserve intial est.
                revised_record[SmartCap_CONFIG.project_details_field_convert[k]] = v

            if k in SmartCap_CONFIG.project_details_value_convert:
                options = SmartCap_CONFIG.project_details_value_convert[k]
                revised_record[k] = options[v]
                value_converted = True

            if value_converted == False:
                if v == '':
                    continue
                else:
                    revised_record[k] = v

        return revised_record


    def update_smartsheets(self, obj_update, close_record):
        '''
        Given a list of record numbers, this function deletes the records from Smartsheets Projects Details sheet.
        '''
        print(obj_update.__dict__)
        print(self.__dict__)
        # sys.exit(0)
        # # START HERE
        # for field, value in close_record.items():
        #     if field in SmartCap_CONFIG.dict_field_convert.values():
        #         print(field)

        # sys.exit(0)
        # print(close_record)
        # print(obj_update.__dict__)
        # sys.exit(0)
        # result = obj_close.delete_project_details_rows(self, close_records)


        # Closed_Project_Update.convert_fields(self, close_records)


    def post_update(self, record):

        ls_record_id = []
        ls_record_id.append(record['Update ID'])
        post_record = REDCAP.get_select_records(self, ls_record_id, raw = True)
        for record in post_record:
            record['time_allocation_complete'] = '2'

        project = redcap.Project(self.redcap_url, self.redcap_allocation_token)
        submit = project.import_records(post_record)


class UPDATE(SMARTSHEET, REDCAP):

    def __init__(self, update=None, redcap_department_token=None, redcap_url=None, smartsheet_token=None,
        workspace_id=None):
        homedir = os.path.expanduser("~")
        api_keys = os.path.join(homedir, 'OneDrive - Ascension','Desktop','LDFR_api_keys.txt')

        if os.path.exists(api_keys) != True:
            api_keys = os.path.join(homedir, 'Desktop','LDFR_api_keys.txt')

        if api_keys == None:
            print('Can\'t find api keys file. Make sure it is on your desktop')

        with open(api_keys, 'r') as f:
            for i, line in enumerate(f):
                if i == 1:
                    RC_department = line.rstrip()
                if i == 4:
                    SSapi = line.rstrip()

        self.redcap_url = 'https://redcap.ascension.org/txaus/api/'
        self.redcap_department_token = RC_department
        self.redcap_form = 'long_form'
        self.smartsheet_token = SSapi


    def new_records(self):
        '''
        Starting point for all new record updates.
        '''

        new_records = UPDATE.find_new_records(self)

        if len(new_records) > 0:
            for record in new_records:
                obj_update = OBJ_UPDATE()
                obj_update.update = self.update

                # obj_update = SMARTSHEET.project_plan_update(self, obj_update, record)

                SMARTSHEET.project_details_update(self, obj_update, record)


    def closed_records(self):
            '''
            - Identifies which sheets will be updated with the closed_projects() update
            '''
            new_records = UPDATE.find_closed_records(self)
            # ADD TROUBLESHOOT STEP HERE TO IDENTIFY ANY MISSING REQUIRED FIELDS!
            if len(new_records) > 0:
                for record in new_records:
                    obj_update = OBJ_UPDATE()
                    obj_update.update = self.update
                    obj_update.close_program_search = self.close_program_search

                    SMARTSHEET.update_smartsheets(self, obj_update, record)


    def find_new_records(self):

        self.update = 'new'

        if SmartCap_CONFIG.testing == True:
            self.workspace_id = SmartCap_CONFIG.test_dict_enterprise_id['workspace_id']
        else: # Live Environment
            self.workspace_id = SmartCap_CONFIG.dict_enterprise_id['workspace_id']

        ls_record_ids = []
        import_record = {}

        ls_SS_proj_ids = SMARTSHEET.get_project_ids(self)
        ls_RC_proj_ids = REDCAP.get_project_ids(self)

        for id_num in ls_RC_proj_ids:
            if id_num not in ls_SS_proj_ids:
                ls_record_ids.append(id_num)

        if len(ls_record_ids) > 0:
            new_records = REDCAP.get_select_records(self, ls_record_ids)
            complete_records = []
            for record in new_records:
                include=True
                for field in SmartCap_CONFIG.ls_required_fields:
                    if field not in record.keys():
                        print('Required fields missing from record ', str(record['Project ID']))
                        include=False
                if include==True:
                    complete_records.append(record)
            print('...new project update with %s new records' %(str(len(complete_records))))
        else:
            complete_records = list()
            print('...No new records found.')

        return complete_records


    def find_closed_records(self):
        self.update = 'closed'

        if SmartCap_CONFIG.testing == True:
            self.workspace_id = SmartCap_CONFIG.test_dict_enterprise_id['workspace_id']
        else:
            self.workspace_id = SmartCap_CONFIG.dict_enterprise_id['workspace_id']

        # ls_programs = ['Redcap', 'Smartsheet']
        ls_programs = ['Smartsheet', 'Redcap']
        close_records = list()
        for program in ls_programs:
            ls_record_ids = []
            self.close_program_search = program
            ls_RC_proj_ids = REDCAP.get_project_ids(self)
            ls_SS_proj_ids = SMARTSHEET.get_project_ids(self)

            if program == 'Redcap':
                for proj_id in ls_RC_proj_ids:
                    if proj_id in ls_SS_proj_ids:
                        ls_record_ids.append(proj_id)

            elif program == 'Smartsheet':
                ls_record_ids = ls_SS_proj_ids

            if len(ls_record_ids) > 0:
                print('....updating %s closed records from %s'%(str(len(ls_record_ids)), program))
                records = SMARTSHEET.get_select_records(self, ls_record_ids)
                close_records.extend(records)
            else:
                print('....no newly closed records from', program)

        return close_records


    # def project_details_update(self, obj_update, record):
    #     '''

    #     '''
    #     obj_update.sheet_type = 'Project Details'
    #     obj_update.project_plan = record['Project Plan']
    #     obj_update.intake_dept = record['Intake Form']
    #     obj_update.portfolio = record['Portfolio']
    #     obj_update.record_id = record['Project ID']

    #     print('...project details update', obj_update.record_id)

    #     dict_ids = SMARTSHEET.get_sheet_id(self, SmartCap_CONFIG.project_details_sheetname) # GET SHEET ID FOR PROJECT DETAILS SHEET
    #     obj_update.sheet_id = dict_ids[SmartCap_CONFIG.project_details_sheetname]

    #     obj_update.parent_row_id = OBJ_UPDATE.new_row(self, obj_update)

    #     result = OBJ_UPDATE.cell(self, obj_update, record)

    #     if obj_update.project_plan != 'No Project Plan':
    #         print('...linking Project Plan to projects details')
    #         result = OBJ_UPDATE.cell_link(self, obj_update)


    # def project_plan_update(self, obj_update, record):
    #     '''

    #     '''
    #     if record['Project Plan'] == 'All Common': # FOR PROJECTS GOING INTO AN ALL COMMON TIMELINE SHEET

    #         obj_update.sheet_type = record['Project Plan'] # All Common Sheet Type
    #         obj_update.project_plan = record['Project Plan'] # carried over to project details update
    #         obj_update.portfolio = record['Portfolio']
    #         obj_update.department = record['Department']
    #         obj_update.record_id = record['Project ID']

    #         project_plan = SMARTSHEET.get_project_plan(self, obj_update)
    #         obj_update.project_plan_id = project_plan.id
    #         print('...all common project plan update', obj_update.record_id)

    #         result = UPDATE.common_project_plan(self, obj_update, record)

    #         return obj_update


    #     elif record['Project Plan'] == 'Project Plan': # FOR PROJECTS THAT CREATE THEIR OWN PROJECT PLAN

    #         obj_update.sheet_type = record['Project Plan'] # Project Plan Sheet Type
    #         obj_update.project_plan = record['Project Plan'] # carried over to project details update
    #         obj_update.portfolio = record['Portfolio']
    #         obj_update.department = record['Department']
    #         obj_update.record_id = record['Project ID']
    #         obj_update.project_name = record['Project Name']

    #         print('...new project plan update', obj_update.record_id)

    #         obj_project_plan = UPDATE.create_project_plan(self, obj_update)
    #         obj_update.sheet_id = obj_project_plan.id
    #         obj_update.project_plan_id = obj_project_plan.id # carried over to project details update

    #         result = UPDATE.new_project_plan(self, obj_update, record)

    #         return obj_update

    #     elif record['Project Plan'] == 'No Project Plan': # FOR PROJECTS WITH NO PROJECT PLAN

    #         obj_update.sheet_type = record['Project Plan']

    #         return obj_update


    # def create_project_plan(self, obj_update):
    #     '''
    #     Copies the template project plan and saves it to the appropriate folder before updating information
    #     '''

    #     project_id = obj_update.record_id
    #     if (len(obj_update.project_name) + len(project_id)) > 50: # B/c Smartsheet has a 50 character limit on names, reduce name conv here.
    #         x = len(project_id)
    #         y = 45 - x
    #         project_name = project_id + ' ' + obj_update.project_name[:y]
    #     else:
    #         project_name = project_id + ' ' + obj_update.project_name

    #     folder_pathway = SmartCap_CONFIG.project_plans_folder_pathway
    #     folder_pathway.append(obj_update.portfolio)

    #     dict_folders = SMARTSHEET.get_path_ids(self, folder_pathway)

    #     project_plans_f_id = dict_folders[obj_update.portfolio] # GET PORTFOLIO PROJECT PLAN/TIMELINE FOLDER

    #     ss = SMARTSHEET.get_client(self)
    #     dest_folder = ss.Folders.get_folder(project_plans_f_id)

    #     folder_pathway = ['Templates']
    #     dict_folders = SMARTSHEET.get_path_ids(self, folder_pathway)

    #     template_folder_id = dict_folders['Templates'] # GET TEMPLATES FOLDER ID

    #     folder = ss.Folders.get_folder(template_folder_id)
    #     for sheet in folder.sheets:
    #         if SmartCap_CONFIG.project_plan_sheet in sheet.name: # GET THE ALL COMMON TIMELINE SHEET ID
    #             obj_sheet = sheet

    #     response = ss.Sheets.copy_sheet(obj_sheet.id, ss.models.ContainerDestination({'destination_type': 'folder','destination_id': dest_folder.id, 'new_name': project_name }), include='all')

    #     obj_sheet = response.result

    #     return obj_sheet


    # def common_project_plan(self, obj_update, record):

    #     obj_update.row_type = 'primary'
    #     obj_update.primary_row_id = OBJ_UPDATE.new_row(self, obj_update)

    #     result = OBJ_UPDATE.cell(self, obj_update, record)

    #     obj_update.row_type = 'secondary'
    #     for phase_number in range(SmartCap_CONFIG.all_common_phases):
    #         record['phase_number'] = phase_number
    #         obj_update.secondary_row_id = OBJ_UPDATE.new_row(self, obj_update)
    #         result = OBJ_UPDATE.cell(self, obj_update, record)

    #     return result


    # def new_project_plan(self, obj_update, record):

    #     ss = SMARTSHEET.get_client(self)
    #     sheet = ss.Sheets.get_sheet(obj_update.sheet_id)
    #     column_map = SMARTSHEET.get_column_map(sheet)

    #     for row in sheet.rows:
    #         obj_update.row_number = str(row.row_number)
    #         if row.row_number == 1:
    #             obj_update.row_type = 'primary'
    #             obj_update.primary_row_id = row.id

    #         elif row.parent_id == obj_update.primary_row_id:
    #             obj_update.row_type = 'secondary'
    #             obj_update.secondary_row_id = row.id

    #         elif row.parent_id == obj_update.secondary_row_id:
    #             obj_update.row_type = 'tertiary'
    #             obj_update.tertiary_row_id = row.id

    #         result = OBJ_UPDATE.cell(self, obj_update, record)

    #     return result


    # def no_project_plan_defaults(record):

    #     new_record = {}
    #     new_record = dict(record, **SmartCap_CONFIG.project_details_no_project_plan_defaults)

    #     return new_record


    # def all_common_parent_row(record):

    #     new_record = {}
    #     for k, v in record.items():
    #         if k in SmartCap_CONFIG.all_common_fields:
    #             new_record[k] = v
    #         if k in SmartCap_CONFIG.all_common_timeline_defaults_from_record:
    #             new_record[k] = record[v]

    #     return new_record


    # def all_common_child_row(record):

    #     child_update_record = {}
    #     child_update_record['Dashboard Display'] = 'No Display'
    #     child_update_record['Project ID'] = record['Project ID']
    #     child_update_record['Project Name'] = record['Project Name'] + ' example phase/milestone' + str(record['phase_number']+1)
    #     child_update_record['Report Project Name'] = record['Project Name']
    #     child_update_record['Portfolio'] = record['Portfolio']
    #     child_update_record['Start Date'] = record['Start Date']
    #     child_update_record['Duration'] = record['Duration']
    #     child_update_record['% Complete'] = record['% Complete']
    #     child_update_record['Department'] = record['Department']
    #     child_update_record['Risks'] = '=IF(OR([Issues and Risks]'+str(record['phase_number']+2)+' = "", [Remove Issue/Update from Dashboard]'+str(record['phase_number']+2)+' = true), 0, 1)'
    #     child_update_record['Updates'] = '=IF(OR([Update]'+str(record['phase_number']+2)+' = "", [Remove Issue/Update from Dashboard]'+str(record['phase_number']+2)+' = true), 0, 1)'

    #     return child_update_record


    # def timeline_only_row(record):

    #     only_row_record = {}
    #     for k, v in record.items():
    #         if k != 'Finished Date':
    #             only_row_record[k] = v

    #     return only_row_record


    # def primary_defaults(record):

    #     new_record = {}
    #     for k, v in record.items():
    #         if k in SmartCap_CONFIG.ls_project_plan_primary_row:
    #             new_record[k] = v
    #     new_record['Report Project Name'] = record['Project Name']
    #     new_record = dict(new_record, **SmartCap_CONFIG.dict_project_plan_primary_row_defaults)

    #     return new_record


    # def secondary_defaults(record):

    #     new_record = {}
    #     for k, v in record.items():
    #         if k in SmartCap_CONFIG.ls_project_plan_secondary_row:
    #             new_record[k] = v
    #     new_record['Report Project Name'] = record['Project Name']

    #     return new_record


    # def tertiary_defaults(record, date_update=False):

    #     new_record = {}
    #     for k, v in record.items():
    #         if k in SmartCap_CONFIG.ls_project_plan_tertiary_row:
    #             new_record[k] = v
    #     new_record['Report Project Name'] = record['Project Name']
    #     new_record = dict(new_record, **SmartCap_CONFIG.dict_project_plan_tertiary_row_defaults)

    #     if date_update != False:
    #         for k, v in SmartCap_CONFIG.project_plan_dates.items():
    #             if k == date_update:
    #                 ls_fields = v
    #         for field in ls_fields:
    #             new_record[field] = record[field]

    #     return new_record


    # def new_project_row(self, obj_update, record):

    #     revised_record = {}

    #     for k, v in record.items():
    #         value_converted = False
    #         if obj_update.project_plan == 'No Project Plan':
    #             revised_record = dict(revised_record, **SmartCap_CONFIG.project_details_no_project_plan_defaults)

    #         if k in SmartCap_CONFIG.project_details_field_convert: # Change finshed date to due date to preserve intial est.
    #             revised_record[SmartCap_CONFIG.project_details_field_convert[k]] = v

    #         if k in SmartCap_CONFIG.project_details_value_convert:
    #             options = SmartCap_CONFIG.project_details_value_convert[k]
    #             revised_record[k] = options[v]
    #             value_converted = True

    #         if value_converted == False:
    #             if v == '':
    #                 continue
    #             else:
    #                 revised_record[k] = v

    #     return revised_record


    # def update_smartsheets(self, obj_update, close_record):
    #     '''
    #     Given a list of record numbers, this function deletes the records from Smartsheets Projects Details sheet.
    #     '''
    #     print(obj_update.__dict__)
    #     print(self.__dict__)
    #     sys.exit(0)
    #     # START HERE
    #     for field, value in close_record.items():
    #         if field in SmartCap_CONFIG.dict_field_convert.values():
    #             print(field)

    #     sys.exit(0)
    #     print(close_record)
    #     print(obj_update.__dict__)
    #     sys.exit(0)
    #     result = obj_close.delete_project_details_rows(self, close_records)
    #     # Closed_Project_Update.convert_fields(self, close_records)


    # def post_update(self, record):

    #     ls_record_id = []
    #     ls_record_id.append(record['Update ID'])
    #     post_record = REDCAP.get_select_records(self, ls_record_id, raw = True)
    #     for record in post_record:
    #         record['time_allocation_complete'] = '2'

    #     project = redcap.Project(self.redcap_url, self.redcap_allocation_token)
    #     submit = project.import_records(post_record)

class OBJ_UPDATE():

    def __init__(self, project_plan_sheet_id=None, project_plan=None, update=None, sheet_type=None, sheet_id=None,
        intake_dept=None, redcap_form=None , portfolio=None, department=None, row_type=None, row_number=None,
        primary_row_id=None, secondary_row_id=None, close_program_search=None, tertiary_row_id=None):

        self.project_plan_sheet_id = project_plan_sheet_id
        self.project_plan = project_plan
        self.update = update
        self.sheet_type = sheet_type
        self.sheet_id = sheet_id
        self.intake_dept = intake_dept
        self.redcap_form = redcap_form
        self.portfolio = portfolio
        self.department = department

        self.close_program_search = close_program_search

        # Row details
        self.row_type = row_type

        # Project Plan attributes
        self.row_number = row_number
        self.primary_row_id = primary_row_id
        self.secondary_row_id = secondary_row_id
        self.tertiary_row_id = tertiary_row_id

    def new_row(self, obj_update):

        ss_client = SMARTSHEET.get_client(self)
        if obj_update.sheet_type == 'All Common':
            obj_update.sheet_id = obj_update.project_plan_id
            sheet = ss_client.Sheets.get_sheet(obj_update.sheet_id)
        else:
            sheet = ss_client.Sheets.get_sheet(obj_update.sheet_id) # initialize sheet

        row_list = SMARTSHEET.get_row_list(sheet)
        row = ss_client.models.Row()
        if obj_update.sheet_type != 'Project Details':
            if obj_update.row_type == 'primary':
                row.format = SmartCap_CONFIG.primary_row_color
        response = ss_client.Sheets.add_rows(obj_update.sheet_id, row)
        sheet = ss_client.Sheets.get_sheet(obj_update.sheet_id) # re initialize sheet
        n_row_list = SMARTSHEET.get_row_list(sheet) # Get new list of rows in sheet

        for row in n_row_list:
            if row not in row_list:
                new_row_id = row

        return new_row_id


    def cell(self, obj_update, record):
        '''

        '''
        ss_client = SMARTSHEET.get_client(self)
        sheet = ss_client.Sheets.get_sheet(obj_update.sheet_id) # initialize sheet
        column_map = SMARTSHEET.get_column_map(sheet)

        if obj_update.update == 'new':

            if obj_update.sheet_type == 'All Common':

                record = dict(record, **SmartCap_CONFIG.all_common_defaults)
                record = dict(record, **SmartCap_CONFIG.all_common_primary_row_defaults)
                if obj_update.row_type == 'primary':
                    record = dict(record, **SmartCap_CONFIG.all_common_primary_row_defaults)
                    record = SMARTSHEET.all_common_parent_row(record)
                elif obj_update.row_type == 'secondary':
                    record = dict(record, **SmartCap_CONFIG.all_common_primary_row_defaults)
                    record = SMARTSHEET.all_common_child_row(record)

            elif obj_update.sheet_type == 'Project Plan':

                if obj_update.row_type == 'primary':
                    record = SMARTSHEET.primary_defaults(record)
                elif obj_update.row_type == 'secondary':
                    record = SMARTSHEET.secondary_defaults(record)
                elif obj_update.row_type == 'tertiary':
                    if obj_update.row_number in list(SmartCap_CONFIG.project_plan_dates.keys()):
                        record = SMARTSHEET.tertiary_defaults(record, date_update=obj_update.row_number)
                    else:
                        record = SMARTSHEET.tertiary_defaults(record)

            elif obj_update.sheet_type == 'Project Details':

                record = SMARTSHEET.new_project_row(self, obj_update, record) # revise records

        elif self.update == 'closed':
            pass

        my_cell_obj = OBJ_CELL()

        for k, v in record.items():

            my_cell_obj.field = k
            my_cell_obj.value = v
            new_cell = ss_client.models.Cell()

            if obj_update.sheet_type == 'Project Details':
                if my_cell_obj.field in list(SmartCap_CONFIG.project_details_obj_links.keys()):
                    hyperlink = ss_client.models.Hyperlink()
                    my_cell_obj, new_cell = OBJ_CELL.object_link(self, obj_update, my_cell_obj, new_cell, hyperlink)
                new_cell.value = my_cell_obj.value

            else:
                if my_cell_obj.field in SmartCap_CONFIG.formula_cells:
                    new_cell.formula = my_cell_obj.value
                else:
                    new_cell.value = my_cell_obj.value

            if my_cell_obj.field in column_map:
                new_cell.column_id = int(column_map[my_cell_obj.field]) # set column_id attribute if the cell in in the column map
            else:
                continue

            new_row = ss_client.models.Row() #initialize new row

            if (obj_update.sheet_type == 'Project Plan') or (obj_update.sheet_type == 'All Common'): # for project plans parent rows
                if obj_update.row_type == 'primary':
                    new_row.id = obj_update.primary_row_id
                    new_row.cells.append(new_cell)
                    new_row.to_top = True
                    result = ss_client.Sheets.update_rows(sheet.id, new_row)
                else:
                    if obj_update.row_type == 'secondary':
                        current_id = obj_update.secondary_row_id
                        parent_id = obj_update.primary_row_id
                    elif obj_update.row_type == 'tertiary':
                        current_id = obj_update.tertiary_row_id
                        parent_id = obj_update.secondary_row_id

                    new_row.id = current_id
                    new_row.cells.append(new_cell)
                    result = ss_client.Sheets.update_rows(sheet.id, new_row)

                    child_row = ss_client.models.Row()
                    child_row.format = ",,,,,,,,,,,,,,,"
                    child_row.parent_id = parent_id
                    child_row.id = current_id
                    child_row.to_bottom = True
                    result = ss_client.Sheets.update_rows(sheet.id, child_row)

            elif obj_update.sheet_type == 'Project Details': # for project details rows
                new_row.id = obj_update.parent_row_id
                new_row.cells.append(new_cell)
                new_row.to_top = True
                result = ss_client.Sheets.update_rows(sheet.id, new_row)

        return result


    def cell_link(self, obj_update):

        ss_client = SMARTSHEET.get_client(self)
        projects_details_sheet = ss_client.Sheets.get_sheet(obj_update.sheet_id) # initialize destination sheet
        projects_details_column_map = SMARTSHEET.get_column_map(projects_details_sheet)

        project_plan_sheet = ss_client.Sheets.get_sheet(obj_update.project_plan_id) # initialize source sheet
        project_plan_column_map = SMARTSHEET.get_column_map(project_plan_sheet)

        for row in project_plan_sheet.rows: # GET THE TOP ROW ID OF THE PROJECT PLAN SHEET
            project_id_cell_value =  SMARTSHEET.get_cell_display_value_by_column_name(row, project_plan_column_map, 'Project ID')
            if project_id_cell_value == obj_update.record_id:
                project_plan_row_id = row.id
                break

        for row in projects_details_sheet.rows: # GET THE TOP ROW ID OF THE PROJECT PLAN SHEET
            project_id_cell_value =  SMARTSHEET.get_cell_display_value_by_column_name(row, projects_details_column_map, 'Project ID')
            if project_id_cell_value == obj_update.record_id:
                projects_details_row_id = row.id
                break

        links_column_map = {}
        for field in SmartCap_CONFIG.project_details_cell_links:
            links_column_map[field] = project_plan_column_map[field]

        for k, v in links_column_map.items():
            cell_link = ss_client.models.CellLink()
            cell_link.sheet_id = int(obj_update.project_plan_id)
            cell_link.row_id = int(project_plan_row_id)
            cell_link.column_id = int(v)

            new_cell = ss_client.models.Cell()
            new_cell.value = ss_client.models.ExplicitNull()
            new_cell.link_in_from_cell = cell_link

            if k in list(projects_details_column_map):
                new_cell.column_id = int(projects_details_column_map[k]) # add cell column id

            new_row = ss_client.models.Row()
            new_row.id = projects_details_row_id
            new_row.cells.append(new_cell)

            rows = []
            rows.append(new_row)
            result = ss_client.Sheets.update_rows(obj_update.sheet_id, rows)

        return None



class OBJ_CELL():
    field = None
    value = None

    def object_link(self, obj_update, my_cell_obj, ss_cell_obj, hyperlink):

        data_dict = copy.deepcopy(SmartCap_CONFIG.project_details_obj_links)

        if obj_update.project_plan == 'No Project Plan':
            project_plan = False
        else:
            project_plan = True

        for k, v in data_dict.items():
            if k == my_cell_obj.field:
                options = v

        for k, v in options.items():
            linked_obj = k
            linked_obj_path = v

        if (project_plan == False) and (my_cell_obj.field == 'Project Name'): # Project Name is not a link w/o a project plan
            pass

        else:
            if linked_obj == 'Dashboard':
                linked_obj_name = my_cell_obj.value
                linked_obj_path.append(linked_obj_name)
                dict_linked_obj_folders = SMARTSHEET.get_path_ids(self, linked_obj_path)
                dashboard_name = linked_obj_path[-1]
                dashboard_folder = linked_obj_path[-2]
                dashboard_folder_id = dict_linked_obj_folders[dashboard_folder]
                my_cell_obj.linked_obj_id = SMARTSHEET.get_dashboard_id(self, dashboard_name, dashboard_folder_id)
                hyperlink.sight_id = my_cell_obj.linked_obj_id

            elif linked_obj == 'Sheet':
                if obj_update.project_plan == 'All Common':
                    linked_obj_name = obj_update.portfolio + ' ' + obj_update.project_plan # file name = portfolio + All Common or Project Plan
                else:
                    linked_obj_name = my_cell_obj.value

                linked_obj_path.extend([obj_update.portfolio, linked_obj_name])
                dict_linked_obj_folders = SMARTSHEET.get_path_ids(self, linked_obj_path)
                sheet_name = linked_obj_path[-1]
                sheet_folder = linked_obj_path[-2]
                sheet_folder_id = dict_linked_obj_folders[sheet_folder]
                dict_sheet_id = SMARTSHEET.get_sheet_id(self, sheet_name, sheet_folder_id)
                if not dict_sheet_id:
                    sheet_name = obj_update.record_id
                    dict_sheet_id = SMARTSHEET.get_sheet_id(self, sheet_name, sheet_folder_id)
                for k, v in dict_sheet_id.items():
                    my_cell_obj.linked_obj_id = v

                hyperlink.sheet_id = my_cell_obj.linked_obj_id

            ss_cell_obj.hyperlink = hyperlink

        return my_cell_obj, ss_cell_obj


class OBJ_CLOSE():

    def __init__(sheet_id, project_plan, redcap_form, intake_dept, record_id):
        self.sheet_id = sheet_id
        self.project_plan = project_plan
        self.redcap_form = redcap_form
        self.intake_dept = intake_dept
        self.record_id = record_id

    def delete_project_details_rows(self, close_records):

        ls_row_ids = []
        for record in close_records:
            ls_row_ids.append(record['row_id'])

        ss_client = SMARTSHEET.get_client(self)
        dict_ids = SMARTSHEET.get_sheet_id(self, SmartCap_CONFIG.project_details_sheetname)
        sheet_id = dict_ids[SmartCap_CONFIG.project_details_sheetname]

        response = ss_client.Sheets.delete_rows(sheet_id, ls_row_ids)

        return response

    #########################################################################################################################
    # EVERYTHING BELOW HERE IS UNDER CONSTRUCTION, TESTING, OR ON HOLD FOR FURTHER SMARTSHEET UPDATES.
    #
    #########################################################################################################################



#     def timeline_report_update(self, sheet_id):


#         folder_pathway = ['Portfolio Dashboards', 'Portfolio Sheets', self.portfolio, 'Portfolio Level Reports']
#         timeline_report_names = ['Timeline Report', 'Issues Report', 'Update Report']

#         ss = SMARTSHEET.get_client(self)
#         sheet = ss.Sheets.get_sheet(sheet_id) # initialize sheet
#         column_map = SMARTSHEET.get_column_map(sheet)

#         dict_folders = SMARTSHEET.get_path_ids(self, folder_pathway)

#         for folder_name, folder_id in dict_folders.items():
#             if folder_name in folder_pathway:
#                 f_id = folder_id

#         folder = ss.Folders.get_folder(f_id)
#         for report in folder.reports:
#             if timeline_report_names[1] in report.name: # GET THE PORTFOLIO REPORT ID
#                 obj_report = report

#         report = ss.Reports.get_report(obj_report.id)
#         report_update = ss.models.Report()

#         report_update.source_sheets.append(sheet)

#         # report_update.id = report.id

#         # Currently, cannot update or create reports via the API
#         # https://stackoverflow.com/questions/43374853/smartsheet-api-create-update-reports-possible

#         # response = ss.Sheets.update_sheet(report.id, report_update)

#         # sys.exit(0)


# # class Closed_Project_Update(SMARTSHEET, REDCAP):
# #     def update_smartsheets(self, close_records):
# #         '''
# #         Given a list of record numbers, this function deletes the records from Smartsheets Projects Details sheet.
# #         '''

# #         # result = obj.obj_close.delete_project_details_rows(self, close_records)
# #         Closed_Project_Update.close_project_details(self, close_records)


#     # def upload_to_Redcap(self, close_records):

#     #     for record in close_records:
#     #         print('   updating RedCap with %s information' %(record['Project ID']))
#     #         updated_record = Closed_Project_Update.convert_fields(record)
#     #         obj_close = obj.obj_close()
#     #         pprint(record)



#     #         # update smartsheet with record data


#     def convert_fields(record):



#     def close_project_plan(self, record):

#         folder_pathway = ['Project Plans']
#         dict_folders = SMARTSHEET.get_path_ids(self, folder_pathway)
#         project_plans_f_id = dict_folders['Project Plans'] # GET PORTFOLIO PROJECT PLAN/TIMELINE FOLDER

#         if record['timeline_sheet'] == 'All Common':
#             sheet_name = self.portfolio + ' All Common Timeline'
#         else:
#             sheet_name = self.project_name + ' Project Plan'

#         ss = SMARTSHEET.get_client(self)
#         folder = ss.Folders.get_folder(project_plans_f_id)
#         self.sheet_type = record['timeline_sheet']
#         for sheet in folder.sheets:
#             if sheet_name == sheet.name: # GET THE ALL COMMON TIMELINE SHEET ID
#                 obj_sheet = sheet
#                 self.timeline_id = sheet.id


#         sheet = ss.Sheets.get_sheet(sheet_id)
#         column_map = SMARTSHEET.get_column_map(sheet)

#         for row in sheet.rows:
#             project_name = SMARTSHEET.get_cell_display_value_by_column_name(row, column_map, 'Project Name')
#             if project_name == self.project_name:
#                 if row.parent_id is None: # Get Parent Row
#                     update_row == row.id

#                 # START HERE MOVING THE CODE BELOW TO THE OBJ.PY SHEET
#             #     cell = ss.models.Cell()
#             #     cell.value = True
#             #     col_id = int(column_map['Project Closed'])
#             #     cell.column_id = col_id

#             #     update_row = ss.models.Row()
#             #     update_row.id = row.id
#             #     update_row.cells.append(cell)

#             # else:
#             #     continue

#             #         result = ss.Sheets.update_rows(sheet_id, update_row)

#             #     print(upload_dict)


#     def post_update(self, record):

#         ls_record_id = []
#         ls_record_id.append(record['Update ID'])
#         post_record = REDCAP.get_select_records(self, ls_record_id, raw = True)
#         for record in post_record:
#             record['time_allocation_complete'] = '2'

#         project = redcap.Project(self.redcap_url, self.redcap_allocation_token)
#         submit = project.import_records(post_record)
