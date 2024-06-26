import xml.dom.minidom as md
import logging
from datetime import datetime
import regex as re
import pandas as pd
from openpyxl import load_workbook
import os
import fire

profile_LUT = [
    "Admin - Program Support",
    "Admin - WC",
    "Admin - WC Chief",
    "Clerical Services Unit",
    "Clerical Services Unit - OAIV",
    "Clerical Services Unit Supervisor",
    "Clerical Supervisor - District Office",
    "Facilitator",
    "Facilitator - District Office",
    "Administrator",
    "Hearings Supervisor",
    "Hearings Officer",
    "Hearings Officer - District Office",
    "Hearings Reviewer",
    "Hearings Scheduler",
    "Insurance",
    "Records and Claims Branch Supervisor",
    "Records and Claims Section Supervisor",
    "Records OA4",
    "Records and Claims Section",
    "District Office Manager",
    "Office Assistant - District Office",
    "Admin DCD IT Support",
    "Admin - TDI/PHC",
    "Plans Acceptance Branch",
    "Audit Supervisor",
    "Audit Section",
    "Enforcement Supervisor",
    "Investigations",
    "Investigations - District Office",
    "Vocational Rehabilitation",
    "LIRAB",
    "Research and Statistics",
    "Fiscal",
    "AG/SCF",
    "EDPSO",
]

object_LUT_P1 = [
    "Account",
    "DCD_Account_Contact__c",
    "Employer__c",
    "DCD_Award_Worksheet__c",
    "DCD_Case_Processing__c",
    "DCD_Case_Settlement__c",
    "Contact",
    "DCD_Case__c",
    "DCD_Decision__c",
    "Dependent__c",
    "DCD_Employment__c",
    "DCD_Hearing__c",
    "DCD_Hearing_Purpose__c",
    "DCD_Hearing_Type__c",
    "Injured_Body_Part__c",
    "DCD_Order__c",
    "DCD_Related_Contact__c",
    "DCD_Period_of_Disability__c",
    "DCD_Required_Document__c",
    "DCD_Vocational_Rehabilitation_Process__c",
    "Dependent_relationship__c",
    "DCD_WC_1__c",
    "DCD_WC_2__c",
    "DCD_WC_3A__c",
    "DCD_WC_5__c",
    "DCD_WC_5A__c",
    "DCD_Address_History__c",
    "Status_History__c",
    "Request_for_Reconsideration__c",
    "DCD_Address_History_Archive__c",
    "Calculation_History__c",
    "DCD_Room__c",
    "DCD_Settlement_Agreement__c",
    "DCD_VR_Provider__c"
]

object_LUT = [
    "DCD_Case_Award__c",
    "DCD_Case_Vendor__c",
    "DCD_Complaint__c",
    "DCD_Coverage__c",
    "DCD_Request__c",
    "DCD_Disbursement__c",
    "DCD_Employer_Audit__c",
    "DCD_GLAccount__c",
    "DCD_HC_15s__c",
    "DCD_HC_4s__c",
    "DCD_HC_61s__c",
    "DCD_Health_Care_Plan__c",
    "DCD_Hearing_HRS_HAR__c",
    "DCD_Monthly_Premium__c",
    "DCD_Net_Profit_or_Loss_After_Taxes__c",
    "DCD_Prehearings__c",
    "DCD_Receipt__c",
    "DCD_Routed_Information__c",
    "DCD_Schedule_of_Disbursement__c",
    "DCD_TDI_Case__c",
    "DCD_GL_Mapping__c",
    "DCD_TDI_Plan_Type__c",
    "DCD_Coverage_Equivalency__c",
    "DCD_TDI_Special_Fund_Case__c",
    "DCD_TDI_15s__c",
    "DCD_TDI_30s__c",
    "DCD_TDI_46s__c",
    "DCD_TDI_62s__c",
    "DCD_Tracking_Log__c",
    "DCD_WC_Insurance_Policy__c",
    "DCD_WC_3__c",
    "DCD_WC_3_Benefit_Payment__c",
    "DCD_WC_3_Benefit_Payment_Summary__c",
    "DCD_Prepare_Expenditure_Voucher__c",
    "DCD_HRS_Section__c",
    "DCD_Disability_Benefits_Commencing__c",
    "DCD_Disability_Benefits_Payable__c",
    "DCD_Surety_Bond_SecurityDeposit__c",
    "DCD_Related_Contact_History__c",
    "DCD_TDI_21s__c",
    "R_S_Error_Report__c",
]

disabled_flow = [
    "Account_Trigger_Handler",
    "Autolaunched_Set_Contact_Masked_Fields",
    "Case_Case_Number_Numeric_Update",
    "Create_Party_Name_And_Address_Record",
    "Create_SCF_in_Party",
    "DCD_Complaints_Mapping_Field",
    "DCD_Coverage_Check_and_Update_WC_Status",
    "List_of_Ended_Cases_Report",
    "Non_compliant_Employer_Insurance_Information_Report",
    "Public_Portal_File_Upload",
    "TDI_SFC_File_Upload",
    "Delete_Hearing_Type",
    "Delete_Hearing_Type_Platform",
    "Update_Related_Account_on_update_Coverage",
    "WC_3_Trigger_Handler_Schedule",
    "WC_Carrier_Handler",
    "DCD_TDI_62_Update_the_status_of_form_when_Action_taken_is_change",
    "WC_Coverage_Check_Prior_WC_Status",
    "WC_Coverage_Check_Effective_Date_change_Status",
    "Account_Master_PB",
    "Award_Employment_Master",
    "Award_Worksheet_Master",
    "Case_Settlement_Master2",
    "Contact_Master",
    "DCD_Special_Fund_Case_Master",
    "TDI_Cases",
    "TDI_Special_Fund_Cases",
    "encrypted_contact"
]

#update the file name to be the corrected one
org_permission_file = "org_permission.csv"
#update the file nem to be the corrected one
#get the configuration file
root = md.parse("config.xml")
permission_config = root.getElementsByTagName("permission_matrix")[0]
matrix_file = permission_config.getElementsByTagName("version")[0].firstChild.nodeValue

profile_permission_p1 = []
profile_permission_p2 = []


#helper functions
def get_node_value(tag_name, dom_list):
    return dom_list.getElementsByTagName(tag_name)[0].childNodes[0].nodeValue

def print_dict(dictionary):
    for dict in dictionary:
        print(dict)


formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')

def setup_logger(name, log_file, level):
    #to setup individual logger
    handler = logging.FileHandler(log_file)
    handler.setFormatter(formatter)

    logger = logging.getLogger(name)
    logger.setLevel(level)

    #adding file handler to output to file
    logger.addHandler(handler)     

    stream = logging.StreamHandler()
    stream.setLevel(level)
    streamformat = logging.Formatter("%(asctime)s:%(levelname)s:%(message)s")
    stream.setFormatter(streamformat)

    #adding stream handler to output to console
    logger.addHandler(stream)

    return logger



#parsing functions to parse the data from the profile permission get from Org
def parse_org_permission(profile, file, permission_list, LUT):
    #read csv permission matrix into df
    raw_data = pd.read_csv(file)
    reduced_df = raw_data.drop(columns=["Parent", "_", "Parent.Profile"])
    

    #extract data for a certain profile
    reduced_df = reduced_df[reduced_df['Parent.Profile.Name'] == profile]
    reduced_df = reduced_df.set_index('SobjectType')
    for obj_name in LUT:
        try:
        #extract a data based on index and profile name
            object_dict = {}
            object_dict['name'] = obj_name
            object_dict['C'] = reduced_df.loc[obj_name]['PermissionsCreate']
            object_dict['R'] = reduced_df.loc[obj_name]['PermissionsRead']
            object_dict['U'] = reduced_df.loc[obj_name]['PermissionsEdit']
            object_dict['D'] = reduced_df.loc[obj_name]['PermissionsDelete']
            object_dict['MA'] = reduced_df.loc[obj_name]['PermissionsModifyAllRecords']
            object_dict['VA'] = reduced_df.loc[obj_name]['PermissionsViewAllRecords']
            #append the dictionary to the list
            permission_list.append(object_dict)
        except KeyError:
            pass
        except:
            print("Critical error")

#parsing functions to parse the data from permission matrix
def parse_matrix_csv(profile, matrix_file, permission_list, LUT, sheetname):
    #read Excel permission matrix into df and set API Name as indenx column
    raw_data = pd.read_excel(io=matrix_file,sheet_name=sheetname)
    reduced_df = raw_data.drop(columns=["Object", "Description", "Permissions Legend"])
    reduced_df = reduced_df.set_index("API Name")

    for obj_name in LUT:
        try:
            extracted_cell = reduced_df.loc[obj_name][profile]
            #extract a data based on index and profile name
            if pd.isnull(reduced_df.loc[obj_name, profile]) == False and (extracted_cell != "None" and extracted_cell != "x"):
                permission = re.split('-', reduced_df.loc[obj_name][profile])
                object_dict = {}
                object_dict['name'] = obj_name
                object_dict['C'] = False
                object_dict['R'] = False
                object_dict['U'] = False
                object_dict['D'] = False
                object_dict['MA'] = False
                object_dict['VA'] = False
                
                for val in permission:
                    temp_dict = {val:True}
                    object_dict.update(temp_dict)

                #append the dictionary to the list
                permission_list.append(object_dict)
        except KeyError:
            pass
        except:
            print("Critical Error")

def permission_compare(permission_list, logger):
    for profile in permission_list:
        print("Compare profile: ", profile[0])
        org_list = profile[1]
        matrix_list = profile[2]
        for (org, matrix) in zip(org_list, matrix_list):
            if org['C'] != matrix['C']:
                logger.info('[{}] [{}] Create in Org {}. Create in matrix {}'.format(profile[0], org['name'], org['C'], matrix['C']))
            
            if org['R'] != matrix['R']:
                logger.info('[{}] [{}] Read in Org {}. Read in matrix {}'.format(profile[0], org['name'], org['R'], matrix['R']))

            if org['U'] != matrix['U']:
                logger.info('[{}] [{}] Update in Org {}. Update in matrix {}'.format(profile[0], org['name'], org['U'], matrix['U']))

            if org['D'] != matrix['D']:
                logger.info('[{}] [{}] Delete in Org {}. Delete in matrix {}'.format(profile[0], org['name'], org['D'], matrix['D']))

            if org['MA'] != matrix['MA']:
                logger.info('[{}] [{}] Modify All in Org {}. Modify All in matrix {}'.format(profile[0], org['name'], org['MA'], matrix['MA']))

            if org['VA'] != matrix['VA']:
                logger.info('[{}] [{}] View All in Org {}. View All in matrix {}'.format(profile[0], org['name'], org['VA'], matrix['VA']))



#main loop of the program
def main_permission():

    #create log file with current date as file name
    current_datetime = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    str_current_datetime = str(current_datetime)
    p1_file_name = str_current_datetime+ "_p1" + ".txt"
    p2_file_name = str_current_datetime+ "_p2" + ".txt"
    
    #setup logging module
    p1_logger = setup_logger('P1 Objects Logger', p1_file_name, logging.INFO)
    p2_logger = setup_logger('P2 Objects Logger', p2_file_name, logging.INFO)

    #loop through all profiles
    for profile in profile_LUT:
        org_permission_list_p1 = []
        org_permission_list = []
        matrix_permission_list_p1 = []
        matrix_permission_list = []

        #extract permission for P1 objects from org
        parse_org_permission(profile, 
                             org_permission_file, 
                             org_permission_list_p1, 
                             object_LUT_P1)
        #extract permission for P2 objects from org
        parse_org_permission(profile, 
                             org_permission_file, 
                             org_permission_list, 
                             object_LUT)

        #extract permission for P1 objects from matrix
        parse_matrix_csv(profile,
                         matrix_file,
                         matrix_permission_list_p1, 
                         object_LUT_P1,
                         sheetname="PermissionMatrix_P1")
        #extract permission for P2 objects from matrix
        parse_matrix_csv(profile,
                         matrix_file,
                         matrix_permission_list, 
                         object_LUT,
                         sheetname="PermissonMatrix_P2")


        profile_permission_p1.append([profile, 
                                      org_permission_list_p1, 
                                      matrix_permission_list_p1])
        
        profile_permission_p2.append([profile,
                                      org_permission_list,
                                      matrix_permission_list])

    #compare all profile in the permission list:
    p1_logger.info("P1 RESULT:")
    permission_compare(profile_permission_p1, p1_logger)
    p2_logger.info("P2 RESULT:")
    permission_compare(profile_permission_p2, p2_logger)

def main_flows():
    #create log file with current date as file name
    current_datetime = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    str_current_datetime = str(current_datetime)
    flows_file_name = str_current_datetime+ "_flows" + ".txt"
    
    #setup logging module
    flows_logger = setup_logger('Flows Logger', flows_file_name, logging.INFO)
    flows_logger.debug("Check flow activation")

    flow_csv = "DCD_Flow.csv"
    raw_flow_data = pd.read_csv(flow_csv)
    reduced_flow_df = raw_flow_data.drop(columns=["Id", "Label", "TriggerType", "_"])

    for idx in range (len(reduced_flow_df)):
        flow_name = reduced_flow_df.iloc[idx, 0]
        is_active = reduced_flow_df.iloc[idx, 1]
        is_out_of_date = reduced_flow_df.iloc[idx, 2]
        if flow_name in disabled_flow and is_active == True:
            flows_logger.info('[{}] flow should be disabled'.format(flow_name))
        elif flow_name not in disabled_flow:
            if is_active == False:
                flows_logger.info('[{}] flow should be activated'.format(flow_name))
            else:
                if(is_out_of_date == True):
                    flows_logger.info('[{}] flows is not running the latest version'.format(flow_name))

if __name__ == "__main__":

    #delete old log
    dir_name = os.getcwd()
    folder = os.listdir(dir_name)
   
    for file in folder:
        if file.endswith(".txt"):
            os.remove(os.path.join(dir_name, file))

    fire.Fire({
        'permission': main_permission,
        'flows': main_flows,
    })
    

