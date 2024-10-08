# -*- coding: utf-8 -*-
"""
Created on Sun Nov 12 19:21:00 2023

@author: Eric Moss


The purpose of this script is to prefill a Risk Exposure Table (RET) to the greatest extent possible through automation from a completed Test Case Workbook (TCW)

Updated to follow rev5 tempaltes provided by FedRAMP

Assumptions:
-TCW IS COMPLETE AND READY TO DELIVER; no more findings, all custom formating removed, all checks have been run and reviewed

-If any changes are made to the TCW, all changes must also be updated in the RET, else the RET process/script restarted from scratch.

-Indentified risks, risk statements, and reccomendations should all be written as a single paragraph in the TCW. 
---Multiple sentances can be used, but newlines are used to separate multiple findings in a single Assessment Procedure. 
---Failure to folow this will result in ValueErrors and that control family failing to be processed
---Same rule applies to PL-2 Findings

-For any "refer to" or "See" Findings - these are restricted to referancing another finding within the control or enhancement, not outside of it.
-A similar restriction exists for PL-2 Findings. At least 1 PL-2 (the first ideally) needs to have the direct statement instead of referancing antoher PL-2 or regular finding.
--"Findings/PL-2s that begin with 'Refer to' or "See" will not be processed and added to the RET" 
--"Refer to" findings should be limited to the same type of finding wherever possible (regular findings referance regular findings, PL-2s referance PL-2s)


-I have attempted to account for the most reasonable typos and variations, but the RET should still be given a onceover to ensure that all data transfered correctly.



Initial Start Date:
11/12/2023

Initial Publication Date:

    
Change Log:
2.1 - modified PL-2 functions to point control column to the implementation statement with the actual issue (An issue with the AC-2 implementation statement will refer to AC-2, not PL-2)
2.2 - added functionality to name documentation and SSP findings and iterate over them to add numbers if multiples existed.


Todo list:
    Load and Save document with GUI
    delete finding counter?
    drop Refer to and See
    tests:
        extra paragraph
        See/refer to stuff
    Error handling:
        Incorrect Data input types?
    
    
"""
#%%
#import modules and start timestamp
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook

dateTimeObj = datetime.now()
timestampStr = dateTimeObj.strftime("%d-%b-%Y (%H:%M:%S.%f)")
print('Start Timestamp : ', timestampStr)

pd.options.mode.chained_assignment = None  # default='warn'

#%% Define variables 


#Completed TCW
tcw_path = r'C:\Users\Eric\Documents\Local Documents\Appendix B - 2024 MediaLab Moderate Annual SAP-Appendix-A-FedRAMP-Moderate-Security-Test-Case-Procedures.xlsx'

#Blank RET
ret_path = r'C:\Users\Eric\Documents\Local Documents\RET_data_export.xlsx'

control_families = [
    'AC',
    'AT',
    'AU',
    'CA',
    'CM',
    'CP',
    'IA',
    'IR',
    'MA',
    'MP',
    'PE',
    'PL',
    'PS',
    'RA',
    'SA',
    'SC',
    'SI',
    'SR']



#define RET lists to append to
ret_poam_id = []
ret_controls = []
ret_name = []
ret_name_tmp = []
ret_name_tmp2 = []
ret_name_final = []
ret_description = []
ret_detection_source = []
ret_source_id = []
ret_asset_id = []
ret_detection_date = []
ret_vendor_dep = []
ret_vendor_product = []
ret_original_risk = []
ret_adjusted_risk_rating = []
ret_risk_adjustment = []
ret_false_positive = []
ret_operational_requirement = []
ret_deviation_rationale = []
ret_comments = []
ret_service_name = []
#Likelihood and impact are necessary to calculate original risk  - columns not imported into Final RET
ret_Likelihood = []
ret_Impact = []

#The following list defines the words that will be checked for to determine if a finding is a documentation finding. 
#All findings that have at least one of these words will be labeled as a documentation finding.
documentation_words_to_check = ["policy", "Policy", 'Policies', "Documentation", "documentation"]
#%% Define Functions
#finding_counter = 1

def process_findings_in_sheet(sheet):
    working_df = pd.read_excel(tcw_path, sheet_name=sheet)
    working_df = working_df.dropna(subset=['Control ID']) #drop and FedRAMP formatting rows that muck up the process
    #global finding_counter
    working_df = working_df[['Control ID', 'Assessment Procedure', 'Identified Risk', 'Likelihood Level', 'Impact Level']]    #Cut DF down to necessary columns
    working_df.dropna(subset=['Identified Risk'], inplace=True)    #removes all controls without findings
    working_df.dropna(subset=['Likelihood Level'], inplace=True)    #removes all controls that only have a PL-2 finding
    #working_df = process_multiple_findings(working_df, sheet)
    tmp_ret_controls =[]
    tmp_ret_assessments =[]
    tmp_ret_risks = []
    tmp_ret_likelihoods = []
    tmp_ret_impacts = []
    tmp_ret_poam_id = []
    tmp_ret_detection_source = []
    tmp_ret_source_id = []
    tmp_ret_asset_id = []
    tmp_ret_detection_date = []
    tmp_ret_vendor_dep = []
    tmp_ret_vendor_product = []
    tmp_ret_adjusted_risk_rating = []
    tmp_ret_false_positive = []
    tmp_ret_risk_adjustment = []
    tmp_ret_operational_requirement = []
    tmp_ret_deviation_rationale = []
    tmp_ret_comments = []
    tmp_ret_service_name = []
    tmp_ret_name = []
    ret_proccessing_list = [
        tmp_ret_controls,
        tmp_ret_assessments,
        tmp_ret_risks,
        tmp_ret_likelihoods,
        tmp_ret_impacts
        ]
    finding_replacer_list = [
        'Finding: ',
        'Finding 1: ',
        'Finding 2: ',
        'Finding 3: ',
        'Finding 4: ',
        'Finding 5: ',
        'Finding 6: ',
        'Finding 7: ',
        'Finding 8: ',
        'Finding 9: ',
        'Finding:10 ',
        'Finding 11: ',
        'Finding 12: ',
        'Finding 13: ',
        'Finding 14: ',
        'Finding 15: ',
        'Finding 16: ',
        'Finding 17: ',
        'Finding 18: ',
        'Finding 19: ',
        'Finding:20 ',
        'Finding 21: ',
        'Finding 22: ',
        'Finding 23: ',
        'Finding 24: ',
        'Finding 25: ',
        'Finding 26: ',
        'Finding 27: ',
        'Finding 28: ',
        'Finding 29: ']
    control_ids = working_df['Control ID'].tolist()
    assessments = working_df['Assessment Procedure'].tolist()
    risks = working_df['Identified Risk'].tolist()
    likelihoods = working_df['Likelihood Level'].tolist()
    impacts = working_df['Impact Level'].tolist()
    new_i_risks = []
    new_likelihoods = []
    new_impacts = []
    #split risks, liklihoods, and impacts into lists of sentances to explode
    for risk in risks:
        n_risk = risk.split('\n')
        new_i_risks.append(n_risk)
    for likelihood in likelihoods:
        n_likelihood = likelihood.split('\n')
        new_likelihoods.append(n_likelihood)
    for impact in impacts:
        n_impact = impact.split('\n')
        new_impacts.append(n_impact)
    #remove PL-2 Findings and empty lines
    iter_index=0
    for new_list in new_i_risks:
        for finding in new_list:
            if finding == "":
                continue
            elif finding.startswith("PL-2") == True:
                continue
            else:
                new_finding_stmt = finding
                for prefix in finding_replacer_list:
                    new_finding_stmt = new_finding_stmt.replace(prefix, "")
                tmp_ret_risks.append(new_finding_stmt)
                tmp_ret_controls.append(control_ids[iter_index])
                tmp_ret_assessments.append(assessments[iter_index])
                tmp_ret_poam_id.append("") #starts with first unused POA&M number
                #functionality to check and tag for doocumentation findings.
                found_words = [word for word in documentation_words_to_check if word in new_finding_stmt]
                if found_words:
                    tmp_r_name = control_ids[iter_index] + " - Documentation Deficiency"
                    tmp_ret_name.append(tmp_r_name)
                else:
                    tmp_ret_name.append("")#
                tmp_ret_detection_source.append("Assessment Test Case") #Control testing or documentation Review
                tmp_ret_source_id.append("None")
                tmp_ret_asset_id.append("N/A")#N/A, or leave blank?
                tmp_ret_detection_date.append("")#SAR date
                tmp_ret_vendor_dep.append("No")
                tmp_ret_vendor_product.append("N/A")
                tmp_ret_adjusted_risk_rating.append("N/A")
                tmp_ret_false_positive.append("No")
                tmp_ret_operational_requirement.append("No")
                tmp_ret_deviation_rationale.append("")
                tmp_ret_comments.append("")
                tmp_ret_service_name.append("")
                tmp_ret_risk_adjustment.append("No")
        iter_index+=1
    for new_list in new_likelihoods:
        for likelihood in new_list:
            if likelihood == "":
                continue
            else:
                likelihood_f = likelihood
                for prefix in finding_replacer_list:
                    likelihood_f = likelihood_f.replace(prefix, "")
                likelihood_f = likelihood_f.lstrip()
                tmp_ret_likelihoods.append(likelihood_f)
    for new_list in new_impacts:
        for impact in new_list:
            if impact == "":
                continue
            else:
                impact_f = impact
                for prefix in finding_replacer_list:
                    impact_f = impact_f.replace(prefix, "")
                impact_f = impact_f.lstrip()
                tmp_ret_impacts.append(impact_f)
    it = iter(ret_proccessing_list)
    the_len = len(next(it))
    if not all(len(l) == the_len for l in it):
        print('Unable to process ' + sheet + "Findings. Issues found while combining split data. Ensure that columns have consistent lengths and numbers of rows.")
        print("Control ID's': " + str(len(tmp_ret_controls)))
        print("Asessment Procedures: " + str(len(tmp_ret_assessments)))
        print("Risk Paragraphs: " + str(len(tmp_ret_risks)))
        print("Likelihood Paragraphs: " + str(len(tmp_ret_likelihoods)))
        print("Impact Paragraphs: " + str(len(tmp_ret_impacts)))
    else:
        for item in tmp_ret_poam_id:
            ret_poam_id.append(item)
        for item in tmp_ret_controls:
            ret_controls.append(item)
        for item in tmp_ret_name:
            ret_name.append(item)
        for item in tmp_ret_risks:
            ret_description.append(item)                
        for item in tmp_ret_detection_source:
            ret_detection_source.append(item)
        for item in tmp_ret_source_id:
            ret_source_id.append(item)
        for item in tmp_ret_asset_id:
            ret_asset_id.append(item)
        for item in tmp_ret_detection_date:
            ret_detection_date.append(item)
        for item in tmp_ret_vendor_dep:
            ret_vendor_dep.append(item)                
        for item in tmp_ret_vendor_product:
            ret_vendor_product.append(item)
        for item in tmp_ret_risk_adjustment:
            ret_risk_adjustment.append(item)
        for item in tmp_ret_false_positive:
            ret_false_positive.append(item)    
        for item in tmp_ret_operational_requirement:
            ret_operational_requirement.append(item)
        for item in tmp_ret_deviation_rationale:
            ret_deviation_rationale.append(item)                
        for item in tmp_ret_comments:
            ret_comments.append(item)
        for item in tmp_ret_service_name:
            ret_service_name.append(item)
        for item in tmp_ret_likelihoods:
            ret_Likelihood.append(item)                
        for item in tmp_ret_impacts:
            ret_Impact.append(item) 
    dateTimeObj = datetime.now()
    timestampStr = dateTimeObj.strftime("%d-%b-%Y (%H:%M:%S.%f)")    
    print(sheet + " Findings Completed: "+ timestampStr)        

def process_pl2s_in_sheet(sheet):
    working_df = pd.read_excel(tcw_path, sheet_name=sheet)
    working_df = working_df[['Control ID', 'SSP Implementation Differential?']]    #Cut DF down to necessary columns      
    working_df.dropna(subset=['SSP Implementation Differential?'], inplace=True)    #removes all controls without findings
    tmp_ret_controls =[]
    tmp_ret_risks = []
    tmp_ret_poam_id = []
    tmp_ret_detection_source = []
    tmp_ret_source_id = []
    tmp_ret_asset_id = []
    tmp_ret_detection_date = []
    tmp_ret_vendor_dep = []
    tmp_ret_vendor_product = []
    tmp_ret_risk_adjustment = []
    tmp_ret_false_positive = []
    tmp_ret_operational_requirement = []
    tmp_ret_deviation_rationale = []
    tmp_ret_comments = []
    tmp_ret_service_name = []
    tmp_ret_name = []
    tmp_ret_adjusted_risk_rating =[]
    finding_replacer_list = [
        'Finding: ',
        'Finding 1: ',
        'Finding 2: ',
        'Finding 3: ',
        'Finding 4: ',
        'Finding 5: ',
        'Finding 6: ',
        'Finding 7: ',
        'Finding 8: ',
        'Finding 9: ',
        'Finding:10 ',
        'Finding 11: ',
        'Finding 12: ',
        'Finding 13: ',
        'Finding 14: ',
        'Finding 15: ',
        'Finding 16: ',
        'Finding 17: ',
        'Finding 18: ',
        'Finding 19: ',
        'Finding:20 ',
        'Finding 21: ',
        'Finding 22: ',
        'Finding 23: ',
        'Finding 24: ',
        'Finding 25: ',
        'Finding 26: ',
        'Finding 27: ',
        'Finding 28: ',
        'Finding 29: ']
    risks = working_df['SSP Implementation Differential?'].tolist()
    risk_controls = working_df['Control ID'].tolist()
    new_i_risks = []
    #split risks, liklihoods, and impacts into lists of sentances to explode
    for risk in risks:
        n_risk = risk.split('\n')
        new_i_risks.append(n_risk)
    pl2_controls_list = []
    pl2_findings = []
    index = 0
    for pl2_list in new_i_risks:
        for pl2 in pl2_list:
            if len(pl2)==0:
                continue
            else:
                pl2_findings.append(pl2)
                pl2_controls_list.append(risk_controls[index])
        index+=1
    index = 0
    for pl2 in pl2_findings:
        pl2_stmt = pl2
        for prefix in finding_replacer_list:
            pl2_stmt = pl2_stmt.replace(prefix, "")
        tmp_ret_risks.append(pl2_stmt)
        tmp_ret_controls.append(pl2_controls_list[index])
        tmp_ret_poam_id.append("") #starts with first unused POA&M number
        tmp_ret_name.append("pl-2")#
        tmp_ret_detection_source.append("Assessment Test Case") #Control testing or documentation Review
        tmp_ret_source_id.append("None")
        tmp_ret_asset_id.append("N/A")#N/A, or leave blank?
        tmp_ret_detection_date.append("")#SAR date
        tmp_ret_vendor_dep.append("No")
        tmp_ret_vendor_product.append("N/A")
        tmp_ret_adjusted_risk_rating.append("N/A")
        tmp_ret_false_positive.append("No")
        tmp_ret_operational_requirement.append("No")
        tmp_ret_deviation_rationale.append("")
        tmp_ret_comments.append("")
        tmp_ret_service_name.append("")
        tmp_ret_risk_adjustment.append("No")
        index+=1
    for item in tmp_ret_poam_id:
        ret_poam_id.append(item)
    for item in tmp_ret_controls:
        ret_controls.append(item)
    for item in tmp_ret_name:
        ret_name.append(item)
    for item in tmp_ret_risks:
        ret_description.append(item)                
    for item in tmp_ret_detection_source:
        ret_detection_source.append(item)
    for item in tmp_ret_source_id:
        ret_source_id.append(item)
    for item in tmp_ret_asset_id:
        ret_asset_id.append(item)
    for item in tmp_ret_detection_date:
        ret_detection_date.append(item)
    for item in tmp_ret_vendor_dep:
        ret_vendor_dep.append(item)                
    for item in tmp_ret_vendor_product:
        ret_vendor_product.append(item)
    for item in tmp_ret_risk_adjustment:
        ret_risk_adjustment.append(item)
    for item in tmp_ret_false_positive:
        ret_false_positive.append(item)    
    for item in tmp_ret_operational_requirement:
        ret_operational_requirement.append(item)
    for item in tmp_ret_deviation_rationale:
        ret_deviation_rationale.append(item)                
    for item in tmp_ret_comments:
        ret_comments.append(item)
    for item in tmp_ret_service_name:
        ret_service_name.append(item)
    for item in tmp_ret_risks:
        ret_Likelihood.append('Low')                
    for item in tmp_ret_risks:
        ret_Impact.append('Low')
    dateTimeObj = datetime.now()
    timestampStr = dateTimeObj.strftime("%d-%b-%Y (%H:%M:%S.%f)")    
    print(sheet + " PL-2's Completed: "+ timestampStr) 
   
def calculate_risk(ret_Likelihood, ret_Impact):
    iter_index = 0
    global ret_original_risk
    global ret_adjusted_risk_rating
    for likelihood in ret_Likelihood:
        if likelihood == 'High' and ret_Impact[iter_index] == 'High':
            ret_original_risk.append('High')
            ret_adjusted_risk_rating.append('N/A')
        elif likelihood == 'High' and ret_Impact[iter_index] == 'Moderate':
            ret_original_risk.append('Moderate')
            ret_adjusted_risk_rating.append('N/A')
        elif likelihood == 'Moderate' and ret_Impact[iter_index] == 'High':
            ret_original_risk.append('Moderate')
            ret_adjusted_risk_rating.append('N/A')
        elif likelihood == 'Moderate' and ret_Impact[iter_index] == 'Moderate':
            ret_original_risk.append('Moderate')
            ret_adjusted_risk_rating.append('N/A')
        else:
            ret_original_risk.append('Low')
            ret_adjusted_risk_rating.append('N/A')
        iter_index += 1   

#first formula to perform basic tagging        
def define_risk_names_1(ret_name):
    index = 0
    for name in ret_name:    
        if name == "pl-2":
            string = ret_controls[index] + " - SSP Deficiency"
            ret_name_tmp.append(string)
        elif "Documentation Deficiency" in name:
            string = ret_controls[index] + " - Documentation Deficiency"
            ret_name_tmp.append(string)
        else:
            ret_name_tmp.append("")
        index+=1
        
#second formula to identify duplicate controls and add numbering to applicable controls        
def define_risk_names_2(ret_name_tmp):
    name_count = {} 
    for name in ret_name_tmp:
        if name == "":
            new_name = name
        else:
            if name in name_count:
                name_count[name] += 1
                new_name = f"{name}_{name_count[name]}"
            else:
                name_count[name] = 1
                new_name = f"{name}_{name_count[name]}"
        ret_name_tmp2.append(new_name)
        
#final formula - strips "_1" off of findings that only have a single finding associated with that topic        
def define_risk_names_3(ret_name_tmp2):
    for name in ret_name_tmp2:
        if name.endswith("_1"):
            base_name = name[:-2]  # Remove the '_1' to check for the base name
            # Check if there is a corresponding '_2' version
            if f"{base_name}_2" in ret_name_tmp2:
                ret_name_final.append(name)  # Keep the '_1' if '_2' exists
            else:
                ret_name_final.append(base_name)  # Replace '_1' with ''
        else:
            ret_name_final.append(name)
            
            
#%% Execute Findings Functions

for sheet in control_families:
    process_findings_in_sheet(sheet)
#%% Execute PL-2 Functions
for sheet in control_families:
    process_pl2s_in_sheet(sheet)
#%% Calulate Original Risk scores
calculate_risk(ret_Likelihood, ret_Impact)
#%% run naming process
define_risk_names_1(ret_name)
define_risk_names_2(ret_name_tmp)
define_risk_names_3(ret_name_tmp2)
#%% Save and Export the RET
export_df = pd.DataFrame(list(zip(
    ret_poam_id,
    ret_controls,
    ret_name_final,
    ret_description,
    ret_detection_source,
    ret_source_id,
    ret_asset_id,
    ret_detection_date,
    ret_vendor_dep,
    ret_vendor_product,
    ret_original_risk,
    ret_adjusted_risk_rating,
    ret_risk_adjustment,
    ret_false_positive,
    ret_operational_requirement,
    ret_deviation_rationale,
    ret_comments,
    ret_service_name)), 
columns = 
    ['POA&M ID', 
     'Controls',
     'Weakness Name', 
     'Weakness Description', 
     'Source of Discovery',
     'Weakness Source Identifier',
     'Affected IP Address / Hostname / Database',
     'Detection Date',
     'Vendor Dependencies',
     'Vendor Products',
     'Risk Exposure(before Mitigating Controls / Factors)',
     'Adjusted Risk Rating',
     'Risk Adjustment',
     'False Positive',
     'Operational Requirement',
     'Deviation Rationale',
     'Comments',
     'Service Name'
     ])
#strip out "refer to" and "see" findings
export_df = export_df[~export_df["Weakness Description"].str.startswith('Refer ', na=False)]
export_df = export_df[~export_df["Weakness Description"].str.startswith('See ', na=False)]



book = load_workbook(ret_path)
writer = pd.ExcelWriter(ret_path, engine='openpyxl')
writer.workbook = book
writer.worksheets = dict((ws.title, ws) for ws in book.worksheets)

export_df.to_excel(writer, sheet_name='SAR Risk Exposure Table', index = False)


writer.close()
#%%Export working_df- dev only
#working_df.to_csv(r'C:\Users\Eric\Documents\Local Documents\working.csv', index=False)
