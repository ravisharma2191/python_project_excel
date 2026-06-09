import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime

#from db_connection import get_connection
#from package import output_file as packagefile
import numpy as np




file_path = r"C:\Users\software.support\Desktop\New folder\Output-Package.xlsx"
sheet5 = pd.read_excel(file_path, sheet_name='Working')
sheet6 = pd.read_excel(file_path, sheet_name='Working')
Service_Code = sheet5['Service_Code']
Service_name = sheet5['Service Name']
Itemgroupid = sheet5['Groupid']
DayCare= sheet5['Day Care']
EcoCeling = sheet5['Eco Celing']
TwinCeiling = sheet5['Twin Ceiling']
SingleCeling = sheet5['Single Celing']
Deluxeceiling =sheet5['Deluxe ceiling']

sheet5['SERVICE_MASTER_ID'] = ''
sheet5['SERVICE_CODE'] = Service_Code
sheet5['SERVICE_NAME'] = Service_name
sheet5['SHORT_DESC'] = ''
sheet5['SERVICE_CATEGORY'] = 111
sheet5['SERVICE_TYPE'] = np.select(
    [ Itemgroupid == 1399,
        Itemgroupid == 37312171,
        Itemgroupid == 37312236,Itemgroupid == 37312154,],
    [ 29379, 103505122,1012,1012],default=29522)
sheet5['SERVICE_GENDER_CODE'] = 580
sheet5['ITEMGROUP_ID'] = Itemgroupid
sheet5['CPT_CODE'] = ''
sheet5['CPT_DESC'] = ''
sheet5['GROUP_ID_RP'] = ''
sheet5['VISIT_TYPE_APPLICABLE'] = 581
sheet5['SCHEDULABLE'] = 'N'
sheet5['ALLOW_MULTIPLE_ORDER'] = 'Y'
sheet5['REORDER_DURATION'] = 0
sheet5['REORDER_DURATION_UOM'] = ''
sheet5['EST_DURATION'] = 0
sheet5['IS_SUGICAL_SERVICE'] = np.select(
    [ Itemgroupid == 1399,
        Itemgroupid== 37312171,],
    [ 'N', 'N',],default='Y')
sheet5['IS_OT_REQ'] = ''
sheet5['IS_AUTHORISATION_REQ'] = ''
sheet5['IS_DIAGNOSTIC_SERVICE'] = ''
sheet5['IS_CONSENT_REQ'] = 'N'
sheet5['IS_ACTIVE'] = 'Y'
sheet5['IS_INDIVIDUALLYORDERABLE'] = 'Y'
sheet5['IS_AUTO_PROCESS'] = 'Y'
sheet5['IS_RESTRICTED'] = 'N'
sheet5['SPECIAL_INSTRUCTION'] = ''
sheet5['ORDER_TEMP_ID'] = ''
sheet5['REPORT_TEMP_ID'] = ''
sheet5['AUTO_CANCEL_DAYS'] = '0'
sheet5['CREATEDBY'] = 19413005
sheet5['CREATEDDATETIME'] = datetime.today().strftime('%d-%m-%y')
sheet5['UPDATEDBY'] = ''
sheet5['UPDATEDDATETIME'] = ''
sheet5['MNEMONICS'] = ''
sheet5['AMOUNT'] = ''
sheet5['BILLINGGROUP_ID'] = 1742384
sheet5['ISPORTABLE'] = 'Y'
sheet5['ISEXTERNALSERVICE'] = 'N'
sheet5['ISTARIFFCATEGORY'] = 'N'
sheet5['FINANCIALGROUP'] = 31596868
sheet5['MAX_HOURS_FOR_HOURLY_BILLING'] = 0
sheet5['MIN_MINS_FOR_BILLING'] = 0
sheet5['SERVICERENDERINGMETHOD'] = ''
sheet5['IS_IMAGINGPROCEDURE'] = 'N'
sheet5['CLAIMACTIVITYTYPE'] = ''
sheet5['AVAILABLEFORONLINE'] = 'N'
sheet5['SPECIALINSTRUCTIONS'] = 'N'
sheet5['EQUIPMENTREQUIRED'] = 'N'
sheet5['GSTPROFILEID'] = 7728669
sheet5['VISITTYPE_GST_APPLICABLE'] = 581
sheet5['SACNO'] = ''
sheet5['REVCOMPPERCENTAGE'] = 0
sheet5['MISGROUPID'] = ''
sheet5['ISSERVICECHARGE'] = 'Y'
sheet5['IS_ROUNDOFF'] = 'N'
sheet5['HIBVERSION'] = '0'
sheet5['CHARGABLE'] = 'Y'
sheet5['EDITABLE'] = ''
sheet5['DISCOUNTABLE'] = ''
sheet5['NABLLOGOREQUIRED'] = 'N'
sheet5['SURGERYLEVEL'] = ''
sheet5['SHARABLE'] = np.select(
    [  Itemgroupid == 37312171,],
    [ 'N'],default='Y')
sheet5['IS_SURGICALRPT'] = 'N'
sheet5['IS_ADDITIONAL_CHARGEABLE'] = 'N'
sheet5['SHOWONWEBSITE'] = 0

# Required column order
required_columns = [
    'SERVICE_MASTER_ID',
    'SERVICE_CODE',
    'SERVICE_NAME',
    'SHORT_DESC',
    'SERVICE_CATEGORY',
    'SERVICE_TYPE',
    'SERVICE_GENDER_CODE',
    'ITEMGROUP_ID',
    'CPT_CODE',
    'CPT_DESC',
    'GROUP_ID_RP',
    'VISIT_TYPE_APPLICABLE',
    'SCHEDULABLE',
    'ALLOW_MULTIPLE_ORDER',
    'REORDER_DURATION',
    'REORDER_DURATION_UOM',
    'EST_DURATION',
    'IS_SUGICAL_SERVICE',
    'IS_OT_REQ',
    'IS_AUTHORISATION_REQ',
    'IS_DIAGNOSTIC_SERVICE',
    'IS_CONSENT_REQ',
    'IS_ACTIVE',
    'IS_INDIVIDUALLYORDERABLE',
    'IS_AUTO_PROCESS',
    'IS_RESTRICTED',
    'SPECIAL_INSTRUCTION',
    'ORDER_TEMP_ID',
    'REPORT_TEMP_ID',
    'AUTO_CANCEL_DAYS',
    'CREATEDBY',
    'CREATEDDATETIME',
    'UPDATEDBY',
    'UPDATEDDATETIME',
    'MNEMONICS',
    'AMOUNT',
    'BILLINGGROUP_ID',
    'ISPORTABLE',
    'ISEXTERNALSERVICE',
    'ISTARIFFCATEGORY',
    'FINANCIALGROUP',
    'MAX_HOURS_FOR_HOURLY_BILLING',
    'MIN_MINS_FOR_BILLING',
    'SERVICERENDERINGMETHOD',
    'IS_IMAGINGPROCEDURE',
    'CLAIMACTIVITYTYPE',
    'AVAILABLEFORONLINE',
    'SPECIALINSTRUCTIONS',
    'EQUIPMENTREQUIRED',
    'GSTPROFILEID',
    'VISITTYPE_GST_APPLICABLE',
    'SACNO',
    'REVCOMPPERCENTAGE',
    'MISGROUPID',
    'ISSERVICECHARGE',
    'IS_ROUNDOFF',
    'HIBVERSION',
    'CHARGABLE',
    'EDITABLE',
    'DISCOUNTABLE',
    'NABLLOGOREQUIRED',
    'SURGERYLEVEL',
    'SHARABLE',
    'IS_SURGICALRPT',
    'IS_ADDITIONAL_CHARGEABLE',
    'SHOWONWEBSITE'
]

# Final output
output_df = sheet5[required_columns]

# Export to Excel
output_file = r"C:\Users\software.support\Desktop\New folder\Output-Package.xlsx"

with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:

    output_df.to_excel(writer, sheet_name='Surgery', index=False)
     #service_df.to_excel(writer, sheet_name="Service", index=False)

#output_df.to_excel(output_file, index=False)

sheet6['SERVICE_CODE'] = Service_Code
sheet6['SERVICE_NAME'] = Service_name
sheet5['OP'] =EcoCeling
sheet5['Day Care'] =DayCare
sheet5['Eco Celing'] =EcoCeling
sheet5['Twin Ceiling']=TwinCeiling
sheet5['Single Celing']=SingleCeling
sheet5['Deluxe ceiling']=Deluxeceiling 

required_columns1 = [
    'SERVICE_CODE',
    'SERVICE_NAME',
    'OP',
    'Day Care',
    'Eco Celing',
    'Twin Ceiling',
    'Single Celing',
    'Deluxe ceiling'

]
output_df1 = sheet5[required_columns1]

# Export to Excel
output_file = r"C:\Users\software.support\Desktop\New folder\Output-Package.xlsx"

with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:

    output_df1.to_excel(writer, sheet_name='servicerate', index=False)
     #service_df.to_excel(writer, sheet_name="Service", index=False)

#output_df.to_excel(output_file, index=False)