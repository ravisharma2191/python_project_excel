import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime
from package import output_file as packagefile

#file_path = = r"C:\Users\software.support\Desktop\New folder\Output-Package.xlsx"
file_path= packagefile
#df = pd.read_excel(file_path, sheet_name="Bulk PKG Sheet")
# Read first worksheet
sheet1 = pd.read_excel(file_path, sheet_name='Package')
s_code = sheet1['PACKAGECODE']
s_NAME = sheet1['PACKAGENAME']

# Read second worksheet
#sheet2 = pd.read_excel(file_path, sheet_name='BulkPKGComponent')

sheet1['SERVICE_MASTER_ID'] = ''
sheet1['SERVICE_CODE'] = s_code
sheet1['SERVICE_NAME'] = s_NAME
sheet1['SHORT_DESC'] = ''
sheet1['SERVICE_CATEGORY'] = 111
sheet1['SERVICE_TYPE'] = 29520
sheet1['SERVICE_GENDER_CODE'] = 530
sheet1['ITEMGROUP_ID'] = 32923432
sheet1['CPT_CODE'] = ''
sheet1['CPT_DESC'] = ''
sheet1['GROUP_ID_RP'] = ''
sheet1['VISIT_TYPE_APPLICABLE'] = 109
sheet1['SCHEDULABLE'] = ''
sheet1['ALLOW_MULTIPLE_ORDER'] = ''
sheet1['REORDER_DURATION'] = 0
sheet1['REORDER_DURATION_UOM'] = ''
sheet1['EST_DURATION'] = 0
sheet1['IS_SUGICAL_SERVICE'] = 'N'
sheet1['IS_OT_REQ'] = ''
sheet1['IS_AUTHORISATION_REQ'] = 'N'
sheet1['IS_DIAGNOSTIC_SERVICE'] = 'Y'
sheet1['IS_CONSENT_REQ'] = 'Y'
sheet1['IS_ACTIVE'] = 'Y'
sheet1['IS_INDIVIDUALLYORDERABLE'] = 'N'
sheet1['IS_AUTO_PROCESS'] = 'N'
sheet1['IS_RESTRICTED'] = 'N'
sheet1['SPECIAL_INSTRUCTION'] = ''
sheet1['ORDER_TEMP_ID'] = ''
sheet1['REPORT_TEMP_ID'] = ''
sheet1['AUTO_CANCEL_DAYS'] = ''
sheet1['CREATEDBY'] = 19413005
sheet1['CREATEDDATETIME'] = datetime.today().strftime('%d-%m-%y')
sheet1['UPDATEDBY'] = ''
sheet1['UPDATEDDATETIME'] = ''
sheet1['MNEMONICS'] = ''
sheet1['AMOUNT'] = ''
sheet1['BILLINGGROUP_ID'] = 32928103
sheet1['ISPORTABLE'] = 'N'
sheet1['ISEXTERNALSERVICE'] = 'N'
sheet1['ISTARIFFCATEGORY'] = 'N'
sheet1['FINANCIALGROUP'] = 31596868
sheet1['MAX_HOURS_FOR_HOURLY_BILLING'] = 0
sheet1['MIN_MINS_FOR_BILLING'] = 0
sheet1['SERVICERENDERINGMETHOD'] = ''
sheet1['IS_IMAGINGPROCEDURE'] = 'N'
sheet1['CLAIMACTIVITYTYPE'] = ''
sheet1['AVAILABLEFORONLINE'] = 'N'
sheet1['SPECIALINSTRUCTIONS'] = 'N'
sheet1['EQUIPMENTREQUIRED'] = 'N'
sheet1['GSTPROFILEID'] = ''
sheet1['VISITTYPE_GST_APPLICABLE'] = ''
sheet1['SACNO'] = ''
sheet1['REVCOMPPERCENTAGE'] = 0
sheet1['MISGROUPID'] = ''
sheet1['ISSERVICECHARGE'] = 'Y'
sheet1['IS_ROUNDOFF'] = 'N'
sheet1['HIBVERSION'] = '0'
sheet1['CHARGABLE'] = ''
sheet1['EDITABLE'] = ''
sheet1['DISCOUNTABLE'] = ''
sheet1['NABLLOGOREQUIRED'] = 'N'
sheet1['SURGERYLEVEL'] = ''
sheet1['SHARABLE'] = 'N'
sheet1['IS_SURGICALRPT'] = 'N'
sheet1['IS_ADDITIONAL_CHARGEABLE'] = 'N'
sheet1['SHOWONWEBSITE'] = 0

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
output_df = sheet1[required_columns]

# Export to Excel
output_file = r"C:\Users\software.support\Desktop\New folder\Output-Package.xlsx"

with pd.ExcelWriter(output_file,engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:

    output_df.to_excel(writer, sheet_name='Service', index=False)
     #service_df.to_excel(writer, sheet_name="Service", index=False)

#output_df.to_excel(output_file, index=False)

#print("Excel file generated successfully.")