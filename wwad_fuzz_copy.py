import pandas as pd
import datetime
import os
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
from polyfuzz import PolyFuzz


past_wwad_combined = r'..\WWAD_Combined.xlsx'
past_wwad_Testcombined = r'..\WWAD_Combined - Copy.xlsx'


employee_master = r'..employees' \
                  r' for wwad\employees.csv'

wwad_df = pd.read_excel(past_wwad_Testcombined)
employee_df = pd.read_csv(employee_master)

wwad_df[['LAST_NM', 'FIRST_NM', 'EMAIL_ADDR_TXT',  'LOC_ADDR_LINE_1_TXT', 'LOC_CITY_NM', 'LOC_STATE_NM' , 'LOC_POSTAL_CD']] = \
    wwad_df[['LAST_NM', 'FIRST_NM', 'EMAIL_ADDR_TXT', 'LOC_ADDR_LINE_1_TXT', 'LOC_CITY_NM', 'LOC_STATE_NM', 'LOC_POSTAL_CD']].astype(str)

wwad_df['FIRST_NM'] = wwad_df['FIRST_NM'].str.strip().str.lower()
wwad_df['LAST_NM'] = wwad_df['LAST_NM'].str.strip().str.lower()
wwad_df['EMAIL_ADDR_TXT'] = wwad_df['EMAIL_ADDR_TXT'].str.strip().str.lower()
wwad_df['LOC_ADDR_LINE_1_TXT'] = wwad_df['LOC_ADDR_LINE_1_TXT'].str.strip().str.lower()
wwad_df['LOC_CITY_NM'] = wwad_df['LOC_CITY_NM'].str.strip().str.lower()
wwad_df['LOC_STATE_NM'] = wwad_df['LOC_STATE_NM'].str.strip().str.lower()
wwad_df['LOC_POSTAL_CD'] = wwad_df['LOC_POSTAL_CD'].str.strip()

employee_df[['LAST_NM', 'FIRST_NM', 'EMAIL_ADDR_TXT', 'LOC_ADDR_LINE_1_TXT', 'LOC_CITY_NM', 'LOC_STATE_NM' , 'LOC_POSTAL_CD']] = \
    employee_df[['LAST_NM', 'FIRST_NM', 'EMAIL_ADDR_TXT', 'LOC_ADDR_LINE_1_TXT', 'LOC_CITY_NM', 'LOC_STATE_NM', 'LOC_POSTAL_CD']].astype(str)

employee_df['FIRST_NM'] = employee_df['FIRST_NM'].str.strip().str.lower()
employee_df['LAST_NM'] = employee_df['LAST_NM'].str.strip().str.lower()
employee_df['EMAIL_ADDR_TXT'] = employee_df['EMAIL_ADDR_TXT'].str.strip().str.lower()
employee_df['LOC_ADDR_LINE_1_TXT'] = employee_df['LOC_ADDR_LINE_1_TXT'].str.strip().str.lower()
employee_df['LOC_CITY_NM'] = employee_df['LOC_CITY_NM'].str.strip().str.lower()
employee_df['LOC_STATE_NM'] = employee_df['LOC_STATE_NM'].str.strip().str.lower()
employee_df['LOC_POSTAL_CD'] = employee_df['LOC_POSTAL_CD'].str.strip()

employee_df['FN_LN_FULLADD'] = employee_df['FIRST_NM'] + employee_df['LAST_NM'] + employee_df['LOC_ADDR_LINE_1_TXT'] + \
                           employee_df['LOC_CITY_NM'] + employee_df['LOC_STATE_NM'] + employee_df['LOC_POSTAL_CD']
employee_df['FN_LN_ADD'] = employee_df['FIRST_NM'] + employee_df['LAST_NM'] + employee_df['LOC_POSTAL_CD']

employee_df['FN_LN_FULLADD'] = employee_df['FN_LN_FULLADD'].str.replace(" ", "")
employee_df['FN_LN_ADD'] = employee_df['FN_LN_ADD'].str.replace(" ", "")

wwad_df['FN_LN_FULLADD'] = wwad_df['FIRST_NM'] + wwad_df['LAST_NM'] + wwad_df['LOC_ADDR_LINE_1_TXT'] + \
                           wwad_df['LOC_CITY_NM'] + wwad_df['LOC_STATE_NM'] + wwad_df['LOC_POSTAL_CD']

wwad_df['FN_LN_ADD'] = wwad_df['FIRST_NM'] + wwad_df['LAST_NM'] + wwad_df['LOC_POSTAL_CD']

wwad_df['FN_LN_FULLADD'] = wwad_df['FN_LN_FULLADD'].str.replace(" ", "")
wwad_df['FN_LN_ADD'] = wwad_df['FN_LN_ADD'].str.replace(" ", "")



matches_df = pd.DataFrame()
#filter for matching emails
wwad_emails_list = wwad_df['EMAIL_ADDR_TXT'].to_list()
for participant_email in wwad_emails_list:
    email_match = employee_df[employee_df['EMAIL_ADDR_TXT'].str.contains(participant_email)]
    if not email_match.empty:
        matches_df = matches_df.append(email_match)

#filter for matches and assign for variable
wwad_list = wwad_df['FN_LN_FULLADD'].to_list()
print(wwad_list)
for participant in wwad_list:
    fn_ln_add_match = employee_df[employee_df['FN_LN_FULLADD'].str.contains(participant)]
    if not fn_ln_add_match.empty:
        matches_df = matches_df.append(fn_ln_add_match)

#fuzz with first name, last name and address
fn_ln_add = employee_df['FN_LN_ADD'].to_list()
wwad_fnlnadd = wwad_df['FN_LN_ADD'].to_list()
model = PolyFuzz("TF-IDF")
model.match(fn_ln_add, wwad_fnlnadd)
model.group(link_min_similarity=0.75)
print(model.get_matches())


