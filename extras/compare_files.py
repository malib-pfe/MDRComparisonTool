import pandas as pd
import numpy as np

# Not used in final app. Good place to test out functionality separately.
def isolate_mdr(mdr_df, rcc_df):
    # Function to only get relevant forms for specific RCC study from MDR since MDR has ALL forms.
    
    # Isolate the items to compare from MDR.
    mdr_df = mdr_df[mdr_df["folder"] == "Volume 3"] # Only items from Volume 3
    mdr_df = mdr_df[["folder","f_ver","s_ver","mdes_form_name", "mde_name", "item_refname", "mde_design_instruction", "mandatory_to_be_collected"]]

    # Get the most up to date version of each item. Sort by descending to search longer names first. Removes any duplicates
    mdr_df = mdr_df.sort_values('s_ver', ascending=False).drop_duplicates(["mdes_form_name","item_refname"])
    mdr_df = mdr_df.sort_values("mdes_form_name", ascending = False)

    # Get only the relevant forms and items.
    mdr_df['Form Present in RCC Build'] = np.nan
    mdr_df['Form Present in RCC Build'] = mdr_df["mdes_form_name"].apply(lambda form: form if rcc_df['RefName Path'].str.startswith(form).any() else False) # Lambda: return the form if the RCC Form starts with a mandatory form name.
    mdr_df = mdr_df[mdr_df['Form Present in RCC Build'] != False] # Removes any forms that were not found in RCC export.
    return mdr_df

def map_rcc_formnames(relevant_forms, rcc_df):
    df = rcc_df["RefName Path"].apply((lambda form: searchform if form.startswith(searchform) else np.nan for searchform in relevant_forms), axis = 1) # Lambda: return the matching searchform from the list of relevant MDR forms if the RCC Form starts with the searchform
    df = df.apply(lambda row: pd.Series(row.dropna().values), axis=1) # Moves all results to the left and removes na's.
    df = df.apply(lambda row: max(row.astype(str), key=len), axis = 1) # Get the longest string match. Some forms can match twice, but a longer match is more exact. (ex: AE001 vs. AE001_1)
    df = pd.DataFrame(df, columns=['mdes_form_name'])
    return df

def map_rcc_itemnames(relevant_vars, rcc_df):
    df_vars = rcc_df["Variable Name"].apply((lambda item: searchitem if item.startswith(searchitem) else np.nan for searchitem in relevant_vars), axis = 1) # Lambda: return the form if the RCC Form starts with a mandatory form name.
    df_vars = df_vars.apply(lambda row: pd.Series(row.dropna().values), axis=1) # Gets rid of all na values, and moves any values into one column.
    df_vars = df_vars.apply(lambda row: max(row.astype(str), key=len), axis = 1)
    df_vars = pd.DataFrame(df_vars, columns=["item_refname"])
    return df_vars

def create_fake_study(rcc_df, mdr_df):
    # Creates list of mandatory elements using names from RCC.
    relevant_rcc_form_names = sorted(set(rcc_df["RefName Path"])) # List of RCC forms
    rcc_names_df = pd.DataFrame(relevant_rcc_form_names, columns=["RefName Path"]) # New dataframe with just a list of the RCC Forms.
    mandatory_df = rcc_names_df.merge(rcc_df, on = "RefName Path")[["RefName Path","mdes_form_name"]] # Merge rcc_df onto names df to have mapping of RCC form names to MDR form names.
    mandatory_df = mandatory_df.drop_duplicates(["RefName Path","mdes_form_name"]) # Gives a list that maps RCC Forms to MDR Form Names
    mandatory_df = mandatory_df.merge(mdr_df,how = 'left', on=['mdes_form_name']) # Merge MDR onto the mandatory df where the MDR names match
    mandatory_df = mandatory_df[mandatory_df["mandatory_to_be_collected"] == True] # Keep only mandatory fields for this fake study
    mandatory_df = mandatory_df[['RefName Path', 'item_refname',"mde_design_instruction"]] # Keep only relevant columns.
    # 'RefName Path', 'item_refname',"mde_design_instruction"
    # AE001_1          AESCAT            instruction 

    return mandatory_df

def return_missing_fields(rcc_df, mandatory_df):
    final_df = mandatory_df.merge(rcc_df,how = 'left', on=["RefName Path", "item_refname"]) # Merge export where RCC Form names and MDR item names match.
    final_df = final_df[final_df['Variable Name'].isnull()] # Variable name is RCC item name, if it's null, then a mandatory field is missing.
    final_df.insert(loc=2, column='Description', value=[f"{item} is marked as 'Mandatory' in the MDR Repository; however, it is not being collected in {form}." for form, item in zip(final_df['RefName Path'], final_df['item_refname'])]) # Inserts description of error.
    #final_df['Description'] = [f"{item} is marked as 'Mandatory' in the MDR Repository; however, it is not being collected in {form}." for form, item in zip(final_df['RefName Path'], final_df['item_refname'])] 
    final_df = final_df.rename(columns={'RefName Path': 'Form Name', 'item_refname': 'Item', "mde_design_instruction":'Context'})
    final_df = final_df.drop(columns=["Variable Name", 'mdes_form_name']) # Drops empty columns.
    return final_df


def compare_files(rcc:str,mdr:str) -> pd.DataFrame:
    # Full print option used for development.
    pd.set_option("display.max_rows", None, "display.max_columns", None)
    
    # Make user-provided files into dataframes.
    mdr_df = pd.read_excel(mdr, sheet_name="Data",engine="openpyxl")
    rcc_df = pd.read_excel(rcc, sheet_name="Item",engine="openpyxl")

    # Get relevant dataframe from export metadata.
    rcc_df = rcc_df[["RefName Path","Variable Name"]]
    rcc_df['RefName Path'] = rcc_df['RefName Path'].str.split(' >> ').str[0] # Isolates Form from rest of path in metadata.

    mdr_df = isolate_mdr(mdr_df, rcc_df)

    # Get list of forms and variables from MDR.
    relevant_forms = sorted(set(mdr_df["mdes_form_name"]), reverse = True) # Set gives a list of unique items (removes dups), sorted Z-A to get longer strings first
    relevant_vars = sorted(set(mdr_df["item_refname"].astype("str")), reverse=True)
    
    # Create a df that maps RCC export Form Names to their MDR form Names
    mapped_rcc_formnames = map_rcc_formnames(relevant_forms, rcc_df)
    rcc_df = rcc_df.merge(mapped_rcc_formnames['mdes_form_name'], how='outer', left_index=True, right_index=True) # Merge MDR Form names onto RCC metadata export

    # Create a df that maps RCC export Item Names to their MDR item Names
    mapped_rcc_itemnames = map_rcc_itemnames(relevant_vars, rcc_df)
    rcc_df = rcc_df.merge(mapped_rcc_itemnames['item_refname'], how='outer', left_index=True, right_index=True).dropna(axis=0) # Merge MDR item names onto RCC metadata export. Drop any forms that don't exist in MDR.

    # Create a fake study of mandatory items using a list of RCC forms from the export. This captures all duplicates of the same forms (DM001 and DM001_1: means two sets of DM001 forms, not just presence of DM001 forms)
    mandatory_df = create_fake_study(rcc_df, mdr_df)
    
    # Compares what is found in the required fields df just made to actual study.
    missing_df = return_missing_fields(rcc_df, mandatory_df)
    return missing_df

if __name__ == "__main__":
    compare_files("test.xlsx","MDR_Data_Collection_MDRView_Apr_14_2025.xlsx")