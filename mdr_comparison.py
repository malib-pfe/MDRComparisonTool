import pandas as pd
import numpy as np
import openpyxl as op
from datetime import datetime
import os
from warnings import filterwarnings
from nicegui import app, ui, run, html, native

# Filter out warnings due to weird Workbook naming.
filterwarnings("ignore", message="Workbook contains no default style", category=UserWarning)

def isolate_mdr(mdr_df, rcc_df):
    # Function to only get relevant forms for specific RCC study from MDR since MDR has ALL forms.
    
    # Isolate the items to compare from MDR.
    mdr_df = mdr_df[mdr_df["latest"] == True] # Only items that are the latest
    mdr_df = mdr_df[mdr_df['f_ver'].str.contains('Volume 3')]
    mdr_df = mdr_df[(mdr_df['library'] == 'Core') | (mdr_df['library'] == 'Efficacy')]
    mdr_df = mdr_df[["f_ver","mdes_form_name", "mde_name", "item_refname", "crf_collection_guidance", "mandatory_to_be_collected", "mde_is_cond_reqd"]]

    # Get the most up to date version of each item. Sort by descending to search longer names first. Removes any duplicates
    #mdr_df = mdr_df.sort_values('s_ver', ascending=False).drop_duplicates(["mdes_form_name","item_refname"])
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
    mandatory_df = mandatory_df[['RefName Path', 'item_refname',"crf_collection_guidance", "mde_is_cond_reqd"]] # Keep only relevant columns.
    # 'RefName Path', 'item_refname',"crf_collection_guidance"
    # AE001            AESCAT            instruction
    # AE001_1          AESCAT            instruction 
    # Example row where first column is the RCC create event, the second column is the mandatory fields associated, and the last column is any context.
    return mandatory_df

def return_missing_fields(rcc_df, mandatory_df):
    final_df = mandatory_df.merge(rcc_df,how = 'left', on=["RefName Path", "item_refname"]) # Merge export where RCC Form names and MDR item names match.
    final_df = final_df[final_df['Variable Name'].isnull()] # Variable name is RCC item name, if it's null, then a mandatory field is missing.
    final_df['Type'] = final_df.mde_is_cond_reqd.apply(lambda x: "Optionally Required" if x == True else "Mandatory")
    final_df.insert(loc=3, column='Description', value=[f"{item} is marked as {mand_string} in the MDR Repository; however, it is not being collected in {form}." for form, item, mand_string in zip(final_df['RefName Path'], final_df['item_refname'], final_df['Type'])]) # Inserts description of error.
    #final_df['Description'] = [f"{item} is marked as 'Mandatory' in the MDR Repository; however, it is not being collected in {form}." for form, item in zip(final_df['RefName Path'], final_df['item_refname'])] 
    final_df = final_df.rename(columns={'RefName Path': 'Form Name', 'item_refname': 'Item', "crf_collection_guidance":'Context'})
    final_df = final_df.drop(columns=["Variable Name", 'mdes_form_name', 'mde_is_cond_reqd']) # Drops empty columns.
    final_df = final_df[['Form Name', 'Item', 'Type', 'Description','Context']]
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

async def choose_rcc_file():
    file = await app.native.main_window.create_file_dialog(allow_multiple=False, file_types= ('Excel Files (*.xlsx)',))
    if file is not None:
        if check_file_for_sheet('Item', file[0]):
            n3 = ui.notification("Checking RCC Metadata Export...", type='ongoing', timeout=None, spinner=True)
            is_filtered = await run.cpu_bound(check_file_for_filter, 'Item', file[0])
            if not is_filtered:
                n3.message = "MDR file selected."
                n3.type = "positive"
                n3.timeout = 3
                n3.spinner = False
                rcc_filepath.set_text(file[0])
            else:
                n3.message = "Filter exists in Item sheet. Please check RCC Metadata Export."
                n3.type = "negative"
                n3.timeout = 3
                n3.spinner = False
        else:
            ui.notify("'Item' sheet not found. Please check file.", type='negative')            
    else:
        ui.notify('No file selected.')

async def choose_mdr_file():
    file = await app.native.main_window.create_file_dialog(allow_multiple=False, file_types= ('Excel Files (*.xlsx)',))
    if file is not None:
        if datetime.now().strftime("_%b_%d_%Y") in file[0]:
            if check_file_for_sheet('Data', file[0]):
                n2 = ui.notification("Checking MDR file...", type='ongoing', timeout=None, spinner=True)
                is_filtered = await run.cpu_bound(check_file_for_filter, 'Data', file[0])
                if not is_filtered:
                    is_pmdr = await run.cpu_bound(check_file_for_col, 'latest', file[0])
                    if is_pmdr:
                        n2.message = "MDR file selected."
                        n2.type = "positive"
                        n2.timeout = 3
                        n2.spinner = False
                        mdr_filepath.set_text(file[0])
                    else:
                        n2.message = "'Latest' column not found in selected file. Please use RCC MDR."
                        n2.type = "negative"
                        n2.timeout = 3
                        n2.spinner = False
                else:
                    n2.message = "Filter exists in Data sheet. Please check RCC MDR."
                    n2.type = "negative"
                    n2.timeout = 3
                    n2.spinner = False
            else:
                ui.notify("'Data' sheet not found. Please check file.", type='negative')
        else:
            ui.notify("Today's date not found in MDR filename.", type = 'negative')
    else:
        ui.notify('No file selected.')

def check_file_for_sheet(sheetname, filename):
    xl = pd.ExcelFile(filename)
    return sheetname in xl.sheet_names

def check_file_for_filter(sheetname, filename):
    workbook = op.load_workbook(filename)
    sheet = workbook[sheetname]
    return sheet.auto_filter

def check_file_for_col(colname, filename):
    df = pd.read_excel(filename, sheet_name="Data",engine="openpyxl")
    return colname in df.columns

async def handle_execute():
    n = ui.notification("Executing... Please Wait.", type='ongoing', timeout=None, spinner=True)
    executeBtn.disable()
    rcc = rcc_filepath.text
    mdr = mdr_filepath.text
    global result
    result = await run.cpu_bound(compare_files, rcc, mdr)
    state['table'] = ui.table.from_pandas(result, pagination=0,column_defaults={'style': 'text-wrap: wrap'}).classes('my-sticky-header-table').style('width: 99%')
    state['table'].columns[0]['sortable'] = True
    state['table'].columns[1]['sortable'] = True
    state['table'].columns[2]['sortable'] = True
    state['table'].columns[3]['sortable'] = True
    state['input'] = ui.input('Search table').bind_value(state["table"], 'filter')
    n.message = "Complete!"
    n.type = "positive"
    n.timeout = 3
    n.spinner = False
    exportBtn.enable()
    clearBtn.enable()

async def reset_page():
    state["table"].delete()
    state['input'].delete()
    executeBtn.enable()
    exportBtn.disable()
    clearBtn.disable()
    ui.notify('Table Cleared.')

async def export():
    folder_path = os.path.dirname(rcc_filepath.text)
    filename = '/MDRComparisonOutput_'
    now = datetime.now()
    timestamp_string = folder_path + filename + now.strftime("%Y_%m_%d_%H_%M_%S") + '.csv'
    result.to_csv(timestamp_string, index= False)
    ui.notify("Table exported to CSV located in " + folder_path)

# Define the UI.
ui.add_css(
    """
    .my-sticky-header-table {
        /* height or max-height is important */
        max-height: 400px;
        /* this is when the loading indicator appears */
        /* prevent scrolling behind sticky top row on focus */
    }
    
    .my-sticky-header-table .q-table__top,
    .my-sticky-header-table .q-table__bottom,
    .my-sticky-header-table thead tr:first-child th {
        /* bg color is important for th; just specify one */
        background-color: #00b4ff;
    }
    
    .my-sticky-header-table thead tr th {
        position: sticky;
        z-index: 1;
    }
    
    .my-sticky-header-table thead tr:first-child th {
        top: 0;
    }
    
    .my-sticky-header-table.q-table--loading thead tr:last-child th {
        /* height of all previous header rows */
        top: 48px;
    }
    
    .my-sticky-header-table tbody {
        /* height of all previous header rows */
        scroll-margin-top: 48px;
    }  
    """
)

state = {}
with ui.row():
    ui.label("Link to MDR Folder:")
    ui.link("Link", "https://pfizer.sharepoint.com/:f:/r/sites/TASL/PMO/CDISC/Weekly%20Forum%20Meeting%20Minutes/2.%20MDR%20Library%20(and%20CDISC)%20Content/RCC%20Standard%20Metadata%20Files?csf=1&web=1&e=W9AlxW", new_tab= True)
with ui.header():
    ui.label('MDR Comparison Tool').style('font-size: 200%; font-weight: bold').classes('absolute-center')
    
with ui.row():
    ui.label('RCC Study Metadata Export File Path:')
    rcc_filepath = ui.label()

ui.button('Select RCC Metadata Export',on_click=choose_rcc_file)

with ui.row():
    ui.label('MDR File Path:')
    mdr_filepath = ui.label()

ui.button("Select Today's MDR",on_click=choose_mdr_file)

ui.space()

with ui.row():
    executeBtn = ui.button("Execute Comparison", on_click= lambda: handle_execute() if rcc_filepath.text != '' and mdr_filepath.text != '' else ui.notify('Please select both files to proceed.'))
    clearBtn = ui.button("Clear Table", on_click= reset_page)
    clearBtn.disable()
    exportBtn = ui.button("Export Table to CSV", on_click = export)
    exportBtn.disable()

ui.run(reload=False,native=True, port=native.find_open_port())