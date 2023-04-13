import pandas as pd
import numpy as np
import functions


def initialize_lkups_fields(tab):
    rows_column_names = []
    cols_column_names = []
    
    rows_descs = [tab[tab == 'Y-AXIS : ROWS'].stack().index.tolist()[0],
                 (len(tab)-1, tab.columns[0])]
    cols_descs = [tab[tab == 'X-AXIS : COLUMNS'].stack().index.tolist()[0],
                 (rows_descs[0][0]-2, tab.columns[0])]
    
    rows_fields = [tab.loc[rows_descs[0][0]:rows_descs[1][0], rows_descs[0][1]:rows_descs[1][1]]]
    cols_fields = [tab.loc[cols_descs[0][0]:cols_descs[1][0], cols_descs[0][1]:cols_descs[1][1]]]
    rows_column_names.append('Descriptions')
    cols_column_names.append('Descriptions')
    
    
    rows_items = [tab[tab == 'Row Num'].stack().index.tolist()[0],
                 (len(tab)-1, tab[tab == 'Row Num'].stack().index.tolist()[0][1])]
    cols_items = [tab[tab == 'Col Num'].stack().index.tolist()[0],
                 (rows_items[0][0]-2, tab[tab == 'Col Num'].stack().index.tolist()[0][1])]
    
    rows_fields.append(tab.loc[rows_items[0][0]:rows_items[1][0], rows_items[0][1]:rows_items[1][1]])
    cols_fields.append(tab.loc[cols_items[0][0]:cols_items[1][0], cols_items[0][1]:cols_items[1][1]])
    rows_column_names.append('Items')
    cols_column_names.append('Items')
    
    
    if 'Totals' in tab.iloc[tab[tab == 'Y-AXIS : ROWS'].stack().index.tolist()[0][0]].values:
        
        totals_pos = tab[tab == 'Totals'].stack().index.tolist()
        rows_totals_pos = [i for i in totals_pos if tab[tab == 'Row Num'].stack().index.tolist()[0][0] == i[0]][0]
        
        rows_totals = [rows_totals_pos,
                      (len(tab)-1, rows_totals_pos[1])]
        
        rows_fields.append(tab.loc[rows_totals[0][0]:rows_totals[1][0], rows_totals[0][1]:rows_totals[1][1]])
        rows_column_names.append('Totals')
        
    
    if 'Totals' in tab.iloc[tab[tab == 'X-AXIS : COLUMNS'].stack().index.tolist()[0][0]].values:
        
        totals_pos = tab[tab == 'Totals'].stack().index.tolist()
        cols_totals_pos = [i for i in totals_pos if tab[tab == 'Col Num'].stack().index.tolist()[0][0] == i[0]][0]
        
        cols_totals = [cols_totals_pos,
                      (len(tab)-1, cols_totals_pos[1])]
        
        cols_fields.append(tab.loc[cols_totals[0][0]:cols_totals[1][0], cols_totals[0][1]:cols_totals[1][1]])
        cols_column_names.append('Totals')
    
    
    
    rows = pd.concat(rows_fields, axis=1).astype('str').replace('nan', np.nan)
    rows.columns = rows_column_names
    rows = rows.reset_index(drop=True).drop([0, 1]).reset_index(drop=True)
    
    cols = pd.concat(cols_fields, axis=1).astype('str').replace('nan', np.nan)
    cols.columns = cols_column_names
    cols = cols.reset_index(drop=True).drop([0, 1]).reset_index(drop=True)
    
    return rows, cols
    
    
def preprocess_totals(raw_totals, items):
    if pd.isna(raw_totals):
        return raw_totals
    else:
        totals_list = []
        raw_totals_list = list(map(lambda s: s.split(','), raw_totals.strip('( )').split('+')))
        raw_totals_list = [item for sublist in raw_totals_list for item in sublist]
        raw_totals_list = list(map(lambda s: s.strip(), raw_totals_list))
        
        for i, item in enumerate(raw_totals_list):
            if ':' in item:
                first_item, last_item = tuple(map(lambda s: s.strip(), item.split(':')))
                new_item = items[items.index(first_item):items.index(last_item) + 1]
                totals_list += new_item
            else:
                totals_list.append(item)

        return totals_list
        

def add_totals(lookups_totals, form, tab_name, third_column_name, fourth_column_name):
    totals_list = form.dropna(subset=['Totals']).drop(columns=['Descriptions'])
    totals_dict = {}
    for index, row in totals_list.iterrows():
        totals_dict.setdefault(row[0], row[1])
        
    for key, value in totals_dict.items():
        for v in value:
            buf_desc = f'TOTAL. {form[form["Items"] == key]["Descriptions"].values[0].strip()}'
            buf_df = pd.DataFrame({'Report Name': tab_name,
                                   'Portfolio Item': v,
                                   third_column_name: key,
                                   fourth_column_name: buf_desc}, index=[0])

            lookups_totals = pd.concat([lookups_totals, buf_df], ignore_index=True)
            
    return lookups_totals


def preprocess_of_whichs(lookups_of_whichs, form, tab_name, third_column_name, fourth_column_name):
    sector_nums = form.loc[:, ['Descriptions']].applymap(lambda s: s[:s.index(' ')])
    form_sectors = form.copy(deep=True)
    form_sectors['Sector'] = sector_nums
    
    sector_nums.columns = ['Sector']

    for index, row in sector_nums.iterrows():
        sector_nums[row[0]] = sector_nums.loc[:, ['Sector']].applymap(lambda s: s.startswith(row[0]) and s != row[0])
    
    for index, row in form.iterrows():
        if 'of which' in row[0].lower():
            sector_num = sector_nums.iloc[index, :].values[0]
            sector_df = sector_nums[sector_nums['Sector'] == sector_num]
            is_true = sector_df == True
            
            parent_sectors = list(sector_df.loc[:, is_true.squeeze()].columns)
            
            for parent_sector in parent_sectors:
                parent_item = form_sectors[form_sectors['Sector'] == parent_sector]['Items'].values[0]
                item = row[1]
                
                buf_desc = f'OF WHICH. {form_sectors[form_sectors["Items"] == parent_item]["Descriptions"].values[0].strip()}'
                buf_df = pd.DataFrame({'Report Name': tab_name,
                                       'Portfolio Item': item,
                                       third_column_name: parent_item,
                                       fourth_column_name: buf_desc}, index=[0])

                lookups_of_whichs = pd.concat([lookups_of_whichs, buf_df], ignore_index=True)
                
    
    return lookups_of_whichs


def preprocess_sub_items(lookups_sub_items, tab_name, third_column_name, fourth_column_name):
    extended_lookups = lookups_sub_items[lookups_sub_items[fourth_column_name].str.contains('TOTAL|OF WHICH')]
    sub_items = extended_lookups[third_column_name].unique()
    
    for index, row in extended_lookups.iterrows():
        if row[1] in sub_items:
            buf_df = extended_lookups[extended_lookups[third_column_name] == row[1]]
            buf_df[third_column_name] = row[2]
            buf_df[fourth_column_name] = row[3]
            lookups_sub_items = pd.concat([lookups_sub_items, buf_df], ignore_index=True)
           
        
    return lookups_sub_items


def create_lookups(raw_form, items_name, tab_name):
    import re
    
    third_column_name = f'Report {items_name} Item'
    fourth_column_name = f'Report {items_name} Description'
    
    columns = ['Report Name', 'Portfolio Item', third_column_name, fourth_column_name]
    
    items = raw_form['Items'].dropna().tolist()
    form = raw_form.copy(deep=True)
    form.dropna(how='any', inplace=True, subset=['Descriptions', 'Items'])
    form.reset_index(inplace=True, drop=True)
    form['Descriptions'] = form.loc[:, ['Descriptions']].applymap(lambda s: ' '.join(s.split()))
    form['Descriptions'] = form.loc[:, ['Descriptions']].applymap(lambda s: s[re.search(r"\d", s).start():])
    
    lookups = pd.DataFrame(tab_name, index=np.arange(len(form)), columns=columns)
    lookups['Portfolio Item'] = form['Items']
    lookups[third_column_name] = form['Items']
    lookups[fourth_column_name] = form['Descriptions']
    
    if 'Totals' in form.columns:
        totals_form = form.copy(deep=True)
        totals_form['Totals'] = totals_form.drop(['Descriptions', 'Items'],
                                                 axis=1).applymap(lambda x: preprocess_totals(x, items))
        
        lookups = add_totals(lookups, totals_form, tab_name, third_column_name, fourth_column_name)
    
    
    lookups = preprocess_of_whichs(lookups, form, tab_name, third_column_name, fourth_column_name)
    
    lookups = preprocess_sub_items(lookups, tab_name, third_column_name, fourth_column_name)
    
    return lookups
    
    
def collect_lkups(report_tabs, tab_names):
    lkups = {}
    
    for tab_name in tab_names:
        tab = functions.initialize_tab(report_tabs, tab_name)
        forms = initialize_lkups_fields(tab)
        
        lkups.setdefault(tab_name, dict())
        for i in range(len(forms)):
            items_name = 'Row' if i == 0 else 'Column'
            lookups = create_lookups(forms[i], items_name, tab_name)
            lkups[tab_name].update({items_name: lookups})
            
    return lkups