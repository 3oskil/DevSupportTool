import dataframe_image as dfi
import pandas as pd
import numpy as np
import warnings
import re


def load_model(data_model, data_model_ctryspec):
    """
    Returns tables and database dicts.

            Parameters:
                    data_model (pandas.io.excel._base.ExcelFile): Data Model uploaded as pandas ExcelFile
                    data_model_ctryspec (pandas.io.excel._base.ExcelFile): Data Model Country Specific uploaded as pandas ExcelFile 

            Returns:
                    tables (dict): Dict where keys - names of tables from Data Model,
                                              values - tables from Data Model in view of DataFrames.
                    database (dict): Dict where keys - names of map and lkup tables from Data Model,
                                                values - map and lkup tables from Data Model in view of DataFrames.
    """
    tables, database = dict(), dict()
    for sheet_name in data_model.sheet_names:
        if sheet_name.startswith('Tbl'):
            tables.setdefault(' '.join(sheet_name.lower().split(' ')[1:]), data_model.parse(sheet_name))
        if sheet_name.startswith('LKUP') or sheet_name.startswith('MAP'):
            database.setdefault(sheet_name.lower(), data_model.parse(sheet_name))

    for sheet_name in data_model_ctryspec.sheet_names:
        if sheet_name.startswith('Tbl'):
            tables.setdefault(' '.join(sheet_name.lower().split(' ')[1:]), data_model_ctryspec.parse(sheet_name))
        if sheet_name.startswith('LKUP') or sheet_name.startswith('MAP'):
            database.setdefault(sheet_name.lower(), data_model_ctryspec.parse(sheet_name))

    return tables, database


def load_spec(spec, tab_names):
    """
    Returns report_tabs dict.

            Parameters:
                    spec (pandas.io.excel._base.ExcelFile): Specification uploaded as pandas ExcelFile
                    tab_names (list): List of tab names that need to be analysed

            Returns:
                    report_tabs (dict): Dict where keys - names of tabs from Specification,
                                                   values - tab from Specification in view of DataFrames.
    """
    report_tabs = dict()
    for tab_name in tab_names:
        report_tabs.setdefault(tab_name, spec.parse(tab_name))

    return report_tabs


def initialize_tab(report_tabs, tab_name):
    tab = report_tabs[tab_name]
    return tab


def initialize_location(tab):
    row_loc = [tab[tab == 'Row Num'].stack().index.tolist()[0], (len(tab) - 1, tab.columns[-1])]
    col_loc = [tab[tab == 'Col Num'].stack().index.tolist()[0], (row_loc[0][0] - 2, tab.columns[-1])]
    return col_loc, row_loc


def initialize_form(location, tab):
    if tab.loc[location[0]] not in ('Col Num', 'Row Num'):
        raise Exception(f'Wrong location {location[0]}, value should be "Col Num" or "Row Num", not {tab.loc[location[0]]}')
    
    raw_form = tab.loc[location[0][0]:location[1][0], location[0][1]:]
    raw_form = raw_form.dropna(axis=0, how='all').dropna(axis=1, how='all')
    form = raw_form.drop(columns=raw_form.columns[raw_form.iloc[:2].isnull().all()])
    
#     if tab.loc[location[0]] == 'Col Num':
#         form.drop(columns=[form.columns[1]])
    return form


def identify_totals(form):
    totals = []
    if 'Totals' in form.iloc[0, :].values:
        i = 0
        for index, row in form.iterrows():
            if i > 1:
                if not pd.isna(row[1]):
                    if not pd.isna(row[0]):
                        item = str(row[0])
                        buf_item, item_part = item, 0
                    else:
                        item_part += 1
                        item = buf_item + '_' + str(item_part)
                    totals.append(item)
            i += 1
    return totals


def identify_tables(cols_form, rows_form, tables):
    tab_tables = []
    for table_name in tables:
        first_row = map(lambda s: str(s).strip().lower(), cols_form.iloc[0].to_list() + rows_form.iloc[0].to_list())
        if table_name in first_row:
            tab_tables.append(table_name)
    return tab_tables


def clean_form(form, tab_tables):
    drop_columns = []
    for i, value in enumerate(form.iloc[0].values):
        if isinstance(value, str) and value.strip().lower() not in tab_tables + ['col num', 'row num']:
            drop_columns.append(i)

    clear_form = form.drop(columns=(form.columns[drop_columns])).reset_index(drop=True)
    return clear_form


def veiw_form(form, form_name):
    file_name = f"Logs/{form_name}.jpg"
    dfi.export(form, file_name, max_rows=(-1))
    print(f"Please check {file_name} in root folder")


def split_not(actual_table, raw_value, map_table):
    in_not = re.findall('\\(.[^)]+\\)', raw_value)
    out_not = re.split('NOT\\s\\(.[^)]+\\)', raw_value)
    clear_in_not = [a for b in list(map((lambda string: string.lstrip('(').rstrip(')').split(', ')), in_not)) for a in iter(b)]
    clear_out_not = [a for b in list(map((lambda string: string.strip().split(', ')), out_not)) for a in iter(b) if a]
    value_in_not = add_fill_values(actual_table, clear_in_not, map_table)
    value_out_not = add_fill_values(actual_table, clear_out_not, map_table)
    return value_in_not, value_out_not


def get_actual_table(t, f, tables, joins):
    t_names = tuple(set([item for sublist in [v for k, v in joins.items() if t in v] for item in iter(sublist)]))
    t_dfs = {}
    for t_name in t_names:
        if not 'map' in t_name and not 'lkup' in t_name:
            t_dfs[t_name] = tables[t_name]
            
    if f in t_dfs[t]['Column Name'].values:
        actual_table = t
    else:
        t_dfs.pop(t)
        for t_df_key, t_df_value in t_dfs.items():
            if f in t_df_value['Column Name'].values:
                actual_table = t_df_key
                break
        else:
            actual_table = t
    return actual_table


def get_map_table(actual_table, f, tables, database):
    table = tables[actual_table]
    try:
        map_table_name = table[table['Column Name'] == f]['Data Type'].values[0].lower()
        if 'map' in map_table_name or 'lkup' in map_table_name:
            return database[map_table_name]
        return
    except:
        return


def fill_not(actual_table, value_in_not, map_table):
    if map_table is not None:
        full_allowed_list = map_table.iloc[:, 0].values
        allowed_list = [val for val in full_allowed_list if val not in value_in_not]
    else:
        return ('NOT', value_in_not)
    return allowed_list


def main_fill_values(field, raw_value, tables, database, jdx, joins):
    raw_value = raw_value.strip() if isinstance(raw_value, str) else raw_value
    t, f = field.split('.')
    actual_table = get_actual_table(t, f, tables, joins)
    
    if pd.isna(raw_value):
        value = None
    else:
        map_table = get_map_table(actual_table, f, tables, database)
        if 'NOT' in raw_value:
            value_in_not, value_out_not = split_not(actual_table, raw_value, map_table)
            if value_out_not:
                value = [val for val in value_out_not if val not in value_in_not]
            else:
                value = fill_not(actual_table, value_in_not, map_table)
        else:
            value = add_fill_values(actual_table, raw_value, map_table)
    return actual_table, value


def add_fill_values(actual_table, intl_raw_value, map_table):
    if not isinstance(intl_raw_value, list):
        intl_raw_value = list(map((lambda x: x.strip()), intl_raw_value.split(',')))
    intl_value = []
    for v in intl_raw_value:
        if 'ax_' in v and map_table is not None:
            col_name = list(filter((lambda string: string.startswith('Parent')), map_table.columns))[0]
            for i, row in map_table.iterrows():
                if isinstance(row[col_name], str) and v in row[col_name]:
                    intl_value.append(row['Code'])
        else:
            intl_value.append(v)

    return intl_value


def initialize_interval(value):
    perpetual = {'y': 500, 'm': 6000, 'w': 26071, 'd': 182625}
    lvalue = value.split()
    time_units = lvalue[-1][0]
    interval = [0, time_units]
    if len(lvalue) == 3:
        if 'gt' in lvalue[0]:
            interval[0] = (
             int(lvalue[1]) if 'e' in lvalue[0] else eval(lvalue[1]) + 1, perpetual[time_units])
        if 'lt' in lvalue[0]:
            interval[0] = (
             -perpetual[time_units], int(lvalue[1]) if 'e' in lvalue[0] else eval(lvalue[1]) - 1)
    else:
        if len(lvalue) == 5:
            interval[0] = (int(lvalue[1]) if 'e' in lvalue[0] else eval(lvalue[1]) + 0.1,
             int(lvalue[3]) if 'e' in lvalue[2] else eval(lvalue[3]) - 0.1)
        else:
            raise Exception(f"Unknown case {value}")
            
    return interval


def is_intersects(term1, term2):
    t1, t2 = term1[0], term2[0]
    if term1[1] != term2[1]:
        convert = {'d': 1, 'w': 7, 'm': 30, 'y': 365}
        t1, t2 = tuple(map((lambda x: x * convert[term1[1]]), t1)), tuple(map((lambda x: x * convert[term2[1]]), t2))
        
    return t1[0] <= t2[1] and t1[1] >= t2[0]


def identify_non_reportable(tab_tables, value):
    without_nones = {k: v for k, v in value.items() if v != None}
    rep = set(map((lambda x: x.split('.')[0]), list(without_nones.keys())))
    non_rep = [i for i in tab_tables if i not in rep]
    
    return non_rep


def process_items(items, tab_tables, joins):
    all_tables = set()
    for v in items.values():
        for k1 in v.keys():
            all_tables.add(k1.split('.')[0])

    main_add_tables = dict()
    for main in joins.keys():
        main_add_tables[main] = [i for i in all_tables if i in joins[main]]

    items_merged = dict()
    for k, v in items.items():
        for k1, v1 in v.items():
            main_tables = [key for key, value in main_add_tables.items() if k1.split('.')[0] in value]
            for main_table in main_tables:
                items_merged.setdefault(main_table, {}).setdefault(k, {})[k1] = v1

    non_reportable = {t: [] for t in items_merged.keys()}
    for k, v in items_merged.items():
        for k1 in list(v):
            if not any((i for i in v[k1].values())):
                items_merged[k].pop(k1)
                non_reportable[k].append(k1)

    return items_merged, non_reportable


def set_items(form, tab_tables, tables, database, totals, jdx, logic):
    tables_id, fields_id = dict(), dict()
    fields, items = dict(), dict()
    joins = dict()
    for rec_iter, row in form.iterrows():
        if rec_iter == 0:
            for f_iter, j in enumerate(row):
                if isinstance(j, str) and j.lower().strip() in tab_tables:
                    tables_id.setdefault(f_iter, j.lower().strip())
            
        if rec_iter == 1:
            num = list(tables_id.keys())[0]
            for j in row[num:]:
                fields_id.setdefault(num, j)
                num += 1

            f_keys = list(fields_id.keys())
            for k in f_keys:
                if k in tables_id:
                    fields.setdefault(k, tables_id[k] + '.' + fields_id[k])
                    buf_key = k
                else:
                    fields.setdefault(k, tables_id[buf_key] + '.' + fields_id[k])
            
            for t in tab_tables:
                if t in logic['core'].keys():
                    joins[t] = logic['core'][t] + logic[jdx][t]
                    
            
        if rec_iter > 1:
            if pd.isna(row[0]):
                item_part += 1
                item = buf_item + '_' + str(item_part)
            else:
                item = str(row[0])
                buf_item, item_part = item, 0
            
            if item not in totals:
                items.setdefault(item, {})
                for f_iter, j in enumerate(row):
                    if f_iter == 0:
                        continue
                    if f_iter in fields.keys():
                        raw_value = j
                        field = fields[f_iter]
                        if isinstance(raw_value, str):
                            if 'elsewhere_reported' in raw_value.lower():
                                items[item].setdefault(field, raw_value)
                        t, value = main_fill_values(field, raw_value, tables, database, jdx, joins)
                        items[item].setdefault('.'.join([t, field.split('.')[1]]), value)


    items_merged, non_reportable = process_items(items, tab_tables, joins)
    return items_merged, non_reportable


def is_elsewhere(items, main_table, item):
    return any((i for i in items[main_table][item].values() if isinstance(i, str) if 'elsewhere_reported' in i.lower()))


def identify_interval_fields(items):
    interval_fields = []
    poss_values = ('lt', 'gt')
    temp_dict = {}
    for k, v in items.items():
        for k1, v1 in v.items():
            for k2, v2 in v1.items():
                temp_dict.setdefault(k2, []).extend(v2 if isinstance(v2, list) else [v2])

    for key, value in temp_dict.items():
        temp = []
        for i in value:
            if i not in temp and i:
                temp.append(i)

        temp_dict[key] = temp

    for key, value in temp_dict.items():
        if any((s1 in s2 for s1 in poss_values for s2 in value)) and any((any((c.isdigit() for c in s)) for s in value)):
            interval_fields.append(key)

    return interval_fields


def analyse_overlaps(items):
    overlaps = {x: dict() for x in items.keys()}
    interval_fields = identify_interval_fields(items)
    for t, v in items.items():
        for i, item_1 in enumerate(v):
            if i == len(v) - 1:
                break
            fields = v[item_1]
            is_overlap = False
            for j, item_2 in enumerate(v):
                overlap_fields = []
                if not j <= i:
                    if item_1.split('_')[0] == item_2.split('_')[0]:
                        continue
                    if not is_elsewhere(items, t, item_1):
                        if is_elsewhere(items, t, item_2):
                            continue
                        for k, f in enumerate(fields):
                            v1, v2 = v[item_1][f], v[item_2][f]
                            if v1 and v2:
                                if f in interval_fields:
                                    term1, term2 = initialize_interval(v1[0]), initialize_interval(v2[0])
                                    if not is_intersects(term1, term2):
                                        is_overlap = False
                                        break
                            else:
                                is_overlap = True
                                overlap_fields.append(f)
                                continue
                            if not isinstance(v1, tuple):
                                if isinstance(v2, tuple):
                                    if not isinstance(v1, tuple):
                                        if isinstance(v2, tuple):
                                            if set(v1).issubset(v2[1]):
                                                is_overlap = False
                                                break
                                    elif isinstance(v1, tuple) and not isinstance(v2, tuple):
                                        if set(v2).issubset(v1[1]):
                                            is_overlap = False
                                            break
                                        else:
                                            is_overlap = True
                                            overlap_fields.append(f)
                                            continue
                                if set(v1).isdisjoint(set(v2)):
                                    is_overlap = False
                                    break
                                else:
                                    is_overlap = True
                                    overlap_fields.append(f)
                            else:
                                is_overlap = True
                                overlap_fields.append(f)

                        if is_overlap:
                            if item_1 in overlaps[t].keys():
                                overlaps[t][item_1].update({item_2: overlap_fields})
                            else:
                                overlaps[t].setdefault(item_1, {item_2: overlap_fields})
                    continue

    return overlaps


def main_add(tab, tables, database, date_fields, jdx):
    form_locations = initialize_location(tab)
    form_overlaps = {}
    for form_location in form_locations:
        form = initialize_form(form_location[0], tab)
        form_name = form_location[1]
        totals = identify_totals(form)
        form_tables = identify_tables(form, tables)
        clear_form = clean_form(form, form_tables)
        veiw_form(clear_form, form_name)
        items, non_reportable = set_items(clear_form, form_tables, tables, database, totals, jdx)
        overlaps = analyse_overlaps(items, date_fields)
        form_overlaps.setdefault(form_name, [overlaps, totals])

    return form_overlaps


def main(tables, database, spec, int_cals, jdx, tab_names):
    jdxs = ('at', 'be', 'ch', 'de', 'dk', 'es', 'fin', 'fr', 'gb', 'ie', 'it', 'lu',
            'nl')
    if jdx.lower() not in jdxs:
        raise Exeption(f"There is no {jdx} in the list. Please enter one of:\n{jdxs}")
    report_tabs = load_spec(spec, tab_names)
    date_fields = ('Past Due', 'Interest Past Due', 'Residual Maturity', 'Revised Residual Maturity',
                   'Capital Residual Maturity', 'Original Maturity', 'Forward Start',
                   'Start Interval', 'Trade Interval', 'Interest Reset Interval',
                   'Value Interval', 'interestResetTerm', 'floatingRateResetTerm',
                   'noticePeriod', 'Current Valuation Ratio', 'Current Valuation Ratio (Gross)',
                   'Movement Interval', 'Restructured Interval', 'Cash Flow Month',
                   'Arrears Percentage', 'Gov Protection Request Interval', 'Gov Protection Acceptance Interval',
                   'Gov Protection Denial Interval', 'Debtor Subrogation Interval',
                   'Renewal Interval', 'Purchase Loan Interval', 'Refinanced Interval',
                   'Renegotiation Interval', 'Performance Change Interval', 'Special Surveillance Interval',
                   'Disbursement Interval', 'Forborne Interval', 'Recognition Interval',
                   'Litigation End Interval', 'Transaction Interval')
    overlaps = {}
    for tab_name in report_tabs:
        tab = initialize_tab(report_tabs, tab_name)
        overlaps.setdefault(tab_name, main_add(tab, tables, database, date_fields, jdx))

    return overlaps


def show_result(tab):
    for form_name, form in tab.items():
        ols, tots = form
        ols_ind = False
        for k, v in ols.items():
            if v:
                ols_ind = True
                break

        if ols_ind:
            print(f"\n{form_name}' overlaps:\n")
            for table, ovlps in ols.items():
                if ovlps:
                    print((' '.join(list(map((lambda s: s.capitalize()), table.replace('_', ' ').split())))), ':', sep='', end='\n')
                    for item, fields in ovlps.items():
                        print(item)

                    print()

        else:
            print(f"\n\nNO OVERLAPS IN {form_name}!")
        if tots:
            print(f"\n\n{form_name}' totals:\n")
            for item in tots:
                print(item)