# os.chdir(source_code_dir)
from utils_io import format_excel_sheet_cols
import pandas as pd
import numpy as np
import os, sys, glob
import humanize
import re
import xlrd

import json
import itertools
import requests
from urllib.parse import urlencode
#from urllib.request import urlopen
#import requests, xmltodict
import time, datetime
import math
from pprint import pprint
import gc
from tqdm import tqdm
tqdm.pandas()
import pickle

import logging
import zipfile
import warnings
import argparse

import warnings
warnings.filterwarnings("ignore")

from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from openpyxl.utils import units
from openpyxl.styles import Border, Side, PatternFill, GradientFill, Alignment
from openpyxl import drawing

import ipywidgets as widgets
from IPython.display import display
from ipywidgets import Layout, Box, Label

from utils_io import save_df_lst_to_excel
from utils_io import logger
if len(logger.handlers) > 1:
    for handler in logger.handlers:
        logger.removeHandler(handler)
    from utils_io import logger

def read_df(data_source_dir, fn, sheet_name):
    # fn = 'LP_Pharm_form_unify_2023_03_31.xlsx'
    # sheet_name = 'привязка_ЕСКЛП'
    # fn = fn_check_file_drop_douwn.value
    # sheet_name = sheet_name_drop_douwn.value
    df_lp = None
    if fn is not None and sheet_name is not None:
        try: 
            df_lp = pd.read_excel(os.path.join(data_source_dir, fn), sheet_name=sheet_name)
            display(df_lp.head(2))
            logger.info(f"Входной файл содержит {df_lp.shape[0]} позиций")
        except Exception as err:
            logger.error(f"{str}")
            sys.exit(2)
    else:
        logger.error(f"Не еделены файл и лист Excel")
        sys.exit(2)
    req_cols = ['МНН', 'ФВ', 'ФВ_краткая']
    cur_cols = list(df_lp.columns)
    if not set(req_cols).issubset(cur_cols):
        logger.error(f"Неправлиьные названия колонок в листе '{sheet_name}': {cur_cols}")
        logger.error(f"Названия колонок должны содержать: {req_cols}")
        sys.exit(2)
    return df_lp
      

def update_df(df_lp, smnn_list_df, n_rows=np.inf, debug=False):
    # n_rows = 1
    new_cols = ['ФВ[]', 'n_forms', 'matched', 'ФВ[МНН]', 'n_forms_all']
    for i_row, row in df_lp.iterrows():
        mnn_standard = row['МНН']
        ph_form = row['ФВ']
        ph_form_short = row['ФВ_краткая']
        mask = (smnn_list_df['mnn_standard'].notnull() & (smnn_list_df['mnn_standard']==mnn_standard) &
        # mask = (smnn_list_df['mnn_standard'].notnull() & (smnn_list_df['mnn_standard'].str.contains(fr"^{mnn_standard}$", flags=re.I, regex=True)) &
                smnn_list_df['form_standard'].notnull() & smnn_list_df['form_standard'].str.contains(fr"^{ph_form_short}", flags=re.I, regex=True))
        form_standard_lst = smnn_list_df[mask]['form_standard'].values
        n_forms = len(form_standard_lst)
        if n_forms > 0:
            # form_standard_lst = [ph_f.lower().capitalize() for ph_f in form_standard_lst]
            # form_standard_lst = np_unique_nan(form_standard_lst)
            form_standard_lst = set(form_standard_lst)
            n_forms = len(form_standard_lst)
            form_standard_lst = [ph_f.lower() for ph_f in form_standard_lst]
            form_standard_lst_str = "; ".join(form_standard_lst)
            if n_forms > 1: 
                form_standard_lst_str = '['+ form_standard_lst_str + ']'
            if ph_form in form_standard_lst:
                matched = 1
            else: matched = 0
            form_standard_lst_all_str = None
            n_forms_all = None
        else:
            form_standard_lst_str = '#НД'
            matched = 0
            mask1 = (smnn_list_df['mnn_standard'].notnull() & (smnn_list_df['mnn_standard']==mnn_standard))
            form_standard_lst_all = smnn_list_df[mask1]['form_standard'].values
            form_standard_lst_all = set(form_standard_lst_all)
            n_forms_all = len(form_standard_lst_all)
            form_standard_lst_all = [ph_f.lower() for ph_f in form_standard_lst_all]
            form_standard_lst_all_str = "; ".join(form_standard_lst_all)
            if n_forms_all > 1: 
                form_standard_lst_all_str = '['+ form_standard_lst_all_str + ']'
        if debug: 
            print(i_row, form_standard_lst_str)
            if i_row > n_rows: break
        df_lp.loc[i_row, new_cols] = [form_standard_lst_str, n_forms, matched, form_standard_lst_all_str, n_forms_all]
    pd.options.display.max_colwidth = 200
    display(df_lp.head(3))
    logger.info(F"Позиций с не найдеными формами: {df_lp[df_lp['n_forms'] ==0].shape[0]}")
    return df_lp

def check_file(
    data_source_dir, data_processed_dir,
    fn, sheet_name,
    smnn_list_df,
    n_rows=np.inf, debug=False):
    df_lp = read_df(data_source_dir, fn, sheet_name)
    df_lp = update_df(df_lp, smnn_list_df, n_rows=2, debug=debug)
    # df_lp = update_df(df_lp, smnn_list_df, n_rows=2000, debug=False)
    fn_save = save_df_lst_to_excel([df_lp], [sheet_name], data_processed_dir, fn)
    col_width_lst = [30, 30, 15, 40, 10, 10, 40, 10]
    format_excel_sheet_cols(data_processed_dir, fn_save, col_width_lst, sheet_name)
    logger.info(f"Обработанный фaйл '{fn_save}' сохранене в директорию '{data_processed_dir}'")
    # return df_lp
 
