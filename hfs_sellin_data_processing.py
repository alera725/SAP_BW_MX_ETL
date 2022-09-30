# -*- coding: utf-8 -*-
"""
Created on Mon Aug 29 13:13:49 2022

@author: mxka1r54
"""

#Importar paqueterias
import os
os.chdir(r'C:\Users\MXKA1R54\OneDrive - Kellogg Company\Documents\ARG\LH4\BW\BW FILES') # relative path: scripts dir is under Lab
import numpy as np
import pandas as pd


distributors_dict = {'DSD' : 'HFS_2022_DSD.xlsx', 
                     'DSF' : 'HFS_2022_DSF.xlsx', 
                     'DUR' : 'HFS_2022_DUR.xlsx', 
                     'ELT' : 'HFS_2022_ELT.xlsx'}

#distributors_dict['DSD']
#distributors_dict[0]
#first_item = list(distributors_dict.items())[0][1]

len(distributors_dict)

for i in range(len(distributors_dict)):
     #i = 0
    #file_name = r'HFS_2022_ELT.xlsx'
    #distributor = 'ELT'
    
    fname = list(distributors_dict.items())[i][1]
    #dist = list(distributors_dict.items())[i][0]
    
    file_name = r'%s'%fname
    print(file_name)
    distributor = list(distributors_dict.items())[i][0]
    print(distributor)

    df_hfs_sellin = pd.read_excel(file_name, sheet_name = 'Sheet1')
    
    df_hfs_sellin.columns
    
    df_hfs_sellin.columns = df_hfs_sellin.columns.str.lower().str.replace(' ', '_')
    df_hfs_sellin.columns = df_hfs_sellin.columns.str.replace('_/_', '_')
    df_hfs_sellin.columns = df_hfs_sellin.columns.str.replace('_-_', '_')
    df_hfs_sellin.columns = df_hfs_sellin.columns.str.replace('-', '_')
    
    
    df_hfs_sellin.rename(columns = {'unnamed:_0': 'meals_snacks' , 'unnamed:_1' : 'sold_to_chain', 'unnamed:_2' : 'sold_to_chain_desc', 
           'unnamed:_3': 'payer', 'unnamed:_4' : 'payer_desc',
           'unnamed:_5' : 'material', 'unnamed:_6' : 'material_desc', 'unnamed:_7': 'fiscal_year_period'}, inplace = True)
    
    mask = ((df_hfs_sellin['sold_to_chain'] == 'Overall Result') | 
        (df_hfs_sellin['payer'] == 'Payer') | 
        (df_hfs_sellin['payer'] == 'Result') | 
        (df_hfs_sellin['material'] == 'Result'))
    df_hfs_sellin_clean = df_hfs_sellin[~mask]
    
    df_hfs_sellin_clean.fiscal_year_period.unique()
    yr = df_hfs_sellin_clean.fiscal_year_period.unique()[0][-4:]
    
    print(yr)
    
    output_file = "C:\\Users\\MXKA1R54\\OneDrive - Kellogg Company\\Documents\\ARG\\LH4\\BW\\BW FILES\\output\\" + distributor + "_" + yr + ".csv"
    
    df_hfs_sellin_clean[[
        'sold_to_chain',
        'sold_to_chain_desc',
        'payer',
        'payer_desc',
        'material',
        'material_desc',
        'fiscal_year_period',
        'net_kilos',
        'invoiced_kilos',
        'saleable_kilos',
        'gross_sales',
        'invoiced_value',
        'saleable_value',
        'temporary_price_reduction',
        'total_allowances',
        'cash_discounts',
        'government_discounts',
        'everyday_low_price',
        'trade_promotion_other',
        'sponsorship_in_store_adv',
        'store_openings_and_info_exchange',
        'unsaleables',
        'cleareance',
        'rollbacks',
        'growth_program',
        'free_goods',
        'margin_support',
        'distribution_comision',
        'write_off_pa',
        'net_sales',
        'net_cases',
        'invoiced_cases',
        'saleable_cases',
        'unsaleable_cases',
        'unsaleable_kilos',
        'price_per_kilo_gross_sales',
        'price_per_kilo_net_sales',
        'cost_of_goods_sold',
        'raw_material',
        'purchased_product',
        'fixed_overhead',
        'variable_overhead',
        'packaging_material',
        'direct_labor',
        'indirect_labor',
        'gross_margin'
    ]].to_csv(output_file, index = False)