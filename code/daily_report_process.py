import pandas as pd
import datetime
import os
from datetime import timedelta
import xlsxwriter
import warnings
warnings.filterwarnings('ignore')

today = datetime.datetime.today().date()
tday  =today.strftime("%d/%m/%Y")
yesterday = today - timedelta(days=1)

##======================================================================================================================

def mapping_type(mappath):
    msterdatapath_ =  mappath + "/Master_data/"

    usetype = pd.read_csv(msterdatapath_ + "usetype.csv")
    usemap = dict(zip(usetype['usetypekey'],usetype['eng_usename']))

    consttype = pd.read_csv(msterdatapath_ + "constructiontype.csv")
    construcmap = dict(zip(consttype['constructiontypekey'],consttype['eng_constructiontypename']))

    occptype=  pd.read_csv(msterdatapath_ + "occupancy.csv")
    occpmap = dict(zip(occptype['occupancykey'],occptype['occupancyname']))

    subusetype= pd.read_csv(msterdatapath_ + "subusetype.csv")
    subusemap = dict(zip(subusetype['subusetypekey'],subusetype['eng_subusename']))

    zonetype =pd.read_csv(msterdatapath_ + "zone.csv")
    zonemap = dict(zip(zonetype['zonename'], zonetype['eng_zonename']))
    # zonemap = dict(zip(zonetype['zonekey'],zonetype['eng_zonename']))
    # final_output['Zone_Type'] = final_output['zonekey'].map(zzz)

    gattype = pd.read_csv(msterdatapath_ + "gat.csv")
    # gattype['gatname_z'] = gattype['gatname'].astype(str) + "_" + gattype['zonetype'].astype(str)
    gattype['gatname_z'] = gattype['gatname'].astype(str) + "_" + gattype['zonetype'].astype(str) + "_" + gattype['mar_gatname'].astype(str)

    gatnamemap = dict(zip(gattype['gat'], gattype['gatname_z']))

    return zonemap,usemap,construcmap,occpmap,subusemap,gatnamemap,msterdatapath_


##======================================================================================================================
def zonegatwise_TDdailyreport(infile,mappath,outpath,zonemap,mailreport):
    tddata = pd.read_excel(infile)
    tddata = tddata[tddata['receiptdate'] == '2023-03-31']
    tddata['receiptdate'] = pd.to_datetime(tddata['receiptdate']).dt.date
    # tddata = pd.read_excel(infile, sheet_name=tdsheetname)
    #-------------------------------------------------------------------------------------------------------------------
    grpbytot = tddata.groupby(['ezname', 'gatname']).agg({'paidamount': 'sum'}).reset_index()
    grpbytot['eng_zone'] = grpbytot['ezname'].map(zonemap)
    # aaaa = grpbytot.melt(id_vars=['ezname', 'gatname','eng_zone'], var_name='groubysum', value_name='paidamount')
    pvotable = grpbytot.pivot_table(index=['eng_zone', 'ezname'], columns='gatname', values='paidamount')
    bbb = list(range(1, 19))
    dfff = pd.DataFrame(pvotable, columns=bbb)
    pp = dfff.reset_index()
    #-------------------------------------------------------------------------------------------------------------------
    lissy = ['Nigdi Pradhikaran', 'Akurdi', 'Chinchwad', 'Thergaon', 'Sangvi', 'Pimpri Waghere',
             'Pimpri Nagar', 'MNP Bhavan', 'Fugewadi Dapodi', 'Bhosari', 'Charholi',
             'Moshi', 'Chikhali', 'Talvade', 'Kivle', 'Dighi Bopkhel', 'Wakad']

    d = {v: k for k, v in enumerate(lissy)}
    df_TD = pp.sort_values('eng_zone', key=lambda x: x.map(d), ignore_index=True)
    #-------------------------------------------------------------------------------------------------------------------
    collen = df_TD.columns.to_list()[2:]
    df_TD['Grand Total'] = df_TD[collen].sum(axis=1)
    df_TD['Recovery'] = df_TD['Grand Total'].apply(lambda x: round(x / 10000000, 2))
    df_TD.index = df_TD.index + 1
    df_TD.loc["Grand Total"] = df_TD.sum(numeric_only=True)
    #-------------------------------------------------------------------------------------------------------------------
    df_TD =  df_TD.drop(columns='eng_zone')
    df_TD1 = df_TD.reset_index()
    final_df_TD = df_TD1.rename(
        columns={'index': 'अ.क्र.', 'ezname': 'विभागीय कार्यालय', 'Grand Total': 'एकूण', 'Recovery': 'वसूली'})
    final_df_TD = final_df_TD.replace("Grand Total",'एकूण')
    #===================================================================================================================
    maildata = final_df_TD[['अ.क्र.', 'विभागीय कार्यालय', 'वसूली']]
    maildata1 = maildata.T.reset_index()
    # maildata1['वसूली'] = maildata1['वसूली'].astype(float)
    # maildata1 = maildata1.fillna('एकूण')
    maildata1.to_csv(mailreport+f"{today}_collectiondata.csv", index=False,encoding='utf-8-sig')
    # writer = pd.ExcelWriter(mailreport+f"{today}_collectiondata.csv", engine="xlsxwriter")
    # maildata.to_csv(writer, index=False,encoding='utf-8-sig')
    #-------------------------------------------------------------------------------------------------------------------
    return final_df_TD,df_TD

##======================================================================================================================
def totaltax_collectionreport(std_path,infile,mappath,msterdatapath_,zonemap,TDCollection,mailreport):
    # Read ColumN naming mapping & stdandard report values
    namemap = pd.read_excel(mappath + "naming_map.xlsx")
    ytdreport_stdvalue = pd.read_excel(mappath + "reportformat.xlsx")

    #Column marathi to english dictionary
    colmapdict_martoeng = dict(zip(namemap['Marathi Name'], namemap['English Name']))

    ###-----------------------------------------------------------------------------------------------------------------
    # Renaming marathi columns to english
    ytdreport_rename = ytdreport_stdvalue.rename(columns=colmapdict_martoeng)
    ytdreport_rename['Zone'] = ytdreport_rename['Zone'].map(zonemap)
    ## Defined the existing dataframe in new ytdreport_df variable
    ytdreport_df =  ytdreport_rename.copy()

    ###-----------------------------------------------------------------------------------------------------------------
    ## Calculate grand total illegal construction,bloated demand & total demand
    ytdreport_df[f'total_demand/total'] = ytdreport_df['total_demand/arrears'] + ytdreport_df['total_demand/current']
    ytdreport_df[f'illegal_construction/total'] = ytdreport_df['illegal_construction/arrears'] + ytdreport_df['illegal_construction/current']
    ytdreport_df[f'bloated_demand/total'] = ytdreport_df['bloated_demand/arrears'] + ytdreport_df['bloated_demand/current']

    ###-----------------------------------------------------------------------------------------------------------------
    ## calculate the demand arrears & current or their grand total sum
    ytdreport_df['demand/arrears'] = ytdreport_df['total_demand/arrears'] - ytdreport_df['illegal_construction/arrears']
    ytdreport_df['demand/current'] = ytdreport_df['total_demand/current'] - ytdreport_df['illegal_construction/current']
    ytdreport_df[f'grand_total_demand'] = ytdreport_df['demand/arrears'] + ytdreport_df['demand/current']
    manual_ytdreportdf = ytdreport_df.copy()
###---------------------------------------------------------------------------------------------------------------------
    # Read YTD data
    # ytddata = pd.read_excel(infile, sheet_name="Total")
    ytddata = pd.read_excel(infile)
    ytddata['eng_zone'] = ytddata['ezname'].map(zonemap)
    sumgrpby = ytddata.groupby(['ezname', 'eng_zone']).agg({'magil': 'sum', 'chalu': 'sum'}).reset_index()
    sumgrpby['sum_of_magil_in_cr'] = round(sumgrpby['magil'] / 10000000, 2)
    sumgrpby['sum_of_chalu_in_cr'] = round(sumgrpby['chalu'] / 10000000, 2)

    lissy = ['Nigdi Pradhikaran', 'Akurdi', 'Chinchwad', 'Thergaon', 'Sangvi', 'Pimpri Waghere',
             'Pimpri Nagar', 'MNP Bhavan', 'Fugewadi Dapodi', 'Bhosari', 'Charholi',
             'Moshi', 'Chikhali', 'Talvade', 'Kivle', 'Dighi Bopkhel', 'Wakad']
    d = {v: k for k, v in enumerate(lissy)}
    df_YTD = sumgrpby.sort_values('eng_zone', key=lambda x: x.map(d), ignore_index=True)
    #Dump
    df_YTD.to_excel(mailreport+f"{today}_TotalTaxCollection.xlsx", index=False,encoding='utf-8-sig')
    ###-----------------------------------------------------------------------------------------------------------------
    manual_ytdreportdf[f'date_{tday}_recovery_magil'] = df_YTD['sum_of_magil_in_cr']
    manual_ytdreportdf[f'date_{tday}_recovery_chalu'] = df_YTD['sum_of_chalu_in_cr']
    manual_ytdreportdf[f'total_recovery_{tday}'] = df_YTD['sum_of_magil_in_cr'] + df_YTD['sum_of_chalu_in_cr']

    manual_ytdreportdf['percentage/arrears'] = \
        round((manual_ytdreportdf[f'date_{tday}_recovery_magil'] * 100) / manual_ytdreportdf['demand/arrears'], 2)
    manual_ytdreportdf['percentage/current'] = \
        round((manual_ytdreportdf[f'date_{tday}_recovery_chalu'] * 100) / manual_ytdreportdf['demand/current'], 2)
    manual_ytdreportdf['total_percentage'] = \
        round((manual_ytdreportdf[f'total_recovery_{tday}'] * 100) / manual_ytdreportdf['grand_total_demand'], 2)

    manual_ytdreportdf['percentage_of_objective'] = \
        round((manual_ytdreportdf[f'total_recovery_{tday}'] * 100) /
              manual_ytdreportdf['annual_objective'], 2)
    manual_ytdreportdf['balance_objective'] = \
        round((manual_ytdreportdf['annual_objective'] -
               manual_ytdreportdf[f'total_recovery_{tday}']),2)
    ###-----------------------------------------------------------------------------------------------------------------
    # Identify the pending days to the end of current financial year
    # last_year = today.year - 1
    financial_year_start = datetime.date(today.year, 4, 1) if today.month > 3 else datetime.date(today.year - 1, 4, 1)
    financial_year_end = datetime.date(financial_year_start.year + 1, 3, 31)
    ###-----------------------------------------------------------------------------------------------------------------
    # future = datetime.date(2023, 3, 31)
    future = financial_year_end
    diff = (future - today)
    diff_days = diff.days
    # str_yest = str(yesterday)
    tomw = today + timedelta(days=1)
    str_tomw = str(tomw)
    str_tday = str(today)
    ###-----------------------------------------------------------------------------------------------------------------
    # Next day objective
    yest_outpath = std_path + "Output/" + str(yesterday) + "/"

    if not os.path.exists(yest_outpath):
        for i in range(1, 7):
            yyy = yesterday - timedelta(days=i)
            y_outpath = std_path + "Output/" + str(yyy) + "/"
            if os.path.exists(y_outpath):
                tomorrow_objective = pd.read_excel(y_outpath + f"PCMC_PTAX_CollectionReport_{str(yyy)}.xlsx",
                                                   sheet_name="YTDCollection", skiprows=5, skipfooter=1)
                next_day = yyy + timedelta(days=1)
                str_nxtday = str(next_day)
                tomorrow_objective.dropna(subset=[f'{str_nxtday}_उद्द‍िष्ट'], how='all', inplace=True)
                tomorrow_objective = tomorrow_objective.reset_index(drop=True)
                manual_ytdreportdf['daily_objective'] = tomorrow_objective[f'{str_nxtday}_उद्द‍िष्ट']
                break
    else:
        tomorrow_objective = pd.read_excel(yest_outpath + f"PCMC_PTAX_CollectionReport_{str(yesterday)}.xlsx",
                                           sheet_name="YTDCollection", skiprows=5, skipfooter=1)
        tomorrow_objective.dropna(subset=[f'{str_tday}_उद्द‍िष्ट'], how='all', inplace=True)
        tomorrow_objective = tomorrow_objective.reset_index(drop=True)
        manual_ytdreportdf['daily_objective'] = tomorrow_objective[f'{str_tday}_उद्द‍िष्ट']
    ###-----------------------------------------------------------------------------------------------------------------
    manual_ytdreportdf['pending_days'] = diff_days
    manual_ytdreportdf[f'{str_tomw}_objective'] =\
        round((manual_ytdreportdf[f'balance_objective'] / manual_ytdreportdf['pending_days']), 2)

    TDCollection = TDCollection.reset_index(drop=True)
    manual_ytdreportdf['Recovery'] =  TDCollection['Recovery']
    ###-----------------------------------------------------------------------------------------------------------------
    # stdreport_1.to_excel("stdreport.xlsx",index=False)
    manual_ytdreportdf.index = manual_ytdreportdf.index + 1
    # manual_ytdreportdf.loc["grand_total"] = manual_ytdreportdf.sum(numeric_only=True)
    manual_ytdreportdf.loc["grand_total"] = manual_ytdreportdf[['Gat', 'Number of property', 'total_demand/arrears',
                                      'total_demand/current', 'illegal_construction/arrears',
                                      'illegal_construction/current', 'bloated_demand/arrears',
                                      'bloated_demand/current', 'annual_objective',
                                      'revised_annual_objectives', 'total_demand/total',
                                      'illegal_construction/total', 'bloated_demand/total', 'demand/arrears',
                                      'demand/current', 'grand_total_demand',
                                       f'date_{tday}_recovery_magil', f'date_{tday}_recovery_chalu',f'total_recovery_{tday}','balance_objective',
                                      'daily_objective', 'pending_days', f'{str_tomw}_objective', 'Recovery']].sum(numeric_only=True)

    recovery_magil_last_row = manual_ytdreportdf[f'date_{tday}_recovery_magil'].iloc[-1:][0]
    recovery_chalu_last_row = manual_ytdreportdf[f'date_{tday}_recovery_chalu'].iloc[-1:][0]
    total_recovery_last_row = manual_ytdreportdf[f'total_recovery_{tday}'].iloc[-1:][0]
    annual_objective_last_row = manual_ytdreportdf[f'annual_objective'].iloc[-1:][0]
    ###-----------------------------------------------------------------------------------------------------------------
    demand_arrears_last_row = manual_ytdreportdf[f'demand/arrears'].iloc[-1:][0]
    demand_current_last_row = manual_ytdreportdf[f'demand/current'].iloc[-1:][0]
    grand_total_demand_last_row = manual_ytdreportdf[f'grand_total_demand'].iloc[-1:][0]

    manual_ytdreportdf.loc["grand_total" ,'percentage_of_objective'] = round((total_recovery_last_row * 100) /
                                                                annual_objective_last_row, 2)
    manual_ytdreportdf.loc["grand_total", 'percentage/arrears'] = round((recovery_magil_last_row * 100) /
                                                            demand_arrears_last_row, 2)
    manual_ytdreportdf.loc["grand_total" ,'percentage/current'] = round((recovery_chalu_last_row * 100) /
                                                           demand_current_last_row, 2)
    manual_ytdreportdf.loc["grand_total" ,'total_percentage'] = round((total_recovery_last_row * 100) /
                                                          grand_total_demand_last_row, 2)
    manual_ytdreportdf.loc["grand_total", 'pending_days'] = ""
    ###-----------------------------------------------------------------------------------------------------------------
    stdreport_1 = manual_ytdreportdf.reset_index()
    zonetype =pd.read_csv(msterdatapath_ + "zone.csv")
    zonemap = dict(zip(zonetype['eng_zonename'],zonetype['zonename']))
    stdreport_1['Zone'] = stdreport_1['Zone'].map(zonemap)
    ###-----------------------------------------------------------------------------------------------------------------
    ###-----------------------------------------------------------------------------------------------------------------
    # Rearrange the columns as per standard report
    rearrange_ytddata_report = stdreport_1[['index', 'Zone', 'Gat', 'Number of property',
                                            'total_demand/arrears','total_demand/current','total_demand/total',
                                            'illegal_construction/arrears','illegal_construction/current','illegal_construction/total',
                                            'bloated_demand/arrears','bloated_demand/current','bloated_demand/total',
                                            'demand/arrears','demand/current', 'grand_total_demand', f'date_{tday}_recovery_magil',
                                            f'date_{tday}_recovery_chalu',f'total_recovery_{tday}',
                                            'percentage/arrears','percentage/current','total_percentage',
                                            'annual_objective', 'revised_annual_objectives',
                                            'percentage_of_objective', 'balance_objective','daily_objective',
                                            'pending_days','Recovery',f'{str_tomw}_objective']]
    ###-----------------------------------------------------------------------------------------------------------------
    # Renaming the columns name in standard format
    rearrange_ytddata_report_rename = rearrange_ytddata_report.rename(
        columns={f'date_{tday}_recovery_magil': "Arrears", f'date_{tday}_recovery_chalu': "Current",
                 f'total_recovery_{tday}': "Total", f'{str_tomw}_objective': f'{str_tomw}_उद्द‍िष्ट'})
    ###-----------------------------------------------------------------------------------------------------------------
    # final_stdreport_1 = stdreport_1.rename(
    #     columns={'index': 'अ.क्र.', 'Zone': 'विभागीय कार्यालय', 'Grand Total': 'एकूण', 'Recovery': 'वसूली'})
    # Replace the grand total name in the marathi format
    report_ytd = rearrange_ytddata_report_rename.replace("grand_total",'एकूण')
    ###-----------------------------------------------------------------------------------------------------------------
    #Finally,convert the columns name in standard format as per report(column name mentioned in col_map file)
    colmap = pd.read_excel(mappath + "col_map.xlsx")
    colmap_em = dict(zip(colmap['English Name'],colmap['Marathi Name']))
    final_df_YTD = report_ytd.rename(columns=colmap_em)
    ###-----------------------------------------------------------------------------------------------------------------
    return final_df_YTD,str_tomw
