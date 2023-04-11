import pandas as pd
import datetime
import os
import xlsxwriter
import warnings
warnings.filterwarnings('ignore')
import send_email as se

today = datetime.datetime.today().date()
tday  =today.strftime("%d/%m/%Y")


def excelwriter(outpth,logopath,TDCollection, YTD_RecoveryCollection,str_tomw):
    ###==================================================================================================================
    writer = pd.ExcelWriter(outpth + f"PCMC_PTAX_CollectionReport_{str(today)}.xlsx", engine="xlsxwriter")

    ##------------------------------------------------------------------------------------------------
    ## Preparing first page
    label1 = pd.DataFrame([f"पिंपरी चिंचवड महानगरपालिका, पिंपरी - 411 018"])
    label2 = pd.DataFrame([f"कर आकारणी व कर संकलन विभाग"])
    label3 = pd.DataFrame([f"दिनांक {tday} रोजीचा गट निहाय वसूली तक्ता"])
    label4 = pd.DataFrame([f"सन 2022 - 2023 विभागीय कार्यालयनिहाय मागणी तक्ता"])
    ##-----------------------------------------------------------------------------------------------------------
    # Today's collection report process by Gat & zone wise
    label1.to_excel(writer, sheet_name='TodaysCollection', startrow=1, startcol=8, index=False, header=False)
    label2.to_excel(writer, sheet_name='TodaysCollection', startrow=2, startcol=8, index=False, header=False)
    label3.to_excel(writer, sheet_name='TodaysCollection', startrow=3, startcol=8, index=False, header=False)
    TDCollection.to_excel(writer, sheet_name='TodaysCollection', startrow=5, startcol=0, index=False)
    ## select sheet name
    workbook = writer.book
    worksheet = writer.sheets['TodaysCollection']
    ### Hide grid lines
    worksheet.hide_gridlines(2)
    worksheet.freeze_panes(5, 2)

    ### Add Format
    border_format = workbook.add_format({'border': 1,
                                         'align': 'left',
                                         'font_color': '#000000',
                                         'font_size': 20})
    worksheet.conditional_format('A6:V24', {'type': 'cell',
                                            'criteria': '>=',
                                            'value': 0,
                                            'format': border_format})
    ## Add Bold format
    bold_format = workbook.add_format({'bold': True,
                                       'text_wrap': True,
                                       'align': 'center',
                                       'valign': 'center',
                                       'bg_color': '#FFFFFF',
                                       'font_color': '#333333',
                                       'font_size': 30})

    ## merge cells format
    merge_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'})  ##'fg_color': 'yellow'
    woutborder_merge_format = workbook.add_format({
        'bold': 1,
        'align': 'center',
        'valign': 'vcenter'})

    ##------------------------------------------------------------------------------------------------------------------
    ## Main header line bold & modify
    worksheet.conditional_format('H1:M4', {'type': 'no_blanks', 'format': bold_format})
    worksheet.conditional_format('A24:V24', {'type': 'no_blanks', 'format': bold_format})
    ##Insert logo
    worksheet.insert_image('G1', logopath + "PCMC_logo.png",
                           {'x_offset': 20, 'y_offset': 0, 'x_scale': 0.15, 'y_scale': 0.13})
    ## Set Columns
    worksheet.set_column('B1:B1', 15)
    worksheet.set_column('U6:U6', 10)
    worksheet.set_column('C5:T6', 8)
    worksheet.set_column('A1:A1', 4)
    ## merge row
    worksheet.merge_range('A24:B24', 'एकूण',merge_format)
    ## Set rows & their size
    for i in range(6, 23):
        worksheet.set_row(i, 25)

    ##------------------------------------------------------------------------------------------------------------------
    # YTD collection recovery report process by zone wise
    label1.to_excel(writer, sheet_name='YTDCollection', startrow=1, startcol=15, index=False, header=False)
    label2.to_excel(writer, sheet_name='YTDCollection', startrow=2, startcol=15, index=False, header=False)
    label4.to_excel(writer, sheet_name='YTDCollection', startrow=3, startcol=15, index=False, header=False)

    YTD_RecoveryCollection.to_excel(writer, sheet_name='YTDCollection', startrow=6, startcol=0, index=False)

    ## select sheet name
    # workbook = writer.book
    worksheet_ytd = writer.sheets['YTDCollection']
    ### Hide grid lines
    worksheet_ytd.hide_gridlines(2)
    worksheet_ytd.freeze_panes(5, 2)
    ##Insert logo
    worksheet_ytd.insert_image('G1', logopath + "PCMC_logo.png",
                           {'x_offset': 20, 'y_offset': 0, 'x_scale': 0.15, 'y_scale': 0.13})
    ## Border mention range cell
    worksheet_ytd.conditional_format('A6:AD25', {'type': 'cell',
                                            'criteria': '>=',
                                            'value': 0,
                                            'format': border_format})
    worksheet_ytd.conditional_format('P1:T4', {'type': 'no_blanks', 'format': bold_format})
    worksheet_ytd.conditional_format('A25:AD25', {'type': 'no_blanks', 'format': bold_format})

    ## Set Columns with width
    worksheet_ytd.set_column('B1:B1', 15)
    worksheet_ytd.set_column('D1:D1', 13)
    worksheet_ytd.set_column('A1:A1', 4)

    worksheet_ytd.set_column('E5:W6', 10)
    worksheet_ytd.set_column('N6:P6', 8)
    worksheet_ytd.set_column('T6:V6', 8)

    ### Merge columns
    # Cell format 1
    cell_format1 = workbook.add_format()
    # cell_format1.set_bg_color('#96b4cf')  # 609a9a
    cell_format1.set_align('center')
    cell_format1.set_bold(True)
    cell_format1.set_border()
    cell_format1.set_text_wrap()
    cell_format1.set_font_size(10)
    #### Cell format 2
    cell_format2 = workbook.add_format()
    # cell_format1.set_bg_color('#96b4cf')  # 609a9a
    cell_format2.set_align('center')
    cell_format2.set_bold(True)
    cell_format2.set_border()
    cell_format2.set_text_wrap()
    cell_format2.set_font_size(7.5)

    merge_col = ["एकूण मागणी","अवैध बांधकाम शास्ती","फुगीर मागणी","मागणी",f"दिनांक {tday} अखेर वसूली","टक्केवारी"]
    cell_range = ['E6:G6','H6:J6','K6:M6','N6:P6','Q6:S6','T6:V6']
    dict_range = dict(zip(merge_col,cell_range))
    for col_lst, colcell_range in dict_range.items():
        worksheet_ytd.merge_range(colcell_range, col_lst, cell_format1)

####----------------------------------------------------------------------------------------------------------------
    col_list = ['अ.क्र.', 'विभागीय कार्यालय', 'गट संख्या', 'मालमत्ता संख्या', 'वार्षिक उद्दिष्ट',
                'उद्द‍िष्ट टक्केवारी', 'शिल्लक उद्द‍िष्ट', 'दैनंदिन उद्द‍िष्ट', 'वसूली', f'{str_tomw}_उद्द‍िष्ट']
    alphabet_cell = ["A", "B", "C", "D", "W", "Y", "Z", "AA", "AC", "AD"]
    col_list2 = ['सुधारित वार्षिक उद्द‍िष्ट', 'आर्थिकवर्षास बाकीदिवस']
    alphabet_cell2 = ["X", "AB"]
    lst_dict = dict(zip(col_list,alphabet_cell))
    lst_dict2 = dict(zip(col_list2,alphabet_cell2))
    for i,j in lst_dict.items():
            worksheet_ytd.merge_range(f'{j}6:{j}7', i, cell_format1)
    for k,l in lst_dict2.items():
            worksheet_ytd.merge_range(f'{l}6:{l}7', k, cell_format2)

    worksheet_ytd.merge_range('A25:B25', 'एकूण', merge_format)
    worksheet_ytd.merge_range('AB5:AD5', 'रक्कम रुपये कोटीमध्ये', woutborder_merge_format)

    ## Set rows & their size
    for i in range(7, 25):
        worksheet_ytd.set_row(i, 22)

    ### Hide Columns
    worksheet_ytd.set_column("E5:M5", None, None, {'hidden': 1})
    worksheet_ytd.set_column("X5:X5", None, None, {'hidden': 1})
    #-------------------------------------------------------
    writer.save()
    writer.close()
    print(f"{tday} Report Prepared Successfully\n============================="
          "---------------------------------------------------------------------------------------------------")

    # se.send()
    # print(f"{tday} Report mailed Successfully\n============================="
    #       "---------------------------------------------------------------------------------------------------")
##==============================================================================================================================================


    # worksheet_ytd.merge_range('E6:G6',"एकूण मागणी",cell_format1)
    # worksheet_ytd.merge_range('H6:J6',"अवैध बांधकाम शास्ती",cell_format1)
    # worksheet_ytd.merge_range('K6:M6',"फुगीर मागणी",cell_format1)
    # worksheet_ytd.merge_range('N6:P6',"मागणी",cell_format1)
    # worksheet_ytd.merge_range('Q6:S6',f"दिनांक {tday} अखेर वसूली ",cell_format1)
    # worksheet_ytd.merge_range('T6:V6',"टक्केवारी",cell_format1)
