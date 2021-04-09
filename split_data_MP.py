import pandas as pd
import os
from tqdm import tqdm
import numpy as np
from datetime import datetime
import locale


# Import data skip 3 row pertama dan menjadikan row kesatu sbg header
xl = pd.ExcelFile('resources/1. Marketplace Parcel.xlsx')
banyaksheet = len(xl.sheet_names)

for i in range(banyaksheet):
    exec("data{0}=pd.read_excel(xl, sheet_name={0}, skiprows=3, header=1)".format(i))
    for y in eval("data{}.columns".format(i)):
        if(eval("data{0}[y].dtype == np.float64".format(i))):
            exec("data{0}[y] = np.float32(data{0}[y])".format(i))
        else:
              continue
    exec("df{0} = data{0}[['DP ID', 'Nama Mitra']]".format(i))
    exec("df{0} = df{0}.drop_duplicates().reset_index(drop=True)".format(i))

# Mengambil path file dan membuat path hasil split data
path = os.getcwd()
result_path = path + "/results/"

# Fungsi untuk autofit
def get_col_widths(dataframe):
    # Menentukan panjang maksimum dari kolom index
    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
    # Menggabungkan maksimum dari panjang kolom dan panjang data dari kolom tersebut
    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]


for i in tqdm(range(len(df0))):
    direct = result_path + str(df0['DP ID'][i]) + " " + df0['Nama Mitra'][i]
    if not os.path.exists(direct):
    	os.makedirs(direct)
    writer = pd.ExcelWriter(direct + "/" + str(df0['DP ID'][i]) + " "+"Marketplace"+" " + df0['Nama Mitra'][i] + ".xlsx")
    if df1['DP ID'].astype(str).str.contains(str(df0.iloc[i]['DP ID'])).any() == True:
        temp = data0[data0['DP ID'] == df0['DP ID'][i]].reset_index(drop=True)
        temp1 = data1[data1['DP ID'] == df0['DP ID'][i]].reset_index(drop=True)
        #ini buat edit kolom dari sheetnya
        #temp buat sheet1
        #temp1 buat sheet2
        #parameter dari temp itu nama kolom dari sheet
        temp['Create Time'] = temp['Create Time'].apply(lambda x: datetime.strptime(str(x),'%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y') if pd.notna(x)==True else None)
        temp1['Create Time'] = temp1['Create Time'].apply(lambda x: datetime.strptime(str(x),'%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y') if pd.notna(x)==True else None)
        temp1['Commission'] = temp1['Commission'].apply(np.round)
        temp1['Delivery Fee'] = temp1['Delivery Fee'].apply(np.round)
        #hapus kalo ada duplikasi
        temp = temp.drop(['DP ID', 'Nama Mitra'], axis=1)
        temp1 = temp1.drop(['DP ID', 'Nama Mitra'], axis=1)
        #ngebuat sheet1
        temp.to_excel(writer, sheet_name="Marketplace", startrow=4, index=False)
        #ngebuat sheet2
        temp1.to_excel(writer, sheet_name="Marketplace Claim", startrow=4, index=False)
    else:
        temp = data0[data0['DP ID'] == df0['DP ID'][i]].reset_index(drop=True)
        temp['Create Time'] = temp['Create Time'].apply(lambda x: datetime.strptime(str(x),'%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y') if pd.notna(x)==True else None)
        temp = temp.drop(['DP ID', 'Nama Mitra'], axis=1)
        temp.to_excel(writer, sheet_name="Marketplace", startrow=4, index=False)

    worksheet = writer.sheets["Marketplace"]
    workbook  = writer.book

    # Buat format tabel excelnya
    bawah = workbook.add_format({'top': 2})
    header_format = workbook.add_format({'bottom': 2, 'bold': True})
    tebal = workbook.add_format({'bold': True})
    center = workbook.add_format({'align': 'center'})
    bulatkan = workbook.add_format({'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'})
    centerheader = workbook.add_format({'bottom': 2, 'bold': True, 'align': 'center'})
    penanggalan = workbook.add_format({'num_format': 'D MMMM YYYY', 'align': 'center'})

    worksheet.set_column('E:G', None, center)
    #for buat ngambil value untuk sheet1
    for col_num, value in enumerate(temp.columns.values):
    	worksheet.write(4, col_num, value, header_format)
    	worksheet.write_blank(len(temp)+5, col_num, '', bawah)
    #for buat bikin format baris dan kolom untuk sheet1
    for j in range(5):
        worksheet.write(4, 4+j, temp.columns[4+j], centerheader)
    worksheet.set_column('G:P', None, bulatkan)
    for j, width in enumerate(get_col_widths(temp)):
        worksheet.set_column(j-1, j-1, width+1.5)
    worksheet.write(0, 0, "PT. Andiarta Muzizat", tebal)
    worksheet.write(1, 0, "Marketplace "+" Parcel Shipment Report Period 2021", tebal)
    worksheet.write(2, 0, "Nama Mitra : "+str(df0['DP ID'][i]) + " s- " + df0['Nama Mitra'][i], tebal)

    if df1['DP ID'].astype(str).str.contains(str(df0.iloc[i]['DP ID'])).any() == True:
        worksheet1 = writer.sheets["Marketplace Claim"]
        worksheet1.set_column('E:I', None, center)
        #for buat ngambil value untuk sheet2
        for col_num, value in enumerate(temp1.columns.values):
            worksheet1.write(4, col_num, value, header_format)
            worksheet1.write_blank(len(temp1)+5, col_num, '', bawah)
        #for buat bikin format baris dan kolom untuk sheet2
        for j in range(5):
            worksheet1.write(4, j, temp1.columns[j], centerheader)
        worksheet1.set_column('J:P', None, bulatkan)
        for j, width in enumerate(get_col_widths(temp1)):
            worksheet1.set_column(j-1, j-1, width+1.5)
        worksheet1.write(0, 0, "PT. Andiarta Muzizat", tebal)
        worksheet1.write(1, 0, "Marketplace Parcel Claim", tebal)
        worksheet1.write(2, 0, "Nama Mitra : "+str(df0['DP ID'][i]) + " - " + df0['Nama Mitra'][i], tebal)
    
    #kalau sudah pakai save(), gk perlu pakai close()    
    writer.save()


print("Done!")
print("Eldho Kece")
