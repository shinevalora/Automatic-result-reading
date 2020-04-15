# _*_coding:utf-8_*_
#  作者    : shinevalora
#  创建时间: 2019/12/30  10:00


import xlrd

data_total = []
data_fam = []
data_vic = []
data_rox = []


# excel_name='./20191213-285-893-560-test_sample1_32_data.xls'


def read_excel(excel_name):
    workbook = xlrd.open_workbook(excel_name, encoding_override="utf-8")
    # print(workbook.encoding)

    # 根据sheet索引或者名称获取sheet内容
    sheet1 = workbook.sheet_by_name('Results')
    # sheet的名称，行数，列数
    nrows = sheet1.nrows
    ncols = sheet1.ncols
    # print(sheet1.name, nrows, ncols)

    for row in range(nrows):
        if row <= 7:
            continue
        data = sheet1.row_values(row)
        # print(data)

        if 'FAM' in data:
            data_fam.append(data)

        if 'VIC' in data:
            data_vic.append(data)

        if 'ROX' in data:
            data_rox.append(data)

    for i in range(96):

        if data_fam[i][6] == "Undetermined":
            data_fam[i][6] = float(data_fam[i][6].replace("Undetermined", "40"))

        if data_vic[i][6] == "Undetermined":
            data_vic[i][6] = float(data_vic[i][6].replace("Undetermined", "40"))

        if data_rox[i][6] == "Undetermined":
            data_rox[i][6] = float(data_rox[i][6].replace("Undetermined", "40"))

        data_ = data_fam[i], data_vic[i], data_rox[i]

        data_total.append(data_)

    csv_name = excel_name.split(".xls")[0] + "结果判断.csv"

    with open(csv_name, 'w+') as file:
        file.write("结果判断"
                   + "," + "Well" + "," + "Sample Name" + "," + "FAM Target Name" + "," + "Task" + "," + "Reporter" + "," + "Quencher" + "," + "Cт" + "," + "Cт  Mean"
                   + "," + "Cт  SD" + "," + "Quantity" + "," + "Quantity Mean" + "," + "Quantity SD" + "," + "Automatic Ct Threshold" + "," + "Ct Threshold"
                   + "," + "Automatic Baseline" + "," + "Baseline Start" + "," + "Baseline End" + "," + "Comments" + "," + "HIGHSD"
                   + "," + "Well" + "," + "Sample Name" + "," + "VIC Target Name" + "," + "Task" + "," + "Reporter" + "," + "Quencher" + "," + "Cт" + "," + "Cт  Mean"
                   + "," + "Cт  SD" + "," + "Quantity" + "," + "Quantity Mean" + "," + "Quantity SD" + "," + "Automatic Ct Threshold" + "," + "Ct Threshold"
                   + "," + "Automatic Baseline" + "," + "Baseline Start" + "," + "Baseline End" + "," + "Comments" + "," + "HIGHSD"
                   + "," + "Well" + "," + "Sample Name" + "," + "ROX Target Name" + "," + "Task" + "," + "Reporter" + "," + "Quencher" + "," + "Cт" + "," + "Cт  Mean"
                   + "," + "Cт  SD" + "," + "Quantity" + "," + "Quantity Mean" + "," + "Quantity SD" + "," + "Automatic Ct Threshold" + "," + "Ct Threshold"
                   + "," + "Automatic Baseline" + "," + "Baseline Start" + "," + "Baseline End" + "," + "Comments" + "," + "HIGHSD" + "\n")

        for i in data_total:
            # print(i)
            fam, fam_ct, vic, vic_ct, rox, rox_ct = i[0], i[0][6], i[1], i[1][6], i[2], i[2][6]

            if 'NTC' in fam:
                # print('NTC  不做判断',i)
                file.write("NTC  不做判断" + "\t" + "," + str(i[0]) + "," + str(i[1]) + "," + str(i[2]) + "\n")

            # 285FAM
            if fam_ct <= 36 and vic_ct > 36 and '285FAM' in fam:
                # print("CYP2C19*2  G/G 纯合野生",i)
                file.write("CYP2C19*2  G/G 纯合野生" + "\t" + "," + str(i[0]) + "," + str(i[1]) + "," + str(i[2]) + "\n")
            if fam_ct <= 36 and vic_ct <= 36 and '285FAM' in fam:
                # print("CYP2C19*2  G/A 杂合突变",i)
                file.write("CYP2C19*2  G/A 杂合突变" + "\t" + "," + str(i[0]) + "," + str(i[1]) + "," + str(i[2]) + "\n")
            if fam_ct > 36 and vic_ct <= 36 and '285FAM' in fam:
                # print("CYP2C19*2  A/A 纯合突变",i)
                file.write("CYP2C19*2  A/A 纯合突变" + "\t" + "," + str(i[0]) + "," + str(i[1]) + "," + str(i[2]) + "\n")

            # 893FAM2
            if fam_ct <= 36 and vic_ct > 36 and '893FAM2' in fam:
                # print("CYP2C19*3  G/G 纯合野生",i)
                file.write("CYP2C19*3  G/G 纯合野生" + "\t" + "," + str(i[0]) + "," + str(i[1]) + "," + str(i[2]) + "\n")
            if fam_ct <= 36 and vic_ct <= 36 and '893FAM2' in fam:
                # print("CYP2C19*3  G/A 杂合突变",i)
                file.write("CYP2C19*3  G/A 杂合突变" + "\t" + "," + str(i[0]) + "," + str(i[1]) + "," + str(i[2]) + "\n")
            if fam_ct > 36 and vic_ct <= 36 and '893FAM2' in fam:
                # print("CYP2C19*3  A/A 纯合突变",i)
                file.write("CYP2C19*3  A/A 纯合突变" + "\t" + "," + str(i[0]) + "," + str(i[1]) + "," + str(i[2]) + "\n")

            # 560FAM2

            if fam_ct <= 36 and vic_ct > 36 and '560FAM2' in fam:
                # print("CYP2C19*17  C/C 纯合野生",i)
                file.write("CYP2C19*17  C/C 纯合野生" + "\t" + "," + str(i[0]) + "," + str(i[1]) + "," + str(i[2]) + "\n")
            if fam_ct <= 36 and vic_ct <= 36 and '560FAM2' in fam:
                # print("CYP2C19*17  C/T 杂合突变",i)
                file.write("CYP2C19*17  C/T 杂合突变" + "\t" + "," + str(i[0]) + "," + str(i[1]) + "," + str(i[2]) + "\n")
            if fam_ct > 36 and vic_ct <= 36 and '560FAM2' in fam:
                # print("CYP2C19*17  T/T 纯合突变",i)
                file.write("CYP2C19*17  T/T 纯合突变" + "\t" + "," + str(i[0]) + "," + str(i[1]) + "," + str(i[2]) + "\n")

# read_excel(excel_name)
