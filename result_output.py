import logging

from xlrd import open_workbook

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)-8s: %(message)s")


def read_excel(path):
    """
    读取表格实验数据信息直接过滤掉从7500导出excle的前六行信息
    """
    data_fam = []
    data_vic = []
    data_rox = []
    data_total = []

    excel = open_workbook(path)
    data = excel.sheet_by_index(0)

    rows = data.nrows
    cols = data.ncols

    logging.info(f"读取表格信息：\t{path},\t{rows}\t行,\t{cols}\t列")

    for i in range(rows):
        if i > 6:
            data_value = data.row_values(i)

            if "FAM" in data_value:
                data_fam.append(data_value)

            if "VIC" in data_value:
                data_vic.append(data_value)

            if "ROX" in data_value:
                data_rox.append(data_value)

    for i, j, k in zip(data_fam, data_vic, data_rox):
        _data = i, j, k
        # logging.info(_data)

        data_total.append(_data)

    return data_total


def save_csv(data, path):
    """
    判断分型结果后csv保存
    """
    logging.info(f"此次实验总共有 {len(data) + 1} 个样本")

    file = open(path, mode="w+")  # 如若出现乱码可指定编码，常用 encoding="gb2312",encoding="utf-8"
    file.write(
        'Result' + "," + 'Well' + "," + 'Sample Name' + "," + 'Target Name' + "," + 'Reporter' + "," + 'Ct' + "," + 'Target Name' + "," + 'Reporter' + "," + 'Ct' + "," + 'Target Name' + "," + 'Reporter' + "," + 'Ct' + "\n")

    for i, j in enumerate(data):
        if j[0][0] == j[1][0] == j[2][0] and j[0][1] == j[1][1] == j[2][1]:
            well, sample_name = j[0][0], j[0][1]
            fam_target, fam_reporter, fam_ct = j[0][2], j[0][4], j[0][6]
            vic_target, vic_reporter, vic_ct = j[1][2], j[1][4], j[1][6]
            rox_target, rox_reporter, rox_ct = j[2][2], j[2][4], j[2][6]

            # print(well,sample_name,fam_target,fam_reporter,fam_ct,vic_target,vic_reporter,vic_ct,rox_target,rox_reporter,rox_ct)

            if sample_name == "NTC":
                if str(fam_ct) == "Undetermined" and str(vic_ct) == "Undetermined" and str(rox_ct) == "Undetermined":
                    file.write("," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")
                else:
                    file.write(
                        "NTC  异常" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                            fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                            vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")
            elif sample_name == "P":
                if float(fam_ct) <= 30 and float(vic_ct) <= 30 and float(rox_ct) <= 30:
                    file.write("," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")
                else:
                    file.write(
                        "阳性对照  异常" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                            fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                            vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")

            elif float(rox_ct) <= 32:
                if type(fam_ct) is float:
                    if type(vic_ct) is float:
                        if fam_ct <= 36 and vic_ct > 36:
                            if fam_target == "285FAM":
                                file.write(
                                    "*2  G/G   纯合野生" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")
                            if fam_target == "893FAM2":
                                file.write(
                                    "*3  G/G   纯合野生" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")
                            if fam_target == "560FAM2":
                                file.write(
                                    "*17  C/C   纯合野生" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")

                        if fam_ct <= 36 and vic_ct <= 36:
                            if fam_target == "285FAM":
                                file.write(
                                    "*2   G/A   杂合突变" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")
                            if fam_target == "893FAM2":
                                file.write(
                                    "*3   G/A   杂合突变" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")
                            if fam_target == "560FAM2":
                                file.write(
                                    "*17  C/T   杂合突变" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")

                        if fam_ct > 36 and vic_ct <= 36:
                            if fam_target == "285FAM":
                                file.write(
                                    "*2   A/A   纯合突变" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")
                            if fam_target == "893FAM2":
                                file.write(
                                    "*3   A/A   纯合突变" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")
                            if fam_target == "560FAM2":
                                file.write(
                                    "*17  T/T   纯合突变" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")

                    if type(vic_ct) is str:
                        if type(fam_ct) is float:
                            if fam_ct <= 36 and vic_ct == "Undetermined":
                                if fam_target == "285FAM":
                                    file.write(
                                        "*2   G/G   纯合野生" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                            fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                            vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")
                                if fam_target == "893FAM2":
                                    file.write(
                                        "*3   G/G   纯合野生" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                            fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                            vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")
                                if fam_target == "560FAM2":
                                    file.write(
                                        "*17  C/C   纯合野生" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                            fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                            vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")

                if type(fam_ct) is str:
                    if type(vic_ct) is float:
                        if vic_ct <= 36 and fam_ct == "Undetermined":
                            if fam_target == "285FAM":
                                file.write(
                                    "*2   A/A   纯合突变" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")
                            if fam_target == "893FAM2":
                                file.write(
                                    "*3   A/A   纯合突变" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")
                            if fam_target == "560FAM2":
                                file.write(
                                    "*17  T/T   纯合突变" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")

            elif float(rox_ct) > 32:
                file.write(
                    "ROX 异常" + "," + well + "," + sample_name + "," + fam_target + "," + fam_reporter + "," + str(
                        fam_ct) + "," + vic_target + "," + vic_reporter + "," + str(
                        vic_ct) + "," + rox_target + "," + rox_reporter + "," + str(rox_ct) + "\n")

            else:
                logging.error("has some  error! ")
                break

    file.close()


if __name__ == "__main__":
    excel_path = "20191213-285-893-560-test_sample1_32_data.xls"
    csv_path = f"{excel_path[:-4]}" + "_output.csv"

    data_total = read_excel(excel_path)
    save_csv(data_total, csv_path)
