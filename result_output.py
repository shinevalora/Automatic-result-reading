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

    # 表格中共有多少行多少列数据信息
    rows = data.nrows
    cols = data.ncols

    logging.info(f"读取表格信息：\t{path},\t{rows}\t行,\t{cols}\t列")

    for i in range(rows):
        # 过滤前六行表头信息
        if i > 6:
            data_value = data.row_values(i)

            if "FAM" in data_value:
                data_fam.append(data_value)

            if "VIC" in data_value:
                data_vic.append(data_value)

            if "ROX" in data_value:
                data_rox.append(data_value)
    # 为每一个样本的所有检测位点都归为一个小列表
    for i, j, k in zip(data_fam, data_vic, data_rox):
        _data = i, j, k
        # logging.info(_data)

        data_total.append(_data)

    return data_total


def save_csv(data, path):
    """
    判断分型结果后csv保存
    """
    logging.info(f"此次实验总共有 {len(data)} 个样本")
    field_name = [
        'Result',
        'Well',
        'Sample Name',
        'Target Name',
        'Reporter',
        'Ct',
        'Target Name',
        'Reporter',
        'Ct',
        'Target Name',
        'Reporter',
        'Ct'
    ]

    file = open(path, mode="w+")  # 如若出现乱码可指定编码，常用 encoding="gb2312",encoding="utf-8"
    file.write(",".join(field_name) + "\n")

    count = 0
    fam_target_map = {
        "fam_ct <= 36 < vic_ct": {
            "285FAM": "*2    G/G   纯合野生",
            "893FAM2": "*3   G/G   纯合野生",
            "560FAM2": "*17  C/C   纯合野生",
        },
        "fam_ct <= 36 and vic_ct <= 36": {
            "285FAM": "*2    G/A   杂合突变",
            "893FAM2": "*3   G/A   杂合突变",
            "560FAM2": "*17  C/T   杂合突变",
        },
        "fam_ct > 36 >= vic_ct": {
            "285FAM": "*2    A/A   纯合突变",
            "893FAM2": "*3   A/A   纯合突变",
            "560FAM2": "*17  T/T   纯合突变",
        },
        "fam_ct <= 36 and vic_ct ==  'Undetermined'": {
            "285FAM": "*2    G/G  纯合野生",
            "893FAM2": "*3   G/G  纯合野生",
            "560FAM2": "*17  C/C  纯合野生",
        },
        "vic_ct <= 36 and fam_ct == 'Undetermined'": {
            "285FAM": "*2    A/A   纯合突变",
            "893FAM2": "*3   A/A   纯合突变",
            "560FAM2": "*17  T/T   纯合突变",
        },
    }

    for item in data:
        # 样本孔号和样本名字信息需一一对等的才可比较进行；
        if item[0][0] == item[1][0] == item[2][0] and item[0][1] == item[1][1] == item[2][1]:
            count += 1
            well, sample_name = item[0][0], item[0][1]
            fam_target, fam_reporter, fam_ct = item[0][2], item[0][4], item[0][6]
            vic_target, vic_reporter, vic_ct = item[1][2], item[1][4], item[1][6]
            rox_target, rox_reporter, rox_ct = item[2][2], item[2][4], item[2][6]

            _data = [
                well, sample_name, fam_target, fam_reporter, str(fam_ct), vic_target, vic_reporter, str(vic_ct),
                rox_target, rox_reporter, str(rox_ct)
            ]

            _data = [str(i) for i in _data]

            logging.debug(_data)

            # logging.info(f"fam_target is {fam_target}")

            result = ""

            if sample_name == "NTC":
                if str(fam_ct) == "Undetermined" and str(vic_ct) == "Undetermined" and str(rox_ct) == "Undetermined":
                    result = ""
                else:
                    result = "NTC  异常"
            elif sample_name == "P":

                if float(fam_ct) <= 30 and float(vic_ct) <= 30 and float(rox_ct) <= 30:
                    result = ""
                else:
                    result = "阳性对照  异常"
            elif float(rox_ct) <= 32:
                if type(fam_ct) is float:
                    if type(vic_ct) is float:
                        if fam_ct <= 36 < vic_ct:
                            # 映射target类别及类型
                            result = fam_target_map["fam_ct <= 36 < vic_ct"][fam_target]

                        if fam_ct <= 36 and vic_ct <= 36:
                            result = fam_target_map["fam_ct <= 36 and vic_ct <= 36"][fam_target]

                        if fam_ct > 36 >= vic_ct:
                            result = fam_target_map["fam_ct > 36 >= vic_ct"][fam_target]

                    if type(vic_ct) is str:
                        if type(fam_ct) is float:
                            if fam_ct <= 36 and vic_ct == "Undetermined":
                                result = fam_target_map["fam_ct <= 36 and vic_ct ==  'Undetermined'"][fam_target]

                if type(fam_ct) is str:
                    if type(vic_ct) is float:
                        if vic_ct <= 36 and fam_ct == "Undetermined":
                            result = fam_target_map["vic_ct <= 36 and fam_ct == 'Undetermined'"][fam_target]

            elif float(rox_ct) > 32:
                result = "ROX 异常"

            else:
                logging.error("has some  error! ")
                break
            file.write(f"{result},{','.join(_data)}\n")
        else:
            logging.warning(f"异常数据: {item}")

    file.close()
    logging.info(f"完成! 处理条数: {count}")


if __name__ == "__main__":
    excel_path = "20191213-285-893-560-test_sample1_32_data.xls"
    csv_path = f"{excel_path[:-4]}" + "_output.csv"

    data_total = read_excel(excel_path)
    save_csv(data_total, csv_path)
