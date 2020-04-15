
import xlrd
from xlrd import open_workbook


excel_path="./20191213-285-893-560-test_sample1_32_data.xls"


def read_excel():

	excel=open_workbook(excel_path)
	data=excel.sheet_by_index(0)

	rows=data.nrows
	cols=data.ncols

	# print(rows,cols)

	for i in range(rows):
		if i > 6:
			data_value=data.row_values(i)

			if "FAM" in data_value:
				data_fam.append(data_value)

			if "VIC" in data_value:
				data_vic.append(data_value)

			if "ROX" in data_value:
				data_rox.append(data_value)

	return data_fam,data_vic,data_rox


def save_csv():

	file=open(csv_path,encoding="gb2312",mode="w+")
	file.write('Result'+","+'Well'+","+'Sample Name'+","+'Target Name'+","+'Reporter'+","+'Ct'+","+'Target Name'+","+'Reporter'+","+'Ct'+","+'Target Name'+","+'Reporter'+","+'Ct'+"\n")

	for i ,j ,k in zip(data_fam,data_vic,data_rox):

		if i[0]==j[0]==k[0] and i[1]==j[1]==k[1]:
			if i[1]=="NTC":
				if  str(i[6])=="Undetermined" and str(j[6])=="Undetermined" and str(k[6])=="Undetermined":		
					file.write(","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")
				else:
					file.write("NTC  异常"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")

			elif i[1]=="P":
				if float(i[6])<=30 and float(j[6])<=30 and float(k[6])<=30:
					file.write(","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")
				else:
					file.write("阳性对照  异常"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")

			elif float(k[6])<=32:
				if type(i[6]) is float:
					if type(j[6]) is float:
						if i[6]<=36 and j[6]>36:
							if i[2]=="285FAM":
								file.write("*2  G/G   纯合野生"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")					
							if i[2]=="893FAM2":
								file.write("*3  G/G   纯合野生"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")
							if i[2]=="560FAM2":
								file.write("*17  C/C   纯合野生"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")

						if i[6]<=36 and j[6]<=36:
							if i[2]=="285FAM":
								file.write("*2   G/A   杂合突变"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")					
							if i[2]=="893FAM2":
								file.write("*3   G/A   杂合突变"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")
							if i[2]=="560FAM2":
								file.write("*17  C/T   杂合突变"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")

						if i[6]>36 and j[6]<=36:
							if i[2]=="285FAM":
								file.write("*2   A/A   纯合突变"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")
							if i[2]=="893FAM2":
								file.write("*3   A/A   纯合突变"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")
							if i[2]=="560FAM2":
								file.write("*17  T/T   纯合突变"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")
	
					if type(j[6]) is str:
						if type(i[6]) is float :
							if i[6]<=36 and j[6]=="Undetermined":
								if i[2]=="285FAM":
									file.write("*2   G/G   纯合野生"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")
								if i[2]=="893FAM2":
									file.write("*3   G/G   纯合野生"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")
								if i[2]=="560FAM2":
									file.write("*17  C/C   纯合野生"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")

				if type(i[6]) is str:
					if type(j[6]) is float:
						if j[6]<=36 and i[6]=="Undetermined":
							if i[2]=="285FAM":
								file.write("*2   A/A   纯合突变"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")
							if i[2]=="893FAM2":
								file.write("*3   A/A   纯合突变"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")
							if i[2]=="560FAM2":
								file.write("*17  T/T   纯合突变"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")
			
			elif float(k[6])>32:
				file.write("ROX 异常"+","+i[0]+","+i[1]+","+i[2]+","+i[4]+","+str(i[6])+","+j[2]+","+j[4]+","+str(j[6])+","+k[2]+","+k[4]+","+str(k[6])+"\n")


	file.close()


if __name__=="__main__":

	data_fam=[]
	data_vic=[]
	data_rox=[]

	read_excel()

	csv_path=f"{excel_path[:-4]}"+"_output.csv"
	
	save_csv()





