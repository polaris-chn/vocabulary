import xlwt
import os

txt1_path = './english_wordlists/CET6_edited.txt'
txt2_path = './english_wordlists/CET4_edited.txt'
txt3_path = './english_wordlists/TOEFL.txt'
save_name1 = 'CET6_edited.xls'
save_name2 = 'CET4_edited.xls'
save_name3 = 'TOEFL.xls'

def txt2excel(txt_path, save_name):
	word = []
	soundmark = []
	meaning = []
	with open(txt_path, 'r') as txt:
		lines = txt.readlines()
		for line in lines:
			if txt_path.find('6') != -1:
				word.append(line.split(' ')[0])
				meaning.append(line.split(' ')[2])
			else:
				word.append(line.split(' ', 1)[0])
				point1 = line.find('[')
				point2 = line.find(']')
				soundmark.append(line[point1:point2+1])
				if line.find(']') != -1:
					meaning.append(line[point2+1: ].strip())
				else:
					meaning.append(line.split(' ')[1].strip())

	txt.close()

	if os.path.exists(save_name):
		os.remove(save_name)

	book = xlwt.Workbook()
	sheet = book.add_sheet('sheet1')
	if len(soundmark) == 0:
		for n in range(len(word)):
			sheet.write(n,0,word[n])
			sheet.write(n,1,meaning[n])
	else:
		for n in range(len(word)):
			sheet.write(n,0,word[n])
			sheet.write(n,1,soundmark[n])
			sheet.write(n,2,meaning[n])
	book.save(save_name)

txt2excel(txt2_path, save_name2)
txt2excel(txt1_path, save_name1)
txt2excel(txt3_path, save_name3)