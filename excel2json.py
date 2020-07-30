import sys
reload(sys)
sys.setdefaultencoding('utf8')

from xlrd import open_workbook
import json

result = {}

with open_workbook('./answers.xlsx') as workbook:
  table_answers = workbook.sheet_by_name('answers')
  result['answers'] = []
  cells = table_answers.col_values(0);
  for cell in cells:
    result['answers'].append(cell);
  table_presets = workbook.sheet_by_name('presets')
  result['presets'] = []
  for col_index in range(table_presets.ncols):
    cells = table_presets.col_values(col_index);
    name = cells[2]
    question = cells[1]
    answers = []
    for index in range(len(cells)):
      if (index > 2 and len(cells[index]) > 0):
        answers.append(cells[index])
    qa = {'name': name, 'q': question, 'a': answers}
    result['presets'].append(qa)

result = json.dumps(result, ensure_ascii = False, encoding = 'utf-8')

print(result)

with open('assets.json','w') as assets:
  assets.write(str(result).decode('utf-8'))
  assets.close()
