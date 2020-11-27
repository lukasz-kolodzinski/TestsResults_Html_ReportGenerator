import xlrd
from typing import List

excel_path = "./Xlsx/results_2020-11-23.xlsx"

input_excel = xlrd.open_workbook(excel_path)
input_excel = input_excel.sheet_by_index(2)

rows_ammount = input_excel.nrows

def get_values_from_row(row_id: int) -> List[str]:
    """
    Function that gathers specific values from particular row
    :param row_id:
    :return:
    """
    out = []
    for clmn in range(20):
        if clmn in range(8,19):
            continue
        value = input_excel.cell_value(row_id, clmn)
        if clmn == 0 and type(value) is not str:
            out.append(str(int(value)))  # float return
        else:
            out.append(str(value))
    return out

headers = get_values_from_row(0)
results_of_tests = []
for itr in range(1, rows_ammount):
    results = input_excel.cell_value(itr,19)
    if results in ["Fail", "NP"]:
        results_of_tests.append(get_values_from_row(itr))

def gen_report(header: List, rows: List[List]) -> None:
    template = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>FAILED/NOT_PERFORMED TESTS STATUS</title>
</head>
<body>
        <table style="width: 90%;" border="2" cellspacing="2" cellpadding="4">>
            <tbody>
                <tr style="border-color: black; background-color: moccasin; text-align: center;">
    """
    for title in header:
        template += "<td>{}</td>\n".format(title)
    template += """<td>Comment</td>
                </>"""
    for row in rows:
        template += "<tr>"
        for column in row:
            template += "<td>{}</td>\n".format(column)
        template += """<td>
        <textarea name="comment" rows="3" cols="">
  </textarea>
        </td></tr>"""
    template += """</tbody>
        </>
</body>
</html>"""
    print(template)

gen_report(headers, results_of_tests)






