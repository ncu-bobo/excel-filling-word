import xlrd
from docxtpl import DocxTemplate
def read_xls(file_path, sheet_name):
    # 打开Excel文件
    excelData = xlrd.open_workbook(filename=file_path)
    # 读取指定的工作表
    sheetData = excelData.sheet_by_name(sheet_name)
    # 统计有效行数范围
    startRow = 2
    endRow = 0
    preValue = 0
    nextIndex = 2
    # 获取合并单元格信息
    merged_cells = sheetData.merged_cells
    # 根据A列判断是否为有效行数
    for rowIndex in range(startRow, sheetData.nrows):
        if rowIndex != nextIndex:
            continue
        # 获取A列的值
        aCellValue = sheetData.cell_value(rowIndex, 0)
        # 检查当前单元格是否被合并，如果是，则获取当前合并的最后一行数，如果不是，则就是下一行
        is_merged = False
        for (rlow, rhigh, clow, chigh) in merged_cells:
            if rlow <= rowIndex < rhigh :
                is_merged = True
                nextIndex = rhigh
                break
        if is_merged == False:
            nextIndex = rowIndex + 1

        if isinstance(aCellValue, (int, float)):
            preValue = aCellValue
            endRow = nextIndex
            # 组装JSON数据并渲染
            docTemplate = DocxTemplate("./files/操作清单.docx")
            bCellValue = sheetData.cell_value(rowIndex, 1)
            jsonData = {
                "JOB_NO": bCellValue,
                "endRow": endRow,
                "preValue": preValue
            }

            docTemplate.render(jsonData)
            docTemplate.save(f"./files/操作清单_{aCellValue}.docx")

        else:
            break
    print(f"endRow= {endRow}, preValue= {preValue}")

    # 读取json数据


if __name__ == "__main__":
    read_xls("./files/2024年海运.xlsx", "9.13")