import os
import sys
import pandas as pd
import datetime


def get_row_height(raw: str) -> int:
    height = 24
    # 1行
    if len(raw) <= 9:
        return height
    raw_lenght = len(raw)
    # 超过9个就换行
    break_count = 9
    row_count = int(raw_lenght / break_count)
    # 2行
    if row_count == 1:
        height += 7
        return height
    # 3行以上
    height = (row_count - 1) * 16

    if raw_lenght % break_count != 0:
        height += 16
    return height


def stamp(worksheet1, row_height_list: list, png_name):
    # 6 = 前面4行 + 合计 + 发货人
    page_height = 760
    page_count = 1
    total_height = 0
    break_list = []
    for i, height in enumerate(row_height_list, 1):
        total_height += height
        if total_height >= page_height * page_count:
            break_list.append(i - 1)
            page_count += 1

    # 如果最后一页只有一条数据，则把倒数第二条数据也放到最后一页
    rest_height = total_height - len(break_list) * page_height
    if rest_height - 60 <= 0:
        break_list.pop()
        break_list.append(len(row_height_list) - 2)

    worksheet1.insert_image(
        f'C2',
        png_name,
        {'x_scale': 4.2 / (4 * 4.62), 'y_scale': 3 / (4 * 3.14)},
    )
    for i in break_list:
        worksheet1.insert_image(
            f'C{i+1}',
            png_name,
            {'x_scale': 25 / 100, 'y_scale': 25 / 100, 'y_offset': 2},
        )
    worksheet1.set_h_pagebreaks(break_list)


def write_one(
    folder_path: str, date: datetime.datetime, df: pd.DataFrame, png_name: str
):
    datetime_object = datetime.datetime.strptime(date, '%Y-%m-%d')

    file_name = datetime_object.strftime('%Y-%m-%d')
    writer = pd.ExcelWriter(
        os.path.join(folder_path, f'{file_name}.xlsx'), engine='xlsxwriter'
    )
    workbook1 = writer.book
    worksheet1 = workbook1.add_worksheet('Sheet1')
    worksheet1.center_horizontally()

    fmt_row1 = workbook1.add_format(
        {
            "font_name": u"宋体",
            'font_size': 16,
            'bold': True,
            'align': 'centre',
            'valign': 'vcentre',
        }
    )
    fmt_row2_3_head = workbook1.add_format(
        {
            "font_name": u"宋体",
            'font_size': 12,
            'align': 'left',
            'valign': 'vcentre',
        }
    )
    fmt_row2_3 = workbook1.add_format(
        {
            "font_name": u"宋体",
            'font_size': 12,
            'align': 'centre',
            'valign': 'vcentre',
        }
    )
    fmt_row4 = workbook1.add_format(
        {
            "font_name": u"宋体",
            'font_size': 12,
            'bold': True,
            'align': 'centre',
            'valign': 'vcentre',
            'border': 1,
        }
    )
    fmt_row5 = workbook1.add_format(
        {
            "font_name": u"宋体",
            'font_size': 12,
            'align': 'centre',
            'valign': 'vcentre',
            'border': 1,
            'text_wrap': 1,
        }
    )
    fmt_row6 = workbook1.add_format(
        {
            "font_name": u"宋体",
            'font_size': 12,
            'align': 'centre',
            'valign': 'vcentre',
            'border': 1,
            'num_format': '#,##0.00',
        }
    )

    fmt_row_sum_head = workbook1.add_format(
        {
            "font_name": u"宋体",
            'font_size': 12,
            'bold': True,
            'align': 'centre',
            'valign': 'vcentre',
            'border': 1,
        }
    )

    fmt_row_sum_value = workbook1.add_format(
        {
            "font_name": u"宋体",
            'font_size': 12,
            'bold': True,
            'align': 'centre',
            'valign': 'vcentre',
            'border': 1,
            'num_format': '#,##0.00',
        }
    )

    fmt_row_last = workbook1.add_format(
        {
            "font_name": u"宋体",
            'font_size': 12,
            'align': 'left',
            'valign': 'vcentre',
        }
    )

    worksheet1.set_column('A:A', 10.67)
    worksheet1.set_column('B:B', 23)
    worksheet1.set_column('C:C', 8.89)
    worksheet1.set_column('D:D', 10)
    worksheet1.set_column('E:E', 10.44)
    worksheet1.set_column('F:F', 13)
    worksheet1.set_column('G:G', 7.33)

    # df = df.astype(str)

    s1 = '配送单'
    s2 = '下单日期：'
    s3 = '司机名称：李尧'
    s4 = '配送日期：'
    s5 = '司机电话：13876242672'

    d1 = (datetime_object + datetime.timedelta(days=-1)).strftime('%Y-%m-%d')
    d2 = datetime_object.strftime('%Y-%m-%d')

    columns = ['序号', '商品名', '单位', '数量', '单价', '下单金额', '备注']

    worksheet1.set_row(0, 35)
    worksheet1.merge_range(0, 0, 0, 6, s1, fmt_row1)

    worksheet1.set_row(1, 27)
    worksheet1.write(1, 0, s2, fmt_row2_3_head)
    worksheet1.write(1, 1, d1, fmt_row2_3)
    worksheet1.merge_range(1, 4, 1, 5, s3, fmt_row2_3_head)

    worksheet1.set_row(2, 27)
    worksheet1.write(2, 0, s4, fmt_row2_3_head)
    worksheet1.write(2, 1, d2, fmt_row2_3)
    worksheet1.merge_range(2, 4, 2, 6, s5, fmt_row2_3_head)

    worksheet1.set_row(3, 27)
    for index, v in enumerate(columns):
        worksheet1.write(3, index, v, fmt_row4)

    row_height_list = [35, 27, 27, 27]

    sum = 0
    df_index = 0
    for _, row in df.iterrows():
        index = df_index + 1
        row_index = df_index + 4
        df_index += 1

        height = get_row_height(row['B'])
        worksheet1.set_row(row_index, height)
        row_height_list.append(height)

        worksheet1.write(row_index, 0, index, fmt_row5)

        worksheet1.write(row_index, 1, row['B'], fmt_row5)
        worksheet1.write(row_index, 2, row['C'], fmt_row5)
        worksheet1.write(row_index, 3, row['D'], fmt_row6)
        worksheet1.write(row_index, 4, row['E'], fmt_row6)
        worksheet1.write(row_index, 5, row['F'], fmt_row6)
        worksheet1.write(row_index, 6, '', fmt_row5)

        sum += row['F']

    # 合计：倒数第二行
    data_row_len = len(df)
    row_index = data_row_len + 4
    worksheet1.set_row(row_index, 24)
    s6 = '合计'
    worksheet1.merge_range(row_index, 0, row_index, 4, s6, fmt_row_sum_head)
    worksheet1.write(row_index, 5, sum, fmt_row_sum_value)
    worksheet1.write(row_index, 6, '', fmt_row_sum_value)

    # 收货人：最后一行
    row_index += 1
    worksheet1.set_row(row_index, 60)
    s7 = '     收货人：                                            监督人：'
    worksheet1.merge_range(row_index, 0, row_index, 6, s7, fmt_row_last)

    row_height_list.extend([24, 60])
    # 打印图章
    stamp(worksheet1, row_height_list, png_name)

    writer.close()
    pass


def read_all(xlsx_path: str, sheet_name_list: list) -> dict[pd.DataFrame]:
    df = pd.read_excel(
        xlsx_path,
        names=['M', 'B', 'C', 'D', 'E', 'F', 'G'],
        dtype={"M": "str"},
        sheet_name=sheet_name_list,
    )
    return df

    pass


def write_all(folder_path, xlsx_path, sheet_name_list, png_name):
    df1 = read_all(xlsx_path, sheet_name_list)
    list = []
    for i, df2 in df1.items():
        list.append(df2)
    df3 = pd.concat(list)

    df3['M'] = df3['M'].apply(lambda x: x.replace(' 00:00:00', ''))
    date_list = df3['M'].values
    date_list.sort()

    date_set = set(date_list)
    for v in date_set:
        if v != '2023-05-30':
            continue
        df4 = df3.query('M == @v')
        write_one(folder_path, v, df4, png_name)


if __name__ == '__main__':
    sheet_name_list = [
        '14964',
        '16477',
        '17360',
        '16039',
        '9839',
        '19867',
        '17674',
        '14340',
    ]
    xlsx_path = 'zj.xlsx'
    folder_path = 'data'
    png_name = 'shian.png'
    if not os.path.isdir(folder_path):
        os.makedirs(folder_path)
    write_all(folder_path, xlsx_path, sheet_name_list, png_name)
    pass
