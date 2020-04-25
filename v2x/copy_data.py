import xlrd


def read_file(file_url):
    try:
        data = xlrd.open_workbook(file_url)
        sheets = data.sheets()

        print('Excel文件 %s 包含以下数据表:' % file_url, end=' ')
        for sheet in sheets:
            print(sheet.name, end=" ")
        return data
    except Exception as e:
        print(str(e))


def process(workbook, sheet_name, from_index, to_index, file_path, x_column, y_column, z_column):
    table = workbook.sheet_by_name(sheet_name)  # 获得表格
    total_rows = table.nrows  # 拿到总共行数
    print('\n表格总行数为%d.\n' % total_rows)

    x_rows = table.col_values(ord(x_column) - ord('A'))  # x值所在的列数（例如, J列为第9列）
    y_rows = table.col_values(ord(y_column) - ord('A'))  # y所在的列数（例如, K列为第10列）
    z_rows = table.col_values(ord(z_column) - ord('A'))  # z所在的列数（例如, K列为第10列）
    row = from_index - 1

    with open(file_path, 'w+') as f:
        for i in range(1, to_index - from_index + 2):
            head = 'pair' + str(i) + ':\n'
            point2d_x = '  point2d_x: ' + str(111) + '\n'
            point2d_y = '  point2d_y: ' + str(111) + '\n'
            point3d_x = '  point3d_x: ' + str(x_rows[row]).strip() + '\n'
            point3d_y = '  point3d_y: ' + str(y_rows[row]).strip() + '\n'
            point3d_z = '  point3d_z: ' + str(z_rows[row]).strip() + '\n'
            f.write(head)
            f.write(point2d_x)
            f.write(point2d_y)
            f.write(point3d_x)
            f.write(point3d_y)
            f.write(point3d_z)
            row += 1


sheet_name = '摄像头安装信息'

excel_path = "/home/jcglqmoyx/Desktop/111.xlsx"  # 要处理的Excel文件路径
file = read_file(excel_path)  # 读取Excel文件
from_index = 291
to_index = 299
# 生成的.yaml文件的路径以及名称

pole = 6
camera = 100
file_path = '/home/jcglqmoyx/Desktop/' + str(pole) + '_' + str(camera) + '_pairs.yaml'

# 处理该Excel文件的“摄像头安装信息”工作表, 输出对第from_index到to_index(两端均包含）行数据进行处理后的结果
process(workbook=file, sheet_name=sheet_name, from_index=from_index, to_index=to_index, file_path=file_path,
        x_column='P', y_column='Q', z_column='L')
