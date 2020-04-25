# 依赖：xlrd 库用于处理Excel表格, pyproj库用来进行地理坐标转换
import xlrd
from pyproj import Proj


# 将纬度由度分秒的格式转化为度（将分和秒转换为小数点部分）
def lat_to_x(lat):
    '''
    :param lat: 从Excel表格中读取的纬度数据（度分秒格式), 例如 38:16:45.99632N
    :return: 度（有小数点）
    '''
    if lat[-1] == 'N':  # N 如果读取的纬度数据末尾有一个字母N, 则将字母N去掉, 经处理后, 输入的纬度数据变为38:16:45.99632
        lat = lat[:-1]
    lat = lat.strip()
    degree = int(lat[0:2])  # 取纬度数据的前两个字符为纬度整数部分（本例中为38）
    minute = int(lat[3:5])  # 取纬度数据的第3和第4个字母为 分值（本例中为16）
    second = float(lat[6:len(lat)])  # 取纬度数据的第7个数据直到末尾为秒值（本例中为45.99632）
    return degree + minute / 60 + second / 3600  # 纬度值由度分秒格式转化为小数点格式, 转换过后的纬度值为38.27944342222222


# 将经度由度分秒的格式转化为度（将分和秒转换为小数点部分）
def lon_to_y(lon):
    if lon[-1] == 'E':
        lon = lon[:-1]
    lon = lon.strip()
    degree = int(lon[0:3])
    minute = int(lon[4:6])
    second = float(lon[7:len(lon)])
    return degree + minute / 60 + second / 3600


def geographic_to_UTM(lon, lat):
    '''
    调用Python库, 将经纬度坐标转换为UTM坐标
    :param lon: 经度值
    :param lat: 纬度值
    :return: UTM坐标
    '''
    # zone值为在网页中输入经纬度值之后产生的zone值（沧州地区的V2X设备所处的zone值均为50, 所以这里默认写成50
    p = Proj(proj='utm', zone=50, ellps='WGS84', preserve_units=False)
    x, y = p(lon, lat)
    return x, y


def read_file(file_url):
    '''
    读取Excel表格
    :param file_url: Excel表格的路径
    :return:
    '''
    try:
        data = xlrd.open_workbook(file_url)
        sheets = data.sheets()

        print('Excel文件 %s 包含以下数据表:' % file_url, end=' ')
        for sheet in sheets:
            print(sheet.name, end=" ")
        return data
    except Exception as e:
        print(str(e))


def process(workbook, sheet_name, fromIndex, toIndex, lat_column, lon_column):
    '''
    读取Excel表中数据, 进行处理后, 在控制台输出结果
    :param fromIndex: 从第excel表第几行开始处理数据
    :param toIndex: 处理到哪一行
    :return:
    '''
    table = workbook.sheet_by_name(sheet_name)  # 获得表格

    total_rows = table.nrows  # 拿到总共行数
    print('\n表格总行数为%d.\n' % total_rows)

    lat_rows = table.col_values(ord(lat_column) - ord('A'))  # LAT值所在的列数（例如, J列为第9列）
    lon_rows = table.col_values(ord(lon_column) - ord('A'))  # LON值所在的列数（例如, K列为第10列）
    row = fromIndex - 1
    while row <= toIndex - 1:
        x, y = geographic_to_UTM(lon=lon_to_y(lon_rows[row]), lat=lat_to_x(lat_rows[row]))
        # print(lon_rows[row])
        print('行数=%4d, %20s%20s' % (row + 1, x, y))
        row += 1


# 使用示例
file_path = "/home/jcglqmoyx/Desktop/road_8.xlsx"  # 要处理的Excel文件路径
file = read_file(file_path)  # 读取Excel文件
# 处理该Excel文件的“摄像头安装信息”工作表, 输出对第98, 99, 100行数据进行处理后的结果
process(workbook=file, sheet_name='摄像头安装信息', fromIndex=287, toIndex=305, lat_column='J', lon_column='K')
