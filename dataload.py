from pyecharts.charts import Line
import pyecharts.options as opts
import xlrd
import datetime
import re

def date_as_tuple(raw_date):
    '''
    用xlrd模块读取日期格式的单元格
    :param raw_date 原始数据,list类型，如：[52234.0, 798798.0]
    :return: 日期列表 date_list
    '''
    date_list = []
    for date in raw_date:
        new_date = datetime.datetime(*xlrd.xldate_as_tuple(date, 0)).strftime('%Y-%m-%d')
        date_list.append(new_date+"-06:00")
        date_list.append(new_date + "-10:00")
        date_list.append(new_date + "-16:00")
        date_list.append(new_date + "-20:00")
    # print(date_list)
    return date_list

def check_data_format(data):
    pattern = re.compile('[0-9]+')
    return pattern.findall(data)

def split_data_format(data):
    return data.split('/')

# 打开存储数据的excel
data = xlrd.open_workbook('1.xlsx')

# 以表格的形式取出数据
table = data.sheets()[0]

# 取出表格中第二列数据
fied = table.col_values(0)[1:]
x = date_as_tuple(fied)
high_list = []
low_list = []
dance_list = []

#总行数
rownumber =table.nrows
for index in range(1,rownumber):
    rowdate =table.row_values(index)[1:]
    # print(rowdate)
    for a, data in enumerate(rowdate):
       if check_data_format(data):
           mid_data = split_data_format(data)
           high_list.append(mid_data[0])
           low_list.append(mid_data[1])
           dance_list.append(mid_data[2])
       else:
           high_list.append('')
           low_list.append('')
           dance_list.append('')
           # print(split_data_format(data))




# 生成一个折线统计图对象
line=(
    Line(init_opts=opts.InitOpts(width="1600px", height="800px"))
    .add_xaxis(xaxis_data=x)
    .add_yaxis(
        series_name="收缩压",
        y_axis=high_list,
        markpoint_opts=opts.MarkPointOpts(
            data=[
                opts.MarkPointItem(type_="max", name="最大值"),
                opts.MarkPointItem(type_="min", name="最小值"),
            ]
        ),
        markline_opts=opts.MarkLineOpts(
            data=[opts.MarkLineItem(type_="average", name="平均值")]
        ),
        is_smooth=True,
        is_connect_nones=True
    )
    .add_yaxis(
        series_name="舒张压",
        y_axis=low_list,
        markpoint_opts=opts.MarkPointOpts(
            data=[
                opts.MarkPointItem(type_="max", name="最大值"),
                opts.MarkPointItem(type_="min", name="最小值"),
            ]
        ),
        markline_opts=opts.MarkLineOpts(
            data=[
                opts.MarkLineItem(type_="average", name="平均值"),
                opts.MarkLineItem(symbol="none", x="90%", y="max"),
            ]
        ),
        is_smooth=True,
        is_connect_nones=True
    )
    .add_yaxis(
        series_name="脉搏",
        y_axis=dance_list,
        markpoint_opts=opts.MarkPointOpts(
            data=[
                opts.MarkPointItem(type_="max", name="最大值"),
                opts.MarkPointItem(type_="min", name="最小值"),
            ]
        ),
        markline_opts=opts.MarkLineOpts(
            data=[
                opts.MarkLineItem(type_="average", name="平均值"),
                opts.MarkLineItem(symbol="none", x="90%", y="max"),
            ]
        ),
        is_smooth=True,
        is_connect_nones=True
    )
    .set_global_opts(
        title_opts=opts.TitleOpts(title="毛昊罡", subtitle="血压可视化自动生成"),
        tooltip_opts=opts.TooltipOpts(trigger="axis"),
        toolbox_opts=opts.ToolboxOpts(is_show=True),
        xaxis_opts=opts.AxisOpts(type_="category", boundary_gap=False),
        datazoom_opts=[
            opts.DataZoomOpts(yaxis_index=0),
            opts.DataZoomOpts(type_="inside", yaxis_index=0),
        ],
    )
)


# 渲染到html页面
line.render('./index.html')
