from utils.loads import loadExcel
from utils.dumps import dumpStyle
import xlwt
# 导入的文本
excel_file = 'data/demo.xlsx'

# 对象
excel_obj = loadExcel(excel_file, 3)

dump_file = 'data/dump.xlsx'

f = xlwt.Workbook()
sheet1 = f.add_sheet(u'统计',cell_overwrite_ok=True)

# 设置sheet1的row和column


# 自定义过滤函数
def tryInt(x):
    try:
        value = x.value
        x = int(value)
    except Exception as e:
        x = e
    return isinstance(x, int)


# 调用相关的方法
excel_obj.filter(53, tryInt) \
    .filter('超时时长', tryInt) \
    .filter('超时时长', lambda v: int(v.value) > 8) \
    .filter('骑手ID', lambda id: int(id.value) == 18535529) \
    .fields("骑手ID,商家名称,骑手,超时时长")

for key, value in enumerate(excel_obj.getColIndex()):
    # 设置样式
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = '微软雅黑'
    font.bold = True
    font.color_index = 10
    font.height = 200
    style.font = font

    borders = xlwt.Borders()
    borders.left = 100
    borders.right = 100
    borders.top = 10
    borders.bottom = 10

    align = xlwt.Alignment()
    align.HORZ_CENTER=True

    style.alignment = align

    # 设置列宽
    sheet1.col(key).width=256*20  #  256为衡量单位，20表示20个字符宽度
    sheet1.col(key).center=True
    sheet1.write(0, key, value, style)



for row, line in enumerate(excel_obj):
    for col, data in enumerate(line):
        sheet1.write(row+1, col, data.value)

f.save('dump.xlsx')