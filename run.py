from utils.loads import loadXlrd

# 导入的文本
excel_file = 'data/demo.xlsx'

# 对象
excel_obj = loadXlrd(excel_file, 3)


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
print(excel_obj)
for data in excel_obj:
    print(data)
