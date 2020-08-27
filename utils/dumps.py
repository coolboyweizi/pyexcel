import xlwt


class dumpStyle(xlwt.XFStyle):
    """
    初始化表格样式
    """

    def font(self, name="宋体", bold=1, height=20):
        _font = xlwt.Font()  # 为样式创建字体
        _font.name = name  # 'Times New Roman'
        _font.bold = bold
        _font.color_index = 4
        _font.height = height
        self.font = _font

    def setBorders(self, left=1, right=1, top=1, bottom=1):
        _borders = xlwt.Borders()
        _borders.left = left
        _borders.right = right
        _borders.top = top
        _borders.bottom = bottom
        self.borders = _borders
