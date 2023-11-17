import xlrd
from pyecharts.charts import Map
from pyecharts import options as opts
import datetime

now = datetime.datetime.now()
fn = now.strftime('%Y年%m月%d日')

xls = xlrd.open_workbook("data.xls")
table = xls.sheet_by_name("Sheet4")
data = []
m = 0
for n in range(4, 38):
    if table.cell(n, 1).value > 0:
        data.append((table.cell(n, 0).value, int(table.cell(n, 1).value)))
        m = max(m, table.cell(n, 1).value)

MyMap = Map(init_opts=opts.InitOpts(
    page_title="个人国内足迹图",
    width="1000px",
    height='640px',
    bg_color="white",
    animation_opts=opts.AnimationOpts(animation_easing="cubicOut"),
))

MyMap.set_global_opts(
    title_opts=opts.TitleOpts(
        title="个人国内足迹",
        subtitle=f"截止日期：{fn}  ",
        pos_right="center",
        pos_top="5%"),
    visualmap_opts=opts.VisualMapOpts(
        pos_right="center",
        pos_top="15%",
        orient="horizontal",
        is_piecewise=False,
        min_=1,
        max_=m,
        range_color=["lightyellow", "red"]),
    legend_opts=opts.LegendOpts(
        is_show=False))

MyMap.add(
    tooltip_opts=opts.TooltipOpts(
        is_show=True,
    ),
    is_map_symbol_show=True,
    label_opts=opts.LabelOpts(
        is_show=True
    ),
    series_name="到访地区数",
    data_pair=data,
    maptype="china",
    is_roam=False)

MyMap.render("Trvel.html")