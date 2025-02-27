#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PDF报告生成模块
用于生成iTop运维服务月报的PDF格式报告。
"""

from io import BytesIO
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.charts.legends import Legend
from reportlab.lib.colors import HexColor
import pandas as pd
import os

# 注册中文字体
font_path = os.path.join(os.path.dirname(__file__), "simkai.ttf")
pdfmetrics.registerFont(TTFont('SimKai', font_path))

def _create_pdf_pie_chart(data, labels, title):
    """创建PDF饼图
    
    Args:
        data (list): 饼图数据列表,每个元素为一个数值
        labels (list): 饼图标签列表,每个元素为一个字符串,与data一一对应
        title (str): 饼图标题
        
    Returns:
        Drawing: reportlab Drawing对象,包含完整的饼图
        
    Example:
        >>> data = [60, 30, 10] 
        >>> labels = ['已解决', '未解决', '已关闭']
        >>> title = '工单状态分布'
        >>> pie_chart = _create_pdf_pie_chart(data, labels, title)
    """
    drawing = Drawing(400, 250)
    pie = Pie()
    pie.x = 100
    pie.y = 25
    pie.width = 200
    pie.height = 200
    pie.data = data
    pie.labels = labels
    
    # 设置样式
    pie.slices.strokeWidth = 0.5
    pie.sideLabels = True
    pie.sideLabelsOffset = 0.1
    pie.simpleLabels = False
    pie.slices.fontName = 'SimKai'
    
    # 计算百分比标签
    total = sum(pie.data)
    pie.labels = ['%.1f%%' % (value/total*100) for value in pie.data]

    # 设置颜色
    colors = [HexColor('#00b8a9'), HexColor('#f6416c'), HexColor('#f8f3d4')]
    for i, color in enumerate(colors[:len(data)]):
        pie.slices[i].fillColor = color

    drawing.add(pie)

    # 添加标题
    title_label = String(200, 250, title)
    title_label.fontName = 'SimKai'
    title_label.fontSize = 12
    title_label.textAnchor = 'middle'
    drawing.add(title_label)

    # 添加图例
    legend = Legend()
    legend.x = 320
    legend.y = 150
    legend.deltay = 15
    legend.fontSize = 10
    legend.fontName = 'SimKai'
    legend.alignment = 'right'
    legend.columnMaximum = 8
    legend.colorNamePairs = list(zip(colors[:len(data)], labels))
    drawing.add(legend)

    return drawing

def _create_pdf_line_chart(x_data, y_data, line_types):
    """创建PDF折线图
    
    Args:
        x_data (list): x轴数据列表,通常为月份
        y_data (list): y轴数据列表,每个元素为一条线的数据元组
        line_types (list): 每条线的名称列表,用于图例显示
        
    Returns:
        Drawing: 返回一个包含折线图的Drawing对象
        
    示例:
        x_data = ['1月', '2月', '3月']
        y_data = [(80, 85, 90), (70, 75, 80)]  # 两条线的数据
        line_types = ['团队A', '团队B']
        drawing = _create_pdf_line_chart(x_data, y_data, line_types)
    """
    drawing = Drawing(600, 300)
    lp = HorizontalLineChart()
    
    # 设置图表位置和大小
    lp.x = 10
    lp.y = 30
    lp.height = 200
    lp.width = 450
    
    # 设置数据
    lp.data = y_data
    lp.joinedLines = 1
    
    # 设置x轴
    catNames = [str(item) for item in x_data]
    lp.categoryAxis.categoryNames = catNames
    lp.categoryAxis.labels.fontName = 'SimKai'
    lp.categoryAxis.labels.boxAnchor = 'n'
    
    # 设置y轴
    lp.valueAxis.valueMin = 0
    lp.valueAxis.valueMax = 105
    lp.valueAxis.valueStep = 10
    
    # 设置线条样式
    line_colors = ['orange', 'grey', 'coral', 'salmon', 'lightcoral', 'tomato', 
                  'darkorange', 'goldenrod', 'khaki', 'lightgoldenrodyellow']
    
    for i in range(len(line_types)):
        lp.lines[i].strokeWidth = 2 if i == 0 else 1.5
        lp.lines[i].strokeColor = getattr(colors, line_colors[i])
        lp.lines[i].name = line_types[i]

    # 设置数据标签
    lp.lineLabelFormat = 'values'
    label_array = [[str(item) + '%' for item in line if item is not None] for line in y_data]
    lp.lineLabelArray = label_array

    drawing.add(lp)
    return drawing

def _create_table_style():
    """创建通用的表格样式"""
    return TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, -1), 'SimKai'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('WORDWRAP', (0, 0), (-1, -1), True),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ])

def _create_data_table(data, normal_style, col_widths=None):
    """创建数据表格
    Args:
        data: 表格数据(包含表头)
        normal_style: 文本样式，可选项包括:
            - Title: 标题样式
            - Heading1: 一级标题样式
            - Heading2: 二级标题样式
            - Heading3: 三级标题样式
            - Heading4: 四级标题样式
            - Heading5: 五级标题样式
            - Heading6: 六级标题样式
            - Normal: 正文样式
            - Italic: 斜体样式
            - Code: 代码样式
            - UnorderedList: 无序列表样式
            - OrderedList: 有序列表样式
            - Definition: 定义样式
            - BodyText: 正文文本样式
        col_widths: 列宽列表
    Returns:
        Table: 格式化的表格对象
    """
    if col_widths is None:
        # 计算表格宽度为页面宽度的85%
        table_width = letter[0] * 0.85
        col_widths = [table_width / len(data[0])] * len(data[0])
    
    table = Table(data, colWidths=col_widths)
    table.setStyle(_create_table_style())
    
    # 设置单元格自动换行
    for i, row in enumerate(data):
        for j, cell in enumerate(row):
            if isinstance(cell, bytes):
                table._cellvalues[i][j] = Paragraph(cell.decode(), normal_style)
            else:
                table._cellvalues[i][j] = Paragraph(str(cell), normal_style)
    
    return table

def _add_service_request_stats(elements, user_request_stats, incident_stats, change_stats, normal_style, subtitle_style):
    """添加服务请求统计"""
    elements.append(Paragraph("1) 服务请求统计", subtitle_style))
    if not user_request_stats.empty:
        total = user_request_stats['total'].iloc[0]
        resolved = user_request_stats['resolved_total'].iloc[0]
        closed = user_request_stats['closed_total'].iloc[0]
        unresolved = user_request_stats['unresolved_total'].iloc[0]
        
        if total and total > 0:
            resolved_percentage = resolved / total * 100
            closed_percentage = closed / resolved * 100 if resolved > 0 else 0
            unresolved_percentage = unresolved / total * 100
            
            elements.append(Paragraph(
                f"本周期内共接收服务请求 {total:g} 个，其中 {resolved:g} 个服务请求被解决，"
                f"占比约 {resolved_percentage:.2f}%；", normal_style))
            elements.append(Paragraph(
                f"已解决的服务请求中，{closed:g} 个服务请求被按时关闭，"
                f"占比约 {closed_percentage:.2f}%；", normal_style))
            elements.append(Paragraph(
                f"未解决的服务请求有 {unresolved:g} 个，"
                f"占比约 {unresolved_percentage:.2f}%。", normal_style))
            
            elements.append(Spacer(1, 12))
            elements.append(Spacer(1, 12))

            # 添加饼图
            pie_data = [resolved, unresolved, closed]
            pie_labels = ['已解决', '未解决', '已关闭']
            elements.append(_create_pdf_pie_chart(pie_data, pie_labels, "服务请求状态分布"))
        else:
            elements.append(Paragraph("本周期内没有接收到服务请求。", normal_style))
    else:
        elements.append(Paragraph("无法获取服务请求统计数据。", normal_style))
    elements.append(Spacer(1, 12))

    # 1.2 事件统计
    elements.append(Paragraph("2) 事件统计", subtitle_style))
    if not incident_stats.empty:
        total = incident_stats['total'].iloc[0]
        resolved = incident_stats['resolved_total'].iloc[0]
        closed = incident_stats['closed_total'].iloc[0]
        unresolved = incident_stats['unresolved_total'].iloc[0]

        if total and total > 0:
            resolved_percentage = resolved / total * 100
            closed_percentage = closed / resolved * 100 if resolved > 0 else 0
            unresolved_percentage = unresolved / total * 100

            elements.append(Paragraph(f"本周期内共发生事件 {total:g} 个，其中 {resolved:g} 个事件被解决，占比约 {resolved_percentage:.2f}%；", normal_style))
            elements.append(Paragraph(f"已解决的事件中，{closed:g} 个事件被按时关闭，占比约 {closed_percentage:.2f}%；", normal_style))
            elements.append(Paragraph(f"未解决的事件有 {unresolved:g} 个，占比约 {unresolved_percentage:.2f}%。", normal_style))
            
            # 添加一行空行
            elements.append(Spacer(1, 12)) 
            
            # 添加饼图
            pie_data = [resolved, unresolved, closed]
            pie_labels = ['已解决', '未解决', '已关闭']
            elements.append(_create_pdf_pie_chart(pie_data, pie_labels, "事件状态分布"))
        else:
            elements.append(Paragraph("本周期内没有发生事件。", normal_style))
    else:
        elements.append(Paragraph("无法获取事件统计数据。", normal_style))
    elements.append(Spacer(1, 12))

    # 1.3 变更统计
    elements.append(Paragraph("3) 变更统计", subtitle_style))
    if not change_stats.empty:
        total = change_stats['total'].iloc[0]
        resolved = change_stats['resolved_total'].iloc[0]
        closed = change_stats['closed_total'].iloc[0]

        if total and total > 0:
            closed_percentage = closed / total * 100
            resolved_percentage = resolved / closed * 100 if closed > 0 else 0

            elements.append(Paragraph(f"本周期内共发生变更 {total:g} 个，其中 {closed:g} 个变更已关闭，占比约 {closed_percentage:.2f}%。", normal_style))
            elements.append(Paragraph(f"已关闭的变更中，{resolved:g} 个变更被成功执行，占比约 {resolved_percentage:.2f}%。", normal_style))
            
            # 添加一行空行
            elements.append(Spacer(1, 12)) 
            
            # 添加饼图
            pie_data = [resolved, total-resolved]
            pie_labels = ['已解决', '未解决']
            elements.append(_create_pdf_pie_chart(pie_data, pie_labels, "变更状态分布"))
        else:
            elements.append(Paragraph("本周期内没有发生变更。", normal_style))
    else:
        elements.append(Paragraph("无法获取变更统计数据。", normal_style))
    elements.append(Spacer(1, 12))

def _add_team_stats(elements, team_stats, normal_style, subtitle_style):
    """添加团队统计"""
    if not team_stats.empty:
        team_data = [team_stats.columns.tolist()] + team_stats.values.tolist()
        elements.append(_create_data_table(team_data, normal_style))

        # 添加团队服务请求解决率趋势图
        df = pd.DataFrame(team_stats)
        df['工单解决率'] = df['工单解决率'].apply(lambda x: float(str(x).rstrip('%')))
        service_request_df = df[df['工单类型'] == '服务请求']

        x_data = sorted(set(service_request_df['月份']))
        y_data = []
        line_types = []
        for team in set(service_request_df['团队']):
            line_types.append(team)
            rates = []
            for month in x_data:
                rate = service_request_df[
                    (service_request_df['团队'] == team) & 
                    (service_request_df['月份'] == month)
                ]['工单解决率'].iloc[0] if len(service_request_df[
                    (service_request_df['团队'] == team) & 
                    (service_request_df['月份'] == month)
                ]) > 0 else None
                rates.append(float(rate) if rate is not None else None)
            y_data.append(tuple(rates))

        drawing = _create_pdf_line_chart(x_data, y_data, line_types)
        elements.append(Paragraph("各团队服务请求月度解决率趋势", subtitle_style))
        elements.append(drawing)
    else:
        elements.append(Paragraph("本周期内没有要处理的工单", normal_style))
    elements.append(Spacer(1, 12))

def _add_kpi_stats(elements, kpi_stats, title, normal_style, subtitle_style):
    """添加KPI统计"""
    elements.append(Paragraph(title, subtitle_style))
    if not kpi_stats.empty:
        # 添加KPI数据表格
        kpi_data = [kpi_stats.columns.tolist()] + kpi_stats.values.tolist()
        elements.append(_create_data_table(kpi_data, normal_style))

        # 添加KPI趋势图
        df = pd.DataFrame(kpi_stats)
        df['KPI总计'] = df['KPI总计'].apply(lambda x: float(str(x).rstrip('%')))
        data = [tuple(df['KPI总计'])]
        drawing = _create_pdf_line_chart(df['月份'].tolist(), data, ['KPI总计'])
        elements.append(Paragraph("KPI总计月度趋势", subtitle_style))
        elements.append(drawing)
    else:
        elements.append(Paragraph("本周期内没有KPI统计数据。", normal_style))
    elements.append(Spacer(1, 12))

def generate_pdf(start_date, end_date, ticket_summary, user_request_stats, incident_stats, 
                change_stats, team_stats, person_stats, unresolved_tickets, overdue_tickets, 
                infra_kpi_stats, app_kpi_stats):
    """生成PDF格式的运维服务月报"""
    try:
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter)
        elements = []

        # 设置样式
        styles = getSampleStyleSheet()
        title_style = styles['Title']
        subtitle_style = styles['Heading2']
        normal_style = styles['Normal']
        for style in [title_style, subtitle_style, normal_style]:
            style.fontName = 'SimKai'

        # 添加标题
        title = (f"<para alignment='center'>iTop 运维服务报表 "
                f"({start_date.year}年{start_date.month}月"
                f"{f'至{end_date.year}年{end_date.month}月' if start_date.month != end_date.month else ''})"
                f"</para>")
        elements.append(Paragraph(title, title_style))
        elements.append(Spacer(1, 12))

        # 添加报表概括
        total_tickets = ticket_summary['total'].iloc[0] if not ticket_summary.empty else 0
        elements.append(Paragraph(f"iTop共接收工单数 {total_tickets} 起，各类工单处理情况如下：", subtitle_style))
        elements.append(Spacer(1, 12))

        # 添加各部分内容
        elements.append(Paragraph("1. 按服务类型统计分析如下：", subtitle_style))
        _add_service_request_stats(elements, user_request_stats, incident_stats, change_stats, normal_style, subtitle_style)
        elements.append(Paragraph("2. 按照工单处理团队统计，具体如下", subtitle_style))
        _add_team_stats(elements, team_stats, normal_style, subtitle_style)
        
        # 添加工程师统计
        elements.append(Paragraph("3. 按照工单处理工程师统计，具体如下", subtitle_style))
        if not person_stats.empty:
            person_data = [person_stats.columns.tolist()] + person_stats.values.tolist()
            elements.append(_create_data_table(person_data, normal_style))
        else:
            elements.append(Paragraph("本周期内没有要处理的工单", normal_style))
        elements.append(Spacer(1, 12))

        # 添加未解决工单
        elements.append(Paragraph("4. 未解决的工单如下", subtitle_style))
        if not unresolved_tickets.empty:
            unresolved_data = [unresolved_tickets.columns.tolist()] + unresolved_tickets.values.tolist()
            elements.append(_create_data_table(unresolved_data, normal_style))
        else:
            elements.append(Paragraph("本周期内没有未解决的工单。", normal_style))
        elements.append(Spacer(1, 12))

        # 添加未解决工单
        elements.append(Paragraph("5. SLA超时的工单如下", subtitle_style))
        if not overdue_tickets.empty:
            overdue_tickets = [overdue_tickets.columns.tolist()] + overdue_tickets.values.tolist()
            elements.append(_create_data_table(overdue_tickets, normal_style))
        else:
            elements.append(Paragraph("本周期内没有SLA超时的工单。", normal_style))
        elements.append(Spacer(1, 12))

        # 添加KPI统计
        _add_kpi_stats(elements, infra_kpi_stats, "6. Infra KPI 统计", normal_style, subtitle_style)
        _add_kpi_stats(elements, app_kpi_stats, "7. 应用 KPI 统计", normal_style, subtitle_style)

        # 生成PDF
        doc.build(elements)
        return buffer.getvalue()
    except Exception as e:
        print(f"生成PDF报告时发生错误: {str(e)}")
        return None
    finally:
        buffer.close()