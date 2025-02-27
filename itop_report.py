import streamlit as st
import pandas as pd
from sqlalchemy import create_engine
from datetime import datetime, timedelta, date
import plotly.express as px
import calendar
import configparser
from pdf_generator import generate_pdf

# 定义饼图
def create_pie_chart(names, values, title):
    """创建饼图
    
    Args:
        names (list): 饼图各部分的名称列表
        values (list): 饼图各部分对应的数值列表
        title (str): 饼图标题
        
    Returns:
        plotly.graph_objects.Figure: 返回一个plotly饼图对象
        
    Example:
        >>> names = ['已解决', '未解决', '已关闭']
        >>> values = [60, 30, 10]
        >>> title = '工单状态分布'
        >>> fig = create_pie_chart(names, values, title)
    """
    fig = px.pie(names=names, 
                    values=values, 
                    title=title, 
                    color_discrete_sequence=['#00b8a9', '#f6416c', '#f8f3d4'])
    fig.update_traces(textposition='inside', 
                    textinfo='label+percent')
    fig.update_layout(title_x=0.35)  # 将标题居中显示

    return fig

# 定义折线图
def create_line_chart(data, x_desc, y_desc, data_desc, line_type, title):
    """创建折线图
    
    Args:
        data (pandas.DataFrame): 包含绘图数据的DataFrame
        x_desc (str): x轴数据列名
        y_desc (str): y轴数据列名 
        data_desc (str): 数据标签列名
        line_type (str): 线条类型列名,用于区分不同线条,可选
        title (str): 图表标题
        
    Returns:
        plotly.graph_objects.Figure: 返回一个plotly折线图对象
        
    Example:
        >>> df = pd.DataFrame({
        ...     '月份': ['1月', '2月', '3月'],
        ...     '解决率': [80, 85, 90],
        ...     '标签': ['80%', '85%', '90%'],
        ...     '团队': ['团队A', '团队A', '团队A']
        ... })
        >>> fig = create_line_chart(df, '月份', '解决率', '标签', '团队', '工单解决率趋势')
    """
    # 根据是否指定线条类型来创建折线图
    if line_type:
        # 如果指定了线条类型，则使用指定的线条类型
        fig = px.line(data, 
                    x=x_desc, 
                    y=y_desc,
                    text=data_desc,  # 添加数值标签
                    markers=True,
                    title=title,
                    color=line_type)
    else:
        # 如果没有指定线条类型，则默认使用橙色
        fig = px.line(data, 
                    x=x_desc, 
                    y=y_desc,
                    text=data_desc,  # 添加数值标签
                    markers=True,
                    title=title,
                    color_discrete_sequence=['orange'])  # 线条颜色设置为橙色
    
    # 配置数值标签的显示位置和格式
    fig.update_traces(
        textposition="top center",  # 将数值显示在点的上方居中
        texttemplate='%{text}'  # 显示格式:显示具体的数值
    )
    
    # 更新布局以调整标题、坐标轴、边距和图例的样式
    fig.update_layout(
        title_x=0.35,
        title_y=0.95, # 将标题向上移动
        xaxis_title=x_desc,
        yaxis_title=y_desc,
        yaxis=dict(range=[0, 110]),  # 设置y轴范围
        xaxis=dict(
            type='category',
            categoryorder='category ascending'  # 设置x轴为类别型，按类别顺序排序
        ),
        margin=dict(t=100), # 增加顶部边距
        legend=dict(
            orientation="h",  # 设置图例方向为水平
            yanchor="bottom",
            y=1.05,  # 调整图例位置,与标题保持10px间距
            xanchor="center",
            x=0.5,  # 图例水平居中
            itemwidth=30,  # 设置图例项的宽度,使团队名称显示在一行
            title=None  # 取消图例标题
        )
    )

    return fig

# 连接到iTop数据库
def connect_to_itop_db():
    # 从配置文件读取数据库连接信息    
    config = configparser.ConfigParser()
    config.read('config.ini')
    
    db_host = config['Database']['host']
    db_user = config['Database']['user']
    db_password = config['Database']['password']
    db_port = config['Database']['port']
    db_name = config['Database']['database']
    
    return create_engine(f'mysql+pymysql://{db_user}:{db_password}@{db_host}:{db_port}/{db_name}')

# 执行SQL查询并返回DataFrame
def execute_query(engine, query, params):
    # 将日期转换为字符串格式
    for key, value in params.items():
        if isinstance(value, (date, datetime)):
            params[key] = value.strftime('%Y-%m-%d')
    
    # print("Executing query:", query)
    # print("With parameters:", params)
    
    with engine.connect() as connection:
        df = pd.read_sql(query, connection, params=params)
    return df

# 1. 工单统计
def get_ticket_summary(engine, start_date, end_date):
    query = """
    SELECT 
        count(1) as total,
        SUM(CASE WHEN t.finalclass = 'UserRequest' THEN 1 ELSE 0 END) as request_total,
        SUM(CASE WHEN t.finalclass LIKE '%%Change%%' THEN 1 ELSE 0 END) as change_total,
        SUM(CASE WHEN t.finalclass = 'Incident' THEN 1 ELSE 0 END) as Incident_total
    FROM ticket t 
    LEFT JOIN ticket_request tr ON tr.id = t.id 
    LEFT JOIN ticket_incident ti ON ti.id = t.id 
    LEFT JOIN `change` c ON c.id = t.id 
    WHERE t.finalclass <> 'Problem' 
    AND t.start_date >= %(start_date)s
    AND t.start_date < %(end_date)s
    AND (tr.status <> 'new' OR ti.status <> 'new' OR c.status <> 'new')
    """
    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 2. 服务请求状态统计
def get_user_request_stats(engine, start_date, end_date):
    query = """
    SELECT
        count(1) as total,
        SUM(CASE WHEN tr.status in ('closed','resolved') THEN 1 ELSE 0 END) as resolved_total,
        SUM(CASE WHEN tr.status = 'closed' THEN 1 ELSE 0 END) as closed_total,
        SUM(CASE WHEN tr.status not in ('closed','resolved') THEN 1 ELSE 0 END) as unresolved_total
    FROM ticket_request tr  
    LEFT JOIN ticket t ON tr.id = t.id 
    WHERE t.start_date >= %(start_date)s
    AND t.start_date < %(end_date)s
    AND tr.status <> 'new'
    """
    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 3. 事件状态统计
def get_incident_stats(engine, start_date, end_date):
    query = """
    SELECT
        count(1) as total,
        SUM(CASE WHEN ti.status in ('closed','resolved') THEN 1 ELSE 0 END) as resolved_total,
        SUM(CASE WHEN ti.status = 'closed' THEN 1 ELSE 0 END) as closed_total,
        SUM(CASE WHEN ti.status not in ('closed','resolved') THEN 1 ELSE 0 END) as unresolved_total
    FROM ticket_incident ti  
    LEFT JOIN ticket t ON ti.id = t.id 
    WHERE t.start_date >= %(start_date)s
    AND t.start_date < %(end_date)s
    AND ti.status <> 'new'
    """
    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 4. 变更状态统计
def get_change_stats(engine, start_date, end_date):
    query = """
    SELECT
        count(1) as total,
        SUM(CASE WHEN c.status in ('closed','resolved') THEN 1 ELSE 0 END) as resolved_total,
        SUM(CASE WHEN c.status = 'closed' THEN 1 ELSE 0 END) as closed_total,
        SUM(CASE WHEN c.status not in ('closed','resolved') THEN 1 ELSE 0 END) as unresolved_total
    FROM `change` c 
    LEFT JOIN ticket t ON c.id = t.id 
    WHERE t.start_date >= %(start_date)s
    AND t.start_date < %(end_date)s
    AND c.status <> 'new'
    """
    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 5. 按团队统计处理时长
def get_team_stats(engine, start_date, end_date):
    query = """
    SELECT 
        DATE_FORMAT(subquery.start_date, '%%Y年%%m月') AS '月份',
        c.name AS '团队',
        subquery.ticket_type AS '工单类型',
        COUNT(*) AS '工单数量',
        SUM(CASE WHEN subquery.status IN ('closed', 'resolved') THEN 1 ELSE 0 END) AS '已解决',
        SUM(CASE WHEN subquery.status NOT IN ('closed', 'new', 'resolved') THEN 1 ELSE 0 END) AS '未解决',
        SUM(CASE WHEN (subquery.tto_100_passed = 1 OR subquery.ttr_100_passed = 1) THEN 1 ELSE 0 END) AS '超时工单',
        CONCAT(
            ROUND(
                (COUNT(*) - SUM(CASE WHEN subquery.status NOT IN ('closed', 'new', 'resolved') THEN 1 ELSE 0 END)) * 100.0 / 
                NULLIF(COUNT(*), 0),
                2
            ),
            '%%'
        ) AS '工单解决率',
        CONCAT(
            ROUND(
                (COUNT(*) - SUM(CASE WHEN (subquery.tto_100_passed = 1 OR subquery.ttr_100_passed = 1) THEN 1 ELSE 0 END)) * 100.0 / 
                NULLIF(COUNT(*), 0),
                2
            ),
            '%%'
        ) AS '工单及时率',
        CASE 
            WHEN subquery.ticket_type = '变更' THEN 'N/A'
            ELSE ROUND(AVG(subquery.response_time) / 60, 2)
        END AS '平均响应时长(分钟)',
        ROUND(AVG(subquery.resolution_time) / 60, 2) AS '平均解决时长(分钟)',
        CASE 
            WHEN subquery.ticket_type = '变更' THEN 'N/A'
            ELSE ROUND(MAX(subquery.response_time) / 60, 2)
        END AS '最大响应时长(分钟)',
        ROUND(MAX(subquery.resolution_time) / 60, 2) AS '最大解决时长(分钟)'
    FROM (
        SELECT 
            t.team_id,
            t.start_date,
            '服务请求' AS ticket_type,
            tr.status,
            tr.tto_100_passed,
            tr.ttr_100_passed,
            TIMESTAMPDIFF(SECOND, tr.tto_started, tr.tto_stopped) AS response_time,
            TIMESTAMPDIFF(SECOND, tr.tto_stopped, tr.ttr_stopped) AS resolution_time
        FROM ticket t 
        JOIN ticket_request tr ON tr.id = t.id 
        WHERE tr.status <> 'new'
            AND t.start_date >= %(start_date)s
            AND t.start_date < %(end_date)s
        
        UNION ALL
        
        SELECT 
            t.team_id,
            t.start_date,
            '事件' AS ticket_type,
            ti.status,
            ti.tto_100_passed,
            ti.ttr_100_passed,
            TIMESTAMPDIFF(SECOND, ti.tto_started, ti.tto_stopped) AS response_time,
            TIMESTAMPDIFF(SECOND, ti.tto_stopped, ti.ttr_stopped) AS resolution_time
        FROM ticket t 
        JOIN ticket_incident ti ON ti.id = t.id
        WHERE ti.status <> 'new'
            AND t.start_date >= %(start_date)s
            AND t.start_date < %(end_date)s
        
        UNION ALL
        
        SELECT 
            t.team_id,
            t.start_date,
            '变更' AS ticket_type,
            c2.status,
            0 AS tto_100_passed,  -- 变更工单没有响应时间要求
            0 AS ttr_100_passed,  -- 变更工单暂不考虑解决时间超时
            NULL AS response_time,
            TIMESTAMPDIFF(SECOND, t.start_date, t.end_date) AS resolution_time
        FROM ticket t 
        JOIN `change` c2 ON c2.id = t.id
        WHERE c2.status <> 'new'
            AND t.start_date >= %(start_date)s
            AND t.start_date < %(end_date)s
    ) AS subquery
    JOIN contact c ON subquery.team_id = c.id 
    WHERE c.finalclass = 'Team'
    GROUP BY DATE_FORMAT(subquery.start_date, '%%Y-%%m'), c.name, subquery.ticket_type
    HAVING COUNT(*) > 0
    ORDER BY DATE_FORMAT(subquery.start_date, '%%Y-%%m') DESC, subquery.ticket_type DESC, c.name
    """
    
    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 6. 按人员统计处理时长
def get_person_stats(engine, start_date, end_date):
    query = """
    WITH agent_info AS (
        SELECT p1.id, CONCAT(COALESCE(c1.name, ''), ' ', COALESCE(p1.first_name, '')) AS agent_name
        FROM person p1
        JOIN contact c1 ON p1.id = c1.id
    )

    SELECT 
        DATE_FORMAT(start_date, '%%Y年%%m月') AS '月份',
        ai.agent_name AS '办理人',
        ticket_type AS '工单类型',
        COUNT(*) AS '工单数量',
        SUM(CASE WHEN status IN ('closed', 'resolved') THEN 1 ELSE 0 END) AS '已解决',
        SUM(CASE WHEN status NOT IN ('closed', 'new', 'resolved') THEN 1 ELSE 0 END) AS '未解决',
        SUM(CASE WHEN (tto_100_passed = 1 OR ttr_100_passed = 1) THEN 1 ELSE 0 END) AS '超时工单',
        CONCAT(
            ROUND(
                (COUNT(*) - SUM(CASE WHEN status NOT IN ('closed', 'new', 'resolved') THEN 1 ELSE 0 END)) * 100.0 / 
                NULLIF(COUNT(*), 0),
                2
            ),
            '%%'
        ) AS '工单解决率',
        CONCAT(
            ROUND(
                (COUNT(*) - SUM(CASE WHEN (tto_100_passed = 1 OR ttr_100_passed = 1) THEN 1 ELSE 0 END)) * 100.0 / 
                NULLIF(COUNT(*), 0),
                2
            ),
            '%%'
        ) AS '工单及时率',
        CASE 
            WHEN ticket_type = '变更' THEN 'N/A'
            ELSE ROUND(AVG(response_time) / 60, 2)
        END AS '平均响应时长(分钟)',
        ROUND(AVG(resolution_time) / 60, 2) AS '平均解决时长(分钟)',
        CASE 
            WHEN ticket_type = '变更' THEN 'N/A'
            ELSE ROUND(MAX(response_time) / 60, 2)
        END AS '最大响应时长(分钟)',
        ROUND(MAX(resolution_time) / 60, 2) AS '最大解决时长(分钟)'
    FROM (
        SELECT 
            t.agent_id,
            '服务请求' AS ticket_type,
            tr.status,
            t.start_date,
            tr.tto_100_passed,
            tr.ttr_100_passed,
            TIMESTAMPDIFF(SECOND, tr.tto_started, tr.tto_stopped) AS response_time,
            TIMESTAMPDIFF(SECOND, tr.tto_stopped, tr.ttr_stopped) AS resolution_time
        FROM ticket_request tr
        JOIN ticket t ON tr.id = t.id
        WHERE tr.status <> 'new'
            AND t.start_date >= %(start_date)s
            AND t.start_date < %(end_date)s
        
        UNION ALL
        
        SELECT 
            t.agent_id,
            '事件' AS ticket_type,
            ti.status,
            t.start_date,
            ti.tto_100_passed,
            ti.ttr_100_passed,
            TIMESTAMPDIFF(SECOND, ti.tto_started, ti.tto_stopped) AS response_time,
            TIMESTAMPDIFF(SECOND, ti.tto_stopped, ti.ttr_stopped) AS resolution_time
        FROM ticket_incident ti
        JOIN ticket t ON ti.id = t.id
        WHERE ti.status <> 'new'
            AND t.start_date >= %(start_date)s
            AND t.start_date < %(end_date)s
        
        UNION ALL
        
        SELECT 
            t.agent_id,
            '变更' AS ticket_type,
            c2.status,
            t.start_date,
            0 AS tto_100_passed,  -- 变更工单没有响应时间要求
            0 AS ttr_100_passed,  -- 变更工单暂不考虑解决时间超时
            NULL AS response_time,
            TIMESTAMPDIFF(SECOND, t.start_date, t.end_date) AS resolution_time
        FROM `change` c2
        JOIN ticket t ON c2.id = t.id
        WHERE c2.status <> 'new'
            AND t.start_date >= %(start_date)s
            AND t.start_date < %(end_date)s
    ) AS subquery
    JOIN agent_info ai ON subquery.agent_id = ai.id
    GROUP BY ai.agent_name, ticket_type, DATE_FORMAT(start_date, '%%Y-%%m')
    ORDER BY DATE_FORMAT(start_date, '%%Y-%%m') DESC, ticket_type DESC, ai.agent_name
    """

    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 7. 未解决的工单
def get_unresolved_tickets(engine, start_date, end_date):
    query = """
    SELECT 
        t.ref AS '工单号', 
        t.title AS '标题', 
        t.start_date AS '开始时间',
        CASE 
            WHEN tr.id IS NOT NULL THEN tr.status
            WHEN ti.id IS NOT NULL THEN ti.status
            WHEN cg.id IS NOT NULL THEN cg.status
        END AS '状态', 
        CONCAT(IFNULL(c.name, ''), ' ', IFNULL(p.first_name, '')) AS '发起人', 
        c2.name AS '团队名称', 
        CONCAT(IFNULL(c1.name, ''), ' ', IFNULL(p1.first_name, '')) AS '办理人'
    FROM ticket t
    LEFT JOIN ticket_request tr ON tr.id = t.id
    LEFT JOIN ticket_incident ti ON ti.id = t.id
    LEFT JOIN `change` cg ON cg.id = t.id
    LEFT JOIN (person p JOIN contact c ON p.id = c.id) ON t.caller_id = p.id 
    LEFT JOIN (person p1 JOIN contact c1 ON p1.id = c1.id) ON t.agent_id = p1.id 
    LEFT JOIN contact c2 ON t.team_id = c2.id 
    WHERE (tr.status NOT IN ('closed','new','resolved') 
        OR ti.status NOT IN ('closed','new','resolved')
        OR cg.status NOT IN ('closed','new','resolved'))
    AND t.start_date >= %(start_date)s
    AND t.start_date < %(end_date)s
    """
    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 8. 超时工单
def get_overdue_tickets(engine, start_date, end_date):
    query = """
    SELECT 
        (CASE WHEN tr.status THEN CAST(CONCAT(COALESCE(t.ref, '')) AS CHAR) ELSE NULL END) AS '工单号', 
        t.title AS '标题',
        tr.status AS '状态', 
        t.start_date AS '开始日期',
        t.last_update AS '最后日期',
        ROUND(tr.tto_100_overrun / 60,2) AS '响应时间超过(分钟)',
        ROUND(tr.ttr_100_overrun / 60,2) AS '解决时间超过(分钟)',
        (CASE WHEN tr.status THEN CAST(CONCAT(COALESCE(c1.name, ''), COALESCE(' ', ''), COALESCE(p1.first_name, '')) AS CHAR) ELSE NULL END) AS '发起人', 
        (CASE WHEN tr.status THEN CAST(CONCAT(COALESCE(c2.name, '')) AS CHAR) ELSE NULL END) AS '团队名称', 
        (CASE WHEN tr.status THEN CAST(CONCAT(COALESCE(c1.name, ''), COALESCE(' ', ''), COALESCE(p1.first_name, '')) AS CHAR) ELSE NULL END) AS '办理人',
        tr.assignment_date AS '实际响应时间',
        tr.resolution_date AS '实际解决时间',
        tr.tto_100_deadline AS '响应最后期限',
        tr.ttr_100_deadline AS '解决最后期限',
        ROUND(AVG(TIMESTAMPDIFF(SECOND,tr.tto_started,tr.tto_stopped) / 60),2) AS '响应时长(分钟)' , 
        ROUND(AVG(TIMESTAMPDIFF(SECOND,tr.tto_stopped,tr.ttr_stopped) / 60),2) AS '解决时长(分钟)'
    FROM ticket t 
    LEFT JOIN ticket_request tr ON tr.id = t.id 
    LEFT JOIN ( person AS p INNER JOIN contact AS c ON p.id = c.id ) ON tr.approver_id = p.id 
    LEFT JOIN ( person AS p1 INNER JOIN contact AS c1 ON p1.id = c1.id ) ON t.agent_id = p1.id 
    LEFT JOIN contact AS c2 ON t.team_id = c2.id 
    WHERE t.finalclass <> 'Problem' 
    AND ( tr.tto_100_passed = 1 or tr.ttr_100_passed = 1 )
    AND t.start_date >= %(start_date)s
    AND t.start_date < %(end_date)s
    GROUP BY tr.status
    HAVING COUNT(*) > 0
    """
    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 9. Infra KPI 统计
def get_infra_kpi_stats(engine, start_date, end_date):
    # 获取本周期内发生工单的服务目录
    service_query = """
    SELECT 
        DISTINCT COALESCE(s.name, COALESCE(s2.name, '未分类')) AS service
    FROM (
        SELECT 
            tr.service_id,
            tr.servicesubcategory_id,
            tr.status,
            t.start_date 
        FROM ticket_request tr
        JOIN ticket t ON tr.id = t.id
        WHERE t.finalclass <> 'Problem'
        AND t.start_date >= %(start_date)s
        AND t.start_date < %(end_date)s
        AND tr.status <> 'new'
        
        UNION ALL
        
        SELECT
            ti.service_id,
            ti.servicesubcategory_id,
            ti.status,
            t.start_date 
        FROM ticket_incident ti
        JOIN ticket t ON ti.id = t.id
        WHERE t.finalclass <> 'Problem'
        AND t.start_date >= %(start_date)s
        AND t.start_date < %(end_date)s
        AND ti.status <> 'new'
    ) subquery
    LEFT JOIN service s ON subquery.service_id = s.id
    LEFT JOIN servicesubcategory s2 ON subquery.servicesubcategory_id = s2.id
    WHERE s.name <> '应用'
    """
    service = execute_query(engine, service_query, {'start_date': start_date, 'end_date': end_date})
    tmp_columns = ''
    columns = ''
    total_columns = ''
    for index, row in service.iterrows():
        tmp_columns = tmp_columns + "CASE WHEN tmp_info.service_name = '{}' THEN tmp_info.ticket_resolution_rate ELSE null END AS '{}',".format(row['service'], row['service']) if row['service'] != '未分类' else tmp_columns
        columns = columns + "CONCAT(AVG(tmp.`{}`), '%%') AS '{}',".format(row['service'], row['service']) if row['service'] != '未分类' else columns
        total_columns = total_columns + "CASE WHEN tmp1.service_name = '{}' THEN tmp1.ticket_resolution_rate ELSE null END AS '{}',".format(row['service'], row['service']) if row['service'] != '未分类' else total_columns

    query = """
    WITH tmp_info AS (
    SELECT 
        DATE_FORMAT(start_date, '%%Y年%%m月') AS ticket_month,
        COALESCE(s2.name, COALESCE(s.name, '未分类')) AS service_name, 
        COUNT(*) as total,
        COUNT(*) - SUM(CASE WHEN subquery.status NOT IN ('closed', 'new', 'resolved') OR subquery.ttr_100_passed = 1 THEN 1 ELSE 0 END) AS resolved,
        CONCAT(
            ROUND(
                (COUNT(*) - SUM(CASE WHEN subquery.status NOT IN ('closed', 'new', 'resolved') OR subquery.ttr_100_passed = 1 THEN 1 ELSE 0 END))  / 
                NULLIF(COUNT(*), 0) * 100.0,
                2
            ),
            '%%'
        ) AS ticket_resolution_rate
    FROM (
        SELECT 
            tr.service_id,
            tr.servicesubcategory_id,
            tr.status,
            tr.tto_100_passed,
            tr.ttr_100_passed,
            t.start_date 
        FROM ticket_request tr
        JOIN ticket t ON tr.id = t.id
        WHERE t.finalclass <> 'Problem'
        AND t.start_date >= %(start_date)s
        AND t.start_date < %(end_date)s
        AND tr.status <> 'new'
        
        UNION ALL
        
        SELECT
            ti.service_id,
            ti.servicesubcategory_id,
            ti.status,
            ti.tto_100_passed,
            ti.ttr_100_passed,
            t.start_date
        FROM ticket_incident ti
        JOIN ticket t ON ti.id = t.id
        WHERE t.finalclass <> 'Problem'
        AND t.start_date >= %(start_date)s
        AND t.start_date < %(end_date)s
        AND ti.status <> 'new'
    ) subquery
    LEFT JOIN service s ON subquery.service_id = s.id
    LEFT JOIN servicesubcategory s2 ON subquery.servicesubcategory_id = s2.id
    WHERE s.name <> '应用'
    GROUP BY DATE_FORMAT(start_date, '%%Y年%%m月'), s.name, s2.name
    ORDER BY '服务' DESC
    )

    SELECT 
        tmp.ticket_month as '月份',
        {}
        CONCAT(AVG(tmp.`未分类`),'%%') AS '未分类',
        CONCAT(
            ROUND(
                SUM(tmp.`已解决`)  / 
                NULLIF(SUM(tmp.`工单总数`), 0) * 100.0,
                2
            ),
            '%%'
        ) AS 'KPI总计',
        SUM(tmp.`工单总数`)  AS '工单总数',
        SUM(tmp.`已解决`)  AS '已解决'
    FROM (
        SELECT 
            tmp_info.ticket_month,
            {}
            CASE WHEN tmp_info.service_name = '未分类' THEN tmp_info.ticket_resolution_rate ELSE null END AS '未分类',
            SUM(tmp_info.total) AS '工单总数',
            SUM(tmp_info.resolved) AS '已解决'
        FROM tmp_info
        GROUP BY tmp_info.ticket_month,tmp_info.service_name

        UNION 
        
        SELECT 
            '总计' AS ticket_month,
            {}
            CASE WHEN tmp1.service_name = '未分类' THEN tmp1.ticket_resolution_rate ELSE null END AS '未分类',
            SUM(tmp1.total) AS '工单总数',
            SUM(tmp1.resolved) AS '已解决'
        FROM (
            SELECT 
                service_name,
                CONCAT(
                    ROUND(
                        SUM(resolved)  / 
                        NULLIF(SUM(total), 0) * 100.0,
                        2
                    ),
                    '%%'
                ) AS ticket_resolution_rate,
                SUM(total)  AS total,
                SUM(resolved)  AS resolved
            FROM tmp_info
            GROUP BY service_name
        ) tmp1
        GROUP BY tmp1.service_name
    ) tmp
    GROUP BY tmp.ticket_month  
    """.format(columns, tmp_columns, total_columns)

    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

# 9. 应用 KPI 统计
def get_app_kpi_stats(engine, start_date, end_date):
    # 获取本周期内发生工单的服务目录
    service_query = """
    SELECT 
        DISTINCT COALESCE(s2.name, COALESCE(s.name, '未分类')) AS service
    FROM (
        SELECT 
            tr.service_id,
            tr.servicesubcategory_id,
            tr.status,
            t.start_date 
        FROM ticket_request tr
        JOIN ticket t ON tr.id = t.id
        WHERE t.finalclass <> 'Problem'
        AND t.start_date >= %(start_date)s
        AND t.start_date < %(end_date)s
        AND tr.status <> 'new'
        
        UNION ALL
        
        SELECT
            ti.service_id,
            ti.servicesubcategory_id,
            ti.status,
            t.start_date 
        FROM ticket_incident ti
        JOIN ticket t ON ti.id = t.id
        WHERE t.finalclass <> 'Problem'
        AND t.start_date >= %(start_date)s
        AND t.start_date < %(end_date)s
        AND ti.status <> 'new'
    ) subquery
    LEFT JOIN service s ON subquery.service_id = s.id
    LEFT JOIN servicesubcategory s2 ON subquery.servicesubcategory_id = s2.id
    WHERE s.name = '应用'
    """
    service = execute_query(engine, service_query, {'start_date': start_date, 'end_date': end_date})
    tmp_columns = ''
    columns = ''
    total_columns = ''
    for index, row in service.iterrows():
        tmp_columns = tmp_columns + "CASE WHEN tmp_info.service_name = '{}' THEN tmp_info.ticket_resolution_rate ELSE null END AS '{}',".format(row['service'], row['service']) if row['service'] != '未分类' else tmp_columns
        columns = columns + "CONCAT(AVG(tmp.`{}`), '%%') AS '{}',".format(row['service'], row['service']) if row['service'] != '未分类' else columns
        total_columns = total_columns + "CASE WHEN tmp1.service_name = '{}' THEN tmp1.ticket_resolution_rate ELSE null END AS '{}',".format(row['service'], row['service']) if row['service'] != '未分类' else total_columns

    query = """
    WITH tmp_info AS (
    SELECT 
        DATE_FORMAT(start_date, '%%Y年%%m月') AS ticket_month,
        COALESCE(s2.name, COALESCE(s.name, '未分类')) AS service_name, 
        COUNT(*) as total,
        COUNT(*) - SUM(CASE WHEN subquery.status NOT IN ('closed', 'new', 'resolved') OR subquery.ttr_100_passed = 1 THEN 1 ELSE 0 END) AS resolved,
        CONCAT(
            ROUND(
                (COUNT(*) - SUM(CASE WHEN subquery.status NOT IN ('closed', 'new', 'resolved') OR subquery.ttr_100_passed = 1 THEN 1 ELSE 0 END))  / 
                NULLIF(COUNT(*), 0) * 100.0,
                2
            ),
            '%%'
        ) AS ticket_resolution_rate
    FROM (
        SELECT 
            tr.service_id,
            tr.servicesubcategory_id,
            tr.status,
            tr.tto_100_passed,
            tr.ttr_100_passed,
            t.start_date 
        FROM ticket_request tr
        JOIN ticket t ON tr.id = t.id
        WHERE t.finalclass <> 'Problem'
        AND t.start_date >= %(start_date)s
        AND t.start_date < %(end_date)s
        AND tr.status <> 'new'
        
        UNION ALL
        
        SELECT
            ti.service_id,
            ti.servicesubcategory_id,
            ti.status,
            ti.tto_100_passed,
            ti.ttr_100_passed,
            t.start_date 
        FROM ticket_incident ti
        JOIN ticket t ON ti.id = t.id
        WHERE t.finalclass <> 'Problem'
        AND t.start_date >= %(start_date)s
        AND t.start_date < %(end_date)s
        AND ti.status <> 'new'
    ) subquery
    LEFT JOIN service s ON subquery.service_id = s.id
    LEFT JOIN servicesubcategory s2 ON subquery.servicesubcategory_id = s2.id
    WHERE s.name = '应用'
    GROUP BY DATE_FORMAT(start_date, '%%Y年%%m月'), s.name, s2.name
    ORDER BY '服务' DESC
    )

    SELECT 
        tmp.ticket_month as '月份',
        {}
        CONCAT(AVG(tmp.`未分类`),'%%') AS '未分类',
        CONCAT(
            ROUND(
                SUM(tmp.`已解决`)  / 
                NULLIF(SUM(tmp.`工单总数`), 0) * 100.0,
                2
            ),
            '%%'
        ) AS 'KPI总计',
        SUM(tmp.`工单总数`)  AS '工单总数',
        SUM(tmp.`已解决`)  AS '已解决'
    FROM (
        SELECT 
            tmp_info.ticket_month,
            {}
            CASE WHEN tmp_info.service_name = '未分类' THEN tmp_info.ticket_resolution_rate ELSE null END AS '未分类',
            SUM(tmp_info.total) AS '工单总数',
            SUM(tmp_info.resolved) AS '已解决'
        FROM tmp_info
        GROUP BY tmp_info.ticket_month,tmp_info.service_name

        UNION 
        
        SELECT 
            '总计' AS ticket_month,
            {}
            CASE WHEN tmp1.service_name = '未分类' THEN tmp1.ticket_resolution_rate ELSE null END AS '未分类',
            SUM(tmp1.total) AS '工单总数',
            SUM(tmp1.resolved) AS '已解决'
        FROM (
            SELECT 
                service_name,
                CONCAT(
                    ROUND(
                        SUM(resolved)  / 
                        NULLIF(SUM(total), 0) * 100.0,
                        2
                    ),
                    '%%'
                ) AS ticket_resolution_rate,
                SUM(total)  AS total,
                SUM(resolved)  AS resolved
            FROM tmp_info
            GROUP BY service_name
        ) tmp1
        GROUP BY tmp1.service_name
    ) tmp
    GROUP BY tmp.ticket_month  
    """.format(columns, tmp_columns, total_columns)

    return execute_query(engine, query, {'start_date': start_date, 'end_date': end_date})

def main():
    # 创建左边栏
    with st.sidebar:
        st.title("iTop 报表查询")
        st.markdown("<style>h1{text-align: center;}</style>", unsafe_allow_html=True)
        # 添加一条横线
        st.markdown("---")

        # 添加日期选择提示
        st.markdown("""
        <div>  </div>
        <div style='color: #808080; font-style: italic;'>
        请选择要查询的开始日期和结束日期\r\n
        (系统默认为上一个月的数据)
        </div>
        """, unsafe_allow_html=True)

        # 日期选择
        today = datetime.now()
        last_month = today.replace(day=1) - timedelta(days=1)
        
        st.markdown("开始日期", unsafe_allow_html=True)
        start_date = st.date_input("", last_month.replace(day=1), key="start_date", label_visibility="collapsed")
        
        st.markdown("结束日期", unsafe_allow_html=True)
        end_date = st.date_input("", last_month.replace(day=calendar.monthrange(last_month.year, last_month.month)[1]), key="end_date", label_visibility="collapsed")

        # 连接数据库
        engine = connect_to_itop_db()

        # 获取数据
        ticket_summary = get_ticket_summary(engine, start_date, end_date)
        user_request_stats = get_user_request_stats(engine, start_date, end_date)
        incident_stats = get_incident_stats(engine, start_date, end_date)
        change_stats = get_change_stats(engine, start_date, end_date)
        team_stats = get_team_stats(engine, start_date, end_date)
        person_stats = get_person_stats(engine, start_date, end_date)
        unresolved_tickets = get_unresolved_tickets(engine, start_date, end_date)
        overdue_tickets = get_overdue_tickets(engine, start_date, end_date)
        infra_kpi_stats = get_infra_kpi_stats(engine, start_date, end_date)
        app_kpi_stats = get_app_kpi_stats(engine, start_date, end_date)

        # 插入一行空行
        st.write("")

        # 添加导出PDF按钮
        col1, col2, col3 = st.columns([1, 1, 2])
        with col3:
            if st.button('导出PDF报表'):
                try:
                    pdf = generate_pdf(start_date, end_date, 
                                       ticket_summary, user_request_stats, 
                                       incident_stats, change_stats, 
                                       team_stats, person_stats, 
                                       unresolved_tickets, overdue_tickets, 
                                       infra_kpi_stats, app_kpi_stats)
                    with col3:
                        st.download_button(
                            label="下载PDF报表",
                            data=pdf,
                            file_name="itop_report.pdf",
                            mime="application/pdf"
                        )
                except Exception as e:
                    st.error(f"生成PDF时发生错误: {str(e)}")
                    st.error("请检查是否安装了所需的中文字体。")

    # 主要内容区域
    st.markdown("<h2 style='text-align: center;'>iTop 运维服务报表</h2>", unsafe_allow_html=True)

    # 显示报告
    if start_date.month == end_date.month:
        st.markdown(f"<div style='text-align: right; color: #808080; font-style: italic;'>服务周期：{start_date.year}年{start_date.month}月</div>", unsafe_allow_html=True)
    else:
        st.markdown(f"<div style='text-align: right; color: #808080; font-style: italic;'>服务周期：{start_date.year}年{start_date.month}月至{end_date.year}年{end_date.month}月</div>", unsafe_allow_html=True)

    # 添加一条横线
    st.markdown("---")
    # 0. 报表概况
    total_tickets = ticket_summary['total'].iloc[0] if not ticket_summary.empty else 0
    if start_date.month == end_date.month:
        st.write(f"#### {start_date.year}年{start_date.month}月iTop共接收工单数 {total_tickets} 起，各类工单处理情况如下：")
    else:
        st.write(f"#### {start_date.year}年{start_date.month}月至{end_date.year}年{end_date.month}月iTop共接收工单数 {total_tickets} 个，各类工单处理情况如下：")

    # 1. 按服务类型统计分析
    st.write("#### 1. 按服务类型统计分析如下：")

    # 1.1 服务请求统计
    st.write("##### 1) 服务请求统计")
    if not user_request_stats.empty:
        total = user_request_stats['total'].iloc[0]
        resolved = user_request_stats['resolved_total'].iloc[0]
        closed = user_request_stats['closed_total'].iloc[0]
        unresolved = user_request_stats['unresolved_total'].iloc[0]

        if total and total > 0:
            resolved_percentage = resolved / total * 100 if total > 0 else 0
            closed_percentage = closed / resolved * 100 if resolved > 0 else 0
            unresolved_percentage = unresolved / total * 100 if total > 0 else 0

            st.write(f"""
                     本周期内共接收服务请求 {total:g} 个，其中 {resolved:g} 个服务请求被解决，占比约 {resolved_percentage:.2f}%；
                     已解决的服务请求中，{closed:g} 个服务请求被按时关闭，占比约 {closed_percentage:.2f}%；
                     未解决的服务请求有 {unresolved:g} 个，占比约 {unresolved_percentage:.2f}%。
                     """)

            # 绘制饼图
            names = ['已解决', '未解决', '已关闭']
            values=[resolved, unresolved, closed]
            title = "服务请求状态分布"
            st.plotly_chart(create_pie_chart(names, values, title))
        else:
            st.write("本周期内没有接收到服务请求。")
    else:
        st.write("无法获取服务请求统计数据。")

    # 1.2 事件统计
    st.write("##### 2) 事件统计")
    if not incident_stats.empty:
        total = incident_stats['total'].iloc[0]
        resolved = incident_stats['resolved_total'].iloc[0]
        closed = incident_stats['closed_total'].iloc[0]
        unresolved = incident_stats['unresolved_total'].iloc[0]

        if total and total > 0:
            resolved_percentage = resolved / total * 100 if total > 0 else 0
            closed_percentage = closed / resolved * 100 if resolved > 0 else 0
            unresolved_percentage = unresolved / total * 100 if total > 0 else 0

            st.write(f"""
                     本周期内共发生事件 {total:g} 个，其中 {resolved:g} 个事件被解决，占比约 {resolved_percentage:.2f}%；
                     已解决的事件中，{closed:g} 个事件被按时关闭，占比约 {closed_percentage:.2f}%；
                     未解决的事件有 {unresolved:g} 个，占比约 {unresolved_percentage:.2f}%。
                     """)

            # 绘制饼图
            names = ['已解决', '未解决', '已关闭']
            values=[resolved, unresolved, closed]
            title = "事件状态分布"
            st.plotly_chart(create_pie_chart(names, values, title))
        else:
            st.write("本周期内没有发生事件。")
    else:
        st.write("无法获取事件统计数据。")

    # 1.3 变更统计
    st.write("##### 3) 变更统计")
    if not change_stats.empty:
        total = change_stats['total'].iloc[0]
        resolved = change_stats['resolved_total'].iloc[0]
        closed = change_stats['closed_total'].iloc[0]

        if total and total > 0:
            closed_percentage = closed / total * 100 if total > 0 else 0
            resolved_percentage = resolved / closed * 100 if closed > 0 else 0
            
            st.write(f"""
                     本周期内共发生变更 {total:g} 个，其中 {closed:g} 个变更已关闭，占比约 {closed_percentage:.2f}%。
                     已关闭的变更中，{resolved:g} 个变更被成功执行，占比约 {resolved_percentage:.2f}%。
                     """)

            # 绘制饼图
            names = ['已解决', '未解决']
            values=[resolved, total-resolved]
            title = "变更状态分布"
            st.plotly_chart(create_pie_chart(names, values, title))
        else:
            st.write("本周期内没有发生变更。")
    else:
        st.write("无法获取变更统计数据。")

    # 2. 按照工单处理团队统计
    st.write("#### 2. 按照工单处理团队统计，具体如下")
    st.dataframe(team_stats, use_container_width=True)

    # 2.1 按照工单处理团队绘制服务请求的解决率
    # 将team_stats转换为pandas DataFrame
    df = pd.DataFrame(team_stats)
    
    # 按月份和团队分组计算平均解决率和及时率
    df['工单解决率'] = df['工单解决率'].apply(lambda x: float(str(x).rstrip('%')))
    
    # 检查是否跨月
    if len(df['月份'].unique()) > 1:
        # 仅保留服务请求数据
        service_request_df = df[df['工单类型'] == '服务请求']
        # 如果工单解决率末尾没有%，则增加%
        service_request_df['工单解决率'] = service_request_df['工单解决率'].apply(lambda x: str(x) + '%' if not str(x).endswith('%') else x)
        # 创建解决率曲线图
        st.plotly_chart(create_line_chart(service_request_df, '月份', '工单解决率', '工单解决率', '团队', '各团队服务请求月度解决率趋势'))

    # 3. 按照工程师统计
    st.write("#### 3. 按照工单处理工程师统计，具体如下")
    st.dataframe(person_stats, use_container_width=True)

    # 5. 未解决的工单
    st.write("#### 4. 未解决的工单如下")
    st.dataframe(unresolved_tickets, use_container_width=True)

    # 5. SLA超时的工单
    st.write("#### 5. SLA超时的工单如下")
    st.dataframe(overdue_tickets, use_container_width=True)

    # 6. Infra KPI 统计
    st.write("#### 6. Infra KPI 统计")
    infra_kpi_stats = infra_kpi_stats.astype(str)
    infra_kpi_stats = infra_kpi_stats.applymap(lambda x: x.replace('b\'', '').replace('\'', '') if isinstance(x, str) else x)
    st.dataframe(infra_kpi_stats, use_container_width=True)   
    # 创建KPI曲线图
    st.plotly_chart(create_line_chart(infra_kpi_stats, '月份', 'KPI总计', 'KPI总计', '', 'KPI总计月度趋势'))

    # 7. 应用 KPI 统计
    st.write("#### 7. 应用 KPI 统计")
    app_kpi_stats = app_kpi_stats.astype(str)
    app_kpi_stats = app_kpi_stats.applymap(lambda x: x.replace('b\'', '').replace('\'', '') if isinstance(x, str) else x)
    st.dataframe(app_kpi_stats, use_container_width=True)   
    # 创建KPI曲线图
    st.plotly_chart(create_line_chart(app_kpi_stats, '月份', 'KPI总计', 'KPI总计', '', 'KPI总计月度趋势'))


if __name__ == "__main__":
    main()