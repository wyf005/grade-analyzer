import streamlit as st
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')
import numpy as np
import os
import io
from PIL import Image, ImageDraw, ImageFont
import matplotlib.font_manager as fm
import urllib.request
import tempfile
import ssl
import hashlib

# 中文字体解决方案 - 使用多个可能的源
CHINESE_FONT_PATH = None
CHINESE_FONT_SIZE = 20

def get_font_download_url():
    """返回可用的字体下载 URL"""
    # 尝试多个 CDN 源
    urls = [
        # 思源黑体 via jsDelivr (可能被墙)
        "https://cdn.jsdelivr.net/gh/googlefonts/noto-cjk@main/Sans/OTF/SimplifiedChinese/NotoSansCJKsc-Regular.otf",
        # 备选
        "https://fonts.gstatic.com/ea/notosanssc/v1/NotoSansSC-Regular.otf",
    ]
    return urls

def try_download_font():
    global CHINESE_FONT_PATH
    
    # 检查本地字体
    local_dir = os.path.dirname(__file__) if __file__ else '.'
    for fname in ['NotoSansSC.ttf', 'NotoSansSC.otf', 'SourceHanSans.ttf', 'simsun.ttc']:
        fpath = os.path.join(local_dir, fname)
        if os.path.exists(fpath) and os.path.getsize(fpath) > 100000:  # 至少 100KB
            CHINESE_FONT_PATH = fpath
            print(f"使用本地字体: {fpath}")
            return fpath
    
    # 尝试下载
    for url in get_font_download_url():
        try:
            print(f"尝试下载字体: {url}")
            ctx = ssl.create_default_context()
            ctx.check_hostname = False
            ctx.verify_mode = ssl.CERT_NONE
            
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, context=ctx, timeout=30) as resp:
                data = resp.read()
                if len(data) > 500000:  # 字体文件应该 > 500KB
                    with tempfile.NamedTemporaryFile(suffix='.otf', delete=False) as f:
                        f.write(data)
                        CHINESE_FONT_PATH = f.name
                    print(f"字体下载成功: {len(data)} bytes -> {CHINESE_FONT_PATH}")
                    return CHINESE_FONT_PATH
        except Exception as e:
            print(f"下载失败: {e}")
    
    return None

def get_pil_font(size=20):
    """获取 PIL 可用的中文字体"""
    global CHINESE_FONT_PATH
    
    if CHINESE_FONT_PATH and os.path.exists(CHINESE_FONT_PATH):
        try:
            return ImageFont.truetype(CHINESE_FONT_PATH, size)
        except:
            pass
    
    # 尝试系统字体
    system_fonts = [
        '/usr/share/fonts/truetype/wqy/wqy-microhei.ttc',
        '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',
    ]
    for fp in system_fonts:
        if os.path.exists(fp):
            try:
                return ImageFont.truetype(fp, size)
            except:
                pass
    
    # 最后用默认字体
    return ImageFont.load_default()

def draw_chinese_on_image(img, text, position, font_size=20, color=(0, 0, 0)):
    """在图片上绘制中文"""
    draw = ImageDraw.Draw(img)
    font = get_pil_font(font_size)
    draw.text(position, text, fill=color, font=font)
    return img

# 初始化字体
font_loaded = try_download_font()
if font_loaded:
    plt.rcParams['font.sans-serif'] = ['Noto Sans CJK SC', 'DejaVu Sans']
else:
    plt.rcParams['font.sans-serif'] = ['DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False



# 测试字体是否可用
def test_font():
    try:
        font = get_pil_font(12)
        # 尝试渲染一个测试字符
        test_img = Image.new('RGB', (50, 20), color='white')
        draw = ImageDraw.Draw(test_img)
        draw.text((0, 0), '测', fill='black', font=font)
        return True
    except:
        return False

st.set_page_config(page_title="历次成绩分析(通用版)", page_icon="📊", layout="wide")

def fmt(v):
    """数值取整"""
    import math
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return '-'
    if isinstance(v, float):
        return int(round(v))
    if isinstance(v, str):
        try:
            return int(round(float(v)))
        except:
            return v
    return v

def extract_sort_number(filename):
    """从文件名提取排序序号"""
    import re
    match = re.match(r'^(\d+)[-_.]', filename)
    if match:
        return int(match.group(1))
    return 999

def auto_detect_columns(df):
    """自动检测列映射"""
    keyword_map = {
        '姓名': '姓名', 'name': '姓名',
        '语文': '语文', 'chinese': '语文',
        '数学': '数学', 'math': '数学',
        '外语': '外语', '英语': '外语', 'english': '外语',
        '语数外': '语数外', '小三门': '语数外',
        '历史': '历史', 'history': '历史',
        '政治': '政治', 'politics': '政治', '道法': '政治',
        '地理': '地理', 'geography': '地理',
        '总分': '总分', 'total': '总分',
        '班排': '班排', '班级排名': '班排',
        '级排': '级排', '年级排名': '级排', 'rank': '级排',
    }
    
    for row_idx in range(min(10, len(df))):
        row = df.iloc[row_idx]
        detected = {}
        
        for col_idx, val in enumerate(row):
            if val is not None:
                val_str = str(val).strip()
                for kw, std_name in keyword_map.items():
                    if kw in val_str:
                        detected[std_name] = col_idx
                        break
        
        if len(detected) >= 3:
            return detected
    
    return None

def detect_subjects(df, col_mapping):
    """自动检测所有科目"""
    # 常见的科目名称
    subject_keywords = [
        '语文', '数学', '外语', '英语', '历史', '政治', '道法', '地理',
        '物理', '化学', '生物', '技术', '信息', '体育', '音乐', '美术',
        '语数外', '小三门', '理综', '文综', '综合'
    ]
    
    subjects = []
    # 从列映射中找科目
    for col_name, col_idx in col_mapping.items():
        if col_name not in ['姓名', '总分', '班排', '级排']:
            if col_idx < len(df.columns):
                subjects.append((col_name, col_idx))
    
    # 如果没找到，去扫描所有列
    if not subjects:
        for col_idx in range(len(df.columns)):
            if col_idx not in col_mapping.values():
                # 尝试获取这一列的表头名
                header_val = df.iloc[0].iloc[col_idx] if len(df) > 0 else None
                if header_val and str(header_val).strip():
                    subjects.append((str(header_val).strip(), col_idx))
    
    # 按列索引排序
    subjects.sort(key=lambda x: x[1])
    return subjects

# 默认列映射
DEFAULT_COLUMNS = {
    '姓名': 2,
    '语文': 3,
    '数学': 4,
    '外语': 5,
    '语数外': 6,
    '历史': 7,
    '政治': 8,
    '地理': 9,
    '总分': 10,
    '班排': 11,
    '级排': 12,
}

st.title("📊 历次成绩分析工具（通用版）")

st.markdown("---")

st.markdown("""
### 📋 使用说明

1. **上传Excel文件**：可一次性上传多个文件
2. **文件命名规则**：序号-考试名称，如：
   - `1-8月月考.xlsx`
   - `2-11月统考.xlsx`  
   - `3-期末统考.xlsx`
3. **自动检测科目**：系统会自动识别表格中的所有科目
4. **生成成绩图**：显示表格和趋势图
""")

# 文件上传
uploaded_files = st.file_uploader("上传成绩Excel文件（可多选）", type=['xlsx', 'xls'], accept_multiple_files=True)

if uploaded_files:
    try:
        st.subheader("📋 数据核对")
        
        # 读取第一个文件预览
        df_preview = pd.read_excel(uploaded_files[0], sheet_name=0, header=None, nrows=10)
        st.write("文件结构预览（前10行）：")
        st.dataframe(df_preview)
        
        # 自动检测列映射
        detected_mapping = auto_detect_columns(df_preview)
        
        if detected_mapping:
            st.success("✅ 自动检测到列映射：")
            col_mapping = detected_mapping
            for col_name, idx in sorted(col_mapping.items(), key=lambda x: x[1]):
                st.write(f"  {col_name} → 第{idx}列")
        else:
            st.warning("⚠️ 使用默认列映射")
            col_mapping = DEFAULT_COLUMNS.copy()
        
        # 检测科目
        subjects = detect_subjects(df_preview, col_mapping)
        st.info(f"📚 检测到 {len(subjects)} 个科目: {[s[0] for s in subjects]}")
        
        # 加载数据
        with st.spinner("正在加载数据..."):
            sorted_files = sorted(uploaded_files, key=lambda x: extract_sort_number(x.name))
            
            all_students = {}
            
            for uploaded_file in sorted_files:
                df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
                df = df.astype(object)
                
                # 提取考试名称
                exam_name = uploaded_file.name
                exam_name = exam_name.replace('.xlsx', '').replace('.xls', '')
                if '-' in exam_name:
                    parts = exam_name.split('-', 1)
                    if len(parts) > 1 and parts[0].isdigit():
                        exam_name = parts[1]
                
                # 简化考试名称
                import re
                name_mapping = {
                    '8月月考': '8月月考', '11月统考': '11月统考', '期末统考': '期末考试',
                    '深一模': '一模', '高二期末': '高二期末'
                }
                simplified = exam_name
                for orig, short in name_mapping.items():
                    if orig in exam_name or exam_name in orig:
                        simplified = short
                        break
                if simplified == exam_name:
                    simplified = re.sub(r'考试|统考|月考|较高二|较\d+月', '', exam_name)
                    if not simplified:
                        simplified = exam_name
                exam_name = simplified
                
                name_col = col_mapping.get('姓名', 2)
                
                for _, row in df.iterrows():
                    name = row.get(name_col)
                    if name and isinstance(name, str) and name != '姓名':
                        if name not in all_students:
                            all_students[name] = []
                        
                        student_data = {'exam': exam_name}
                        
                        # 添加检测到的科目
                        for subj_name, subj_idx in subjects:
                            student_data[subj_name] = row.get(subj_idx)
                        
                        # 添加总分和排名
                        if '总分' in col_mapping:
                            student_data['total'] = row.get(col_mapping['总分'])
                        if '班排' in col_mapping:
                            student_data['class_rank'] = row.get(col_mapping['班排'])
                        if '级排' in col_mapping:
                            student_data['rank'] = row.get(col_mapping['级排'])
                        
                        all_students[name].append(student_data)
        
        students = sorted(all_students.keys())
        st.success(f"成功读取 {len(uploaded_files)} 个文件，共 {len(students)} 位学生")
        
        sorted_file_names = [f.name for f in sorted_files]
        st.info(f"文件顺序: {' → '.join(sorted_file_names)}")
        
        option = st.radio("选择生成方式", ["生成全部", "选择学生"])
        
        def safe_sub(a, b):
            try:
                import math
                if a is None or b is None or (isinstance(a, float) and math.isnan(a)) or (isinstance(b, float) and math.isnan(b)):
                    return None
                a_num = float(a) if not isinstance(a, (int, float)) else a
                b_num = float(b) if not isinstance(b, (int, float)) else b
                return a_num - b_num
            except:
                return None
        
        def safe_float(v, default=0):
            import math
            if v is None or (isinstance(v, float) and math.isnan(v)):
                return default
            try:
                return float(v)
            except:
                return default
        
        def generate_student_figure(name, student_data, subjects_list):
            try:
                student_data = sorted(student_data, key=lambda x: extract_sort_number(x.get('exam', '999')))
                
                table_rows = []
                
                for i, data in enumerate(student_data):
                    row_data = {'exam': data['exam'], 'is_diff': False}
                    for subj_name, _ in subjects_list:
                        row_data[subj_name] = data.get(subj_name)
                    row_data['total'] = data.get('total')
                    row_data['class_rank'] = data.get('class_rank')
                    row_data['rank'] = data.get('rank')
                    table_rows.append(row_data)
                    
                    if i > 0:
                        prev = student_data[i - 1]
                        curr = data
                        diff_data = {'exam': f"{curr['exam']}较{prev['exam']}", 'is_diff': True}
                        
                        for subj_name, _ in subjects_list:
                            diff_data[subj_name] = fmt(safe_sub(curr.get(subj_name), prev.get(subj_name)))
                        
                        diff_data['total'] = fmt(safe_sub(curr.get('total'), prev.get('total')))
                        diff_data['class_rank'] = fmt(safe_sub(prev.get('class_rank'), curr.get('class_rank')))
                        diff_data['rank'] = fmt(safe_sub(prev.get('rank'), curr.get('rank')))
                        
                        table_rows.append(diff_data)
                
                exam_data = [r for r in table_rows if not r['is_diff']]
                exams = [d['exam'] for d in exam_data]
                
                # 构建表头
                header = ['考试'] + [s[0] for s in subjects_list] + ['总分', '班排', '级排']
                header = [h for h in header if h]
                
                # 准备表格数据
                table_data = []
                for row in table_rows:
                    r = [row['exam']]
                    for subj_name, _ in subjects_list:
                        r.append(fmt(row.get(subj_name)))
                    r.append(fmt(row.get('total')))
                    r.append(fmt(row.get('class_rank')))
                    r.append(fmt(row.get('rank')))
                    table_data.append(r)
                
                # ========== 使用 Plotly 生成图表 ==========
                totals = [safe_float(d.get('total'), 0) for d in exam_data]
                ranks = [safe_float(d.get('rank'), 0) for d in exam_data]
                
                from plotly.subplots import make_subplots
                import plotly.graph_objects as go
                
                # 创建图表
                fig = make_subplots(
                    rows=2, cols=1,
                    row_heights=[0.35, 0.65],
                    subplot_titles=('<b>总分与级排趋势</b>', '<b>各科成绩趋势</b>'),
                    vertical_spacing=0.12
                )
                
                # 总分折线
                fig.add_trace(
                    go.Scatter(x=exams, y=totals, name='总分', mode='lines+markers',
                              line=dict(color='#1f77b4', width=3), marker=dict(size=10)),
                    row=1, col=1
                )
                
                # 级排折线（反向）
                if ranks and max(ranks) > 0:
                    fig.add_trace(
                        go.Scatter(x=exams, y=ranks, name='级排', mode='lines+markers',
                                  line=dict(color='#ff7f0e', width=3), marker=dict(size=10),
                                  yaxis='y2'),
                        row=1, col=1
                    )
                
                # 各科折线
                colors = ['#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
                for idx, (subj_name, _) in enumerate(subjects_list):
                    data = [int(safe_float(d.get(subj_name), 0)) for d in exam_data]
                    color = colors[idx % len(colors)]
                    fig.add_trace(
                        go.Scatter(x=exams, y=data, name=subj_name, mode='lines+markers',
                                  line=dict(color=color, width=2), marker=dict(size=8)),
                        row=2, col=1
                    )
                
                # 布局设置
                fig.update_layout(
                    title=dict(text=f'<b>{name} 历次成绩</b>', font=dict(size=20)),
                    height=700,
                    showlegend=True,
                    legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
                    font=dict(size=12),
                    paper_bgcolor='white',
                    plot_bgcolor='white'
                )
                
                # Y轴设置
                fig.update_yaxes(title_text='分数', row=2, col=1, range=[0, 150])
                if ranks and max(ranks) > 0:
                    fig.update_yaxes(title_text='级排', row=1, col=1, secondary_y=True, 
                                   range=[max(ranks) + 20, max(1, min(ranks) - 20)],
                                   autorange='reversed')
                else:
                    fig.update_yaxes(title_text='级排', row=1, col=1, secondary_y=True, showgrid=False)
                
                # X轴旋转
                fig.update_xaxes(tickangle=30)
                
                # 在Streamlit中显示图表
                st.plotly_chart(fig, use_container_width=True)
                
                # 返回None，因为图表直接显示了
                return None
            except Exception as e:
                st.error(f"生成图片时出错: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
                return None
        
        def create_pdf_from_figures(figures):
            """使用kaleido将Plotly图表导出为PDF"""
            pdf_buffer = io.BytesIO()
            try:
                import kaleido
                images = []
                for fig in figures:
                    img_bytes = fig.to_image(format='png', width=1200, height=800, scale=2)
                    img = Image.open(io.BytesIO(img_bytes))
                    if img.mode != 'RGB':
                        img = img.convert('RGB')
                    images.append(img)
                
                if images:
                    images[0].save(pdf_buffer, save_all=True, append_images=images[1:], format='PDF')
                pdf_buffer.seek(0)
                return pdf_buffer
            except Exception as e:
                st.warning(f"PDF生成失败（kaleido问题）: {e}")
                return None
        
        if option == "生成全部":
            if st.button("生成全部学生成绩图"):
                with st.spinner("正在生成..."):
                    figures = []
                    progress_bar = st.progress(0)
                    
                    for i, name in enumerate(students):
                        # 生成图表但不显示
                        try:
                            fig = generate_student_figure(name, all_students[name], subjects)
                            if fig:
                                figures.append((name, fig))
                        except Exception as e:
                            st.error(f"生成 {name} 图表时出错: {e}")
                        progress_bar.progress((i + 1) / len(students))
                    
                    if figures:
                        # 尝试生成PDF
                        pdf_buffer = create_pdf_from_figures([f for _, f in figures])
                        
                        st.success(f"生成完成！共 {len(figures)} 位学生")
                        
                        if pdf_buffer:
                            st.download_button(
                                label="📥 下载合并PDF",
                                data=pdf_buffer,
                                file_name="历次成绩_合并.pdf",
                                mime="application/pdf"
                            )
                        
                        # 预览前几个
                        st.markdown("### 预览")
                        for name, fig in figures[:3]:
                            st.plotly_chart(fig, use_container_width=True)
                        if len(figures) > 3:
                            st.info(f"还有 {len(figures) - 3} 位学生...")
        else:
            selected = st.selectbox("选择学生", students)
            
            if st.button("生成"):
                with st.spinner("正在生成..."):
                    fig = generate_student_figure(selected, all_students[selected], subjects)
                    
                    if fig:
                        st.success("生成完成！")
                        st.plotly_chart(fig, use_container_width=True)
                        
                        # 尝试生成PNG下载
                        try:
                            import kaleido
                            img_bytes = fig.to_image(format='png', width=1200, height=800, scale=2)
                            st.download_button(
                                label="📥 下载图片",
                                data=img_bytes,
                                file_name=f"{selected} 历次成绩.png",
                                mime="image/png"
                            )
                        except Exception as e:
                            st.info("图片下载需要 kaleido，请确保已安装: pip install kaleido")
                    else:
                        st.error("生成失败")
                        
    except Exception as e:
        st.error(f"出错: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
