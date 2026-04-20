import streamlit as st
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')
import numpy as np
import os
import io
from PIL import Image
import matplotlib.font_manager as fm

# 尝试加载中文字体 - 兼容本地和Streamlit Cloud
def setup_chinese_font():
    # 本地字体路径
    local_fonts = [
        os.path.join(os.path.dirname(__file__), f) 
        for f in os.listdir(os.path.dirname(__file__)) 
        if f.endswith(('.ttf', '.otf'))
    ]
    
    chinese_font = None
    
    # 先尝试本地字体
    for fp in local_fonts:
        try:
            fm.fontManager.addfont(fp)
            chinese_font = fm.FontProperties(fname=fp).get_name()
            print(f"使用本地字体: {fp} -> {chinese_font}")
            return chinese_font
        except Exception as e:
            print(f"字体 {fp} 加载失败: {e}")
    
    # 搜索系统字体
    for f in fm.findSystemFonts():
        if any(kw in f.lower() for kw in ['cjk', 'noto', 'chinese', 'hans', 'wqy', 'source']):
            try:
                fm.fontManager.addfont(f)
                chinese_font = fm.FontProperties(fname=f).get_name()
                print(f"使用系统字体: {f} -> {chinese_font}")
                return chinese_font
            except:
                continue
    
    print("警告: 未找到中文字体，汉字可能显示为方框")
    return 'sans-serif'

chinese_font = setup_chinese_font()
plt.rcParams['font.sans-serif'] = [chinese_font, 'DejaVu Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

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
        
        def generate_student_image(name, student_data, subjects_list):
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
                
                # 生成图片
                fig = plt.figure(figsize=(8.27, 11.69))
                fig.subplots_adjust(left=0.04, right=0.96, top=0.93, bottom=0.05)
                
                fig.suptitle(f'{name} 历次成绩', fontsize=20, fontweight='bold', y=0.95)
                
                # 上部：表格
                ax_table = fig.add_axes([0.05, 0.55, 0.9, 0.40])
                ax_table.axis('off')
                
                table = ax_table.table(cellText=table_data, colLabels=header, loc='center', cellLoc='center')
                table.auto_set_font_size(False)
                table.set_fontsize(9)
                table.auto_set_column_width([0])
                table.scale(1.0, 1.4)
                
                for i in range(len(header)):
                    table[(0, i)].set_facecolor('#4472C4')
                    table[(0, i)].set_text_props(color='white', fontweight='bold', fontsize=9)
                
                for i in range(1, len(table_data) + 1):
                    is_diff = '较' in str(table_data[i-1][0])
                    for j in range(len(header)):
                        if is_diff:
                            table[(i, j)].set_facecolor('#FFE6CC')
                        elif i % 2 == 0:
                            table[(i, j)].set_facecolor('#FFFFFF')
                        else:
                            table[(i, j)].set_facecolor('#CCE5FF')
                        
                        val = table_data[i-1][j]
                        if isinstance(val, (int, float)) and val < 0:
                            table[(i, j)].set_text_props(color='red', fontsize=9)
                        else:
                            table[(i, j)].set_text_props(color='black', fontsize=9)
                
                # 中部：总分和级排折线图
                ranks = [safe_float(d.get('rank'), 0) for d in exam_data]
                if ranks and max(ranks) > 0:
                    min_rank, max_rank = min(ranks), max(ranks)
                    rank_min, rank_max = max(1, min_rank - 20), max_rank + 20
                else:
                    rank_min, rank_max = 1, 100
                
                ax_middle = fig.add_axes([0.1, 0.35, 0.8, 0.18])
                ax_middle_twin = ax_middle.twinx()
                
                totals = [safe_float(d.get('total'), 0) for d in exam_data]
                line1, = ax_middle.plot(exams, totals, color='#1f77b4', marker='o', linewidth=2, markersize=8)
                ax_middle.set_ylabel('总分', color='#1f77b4', fontsize=12, fontweight='bold')
                ax_middle.tick_params(axis='y', labelcolor='#1f77b4')
                ax_middle.set_ylim(50, max(max(totals) + 50, 550) if totals else 550)
                
                line2, = ax_middle_twin.plot(exams, ranks, color='#ff7f0e', marker='s', linewidth=2, markersize=8)
                ax_middle_twin.set_ylabel('级排', color='#ff7f0e', fontsize=12, fontweight='bold')
                ax_middle_twin.tick_params(axis='y', labelcolor='#ff7f0e')
                ax_middle_twin.set_ylim(rank_min, rank_max)
                ax_middle_twin.invert_yaxis()
                
                ax_middle.set_title('总分与级排趋势', fontsize=14, fontweight='bold', pad=10)
                ax_middle.tick_params(axis='x', labelsize=10, rotation=30)
                ax_middle.grid(True, alpha=0.3)
                ax_middle.set_xlabel('')
                
                if exams and totals:
                    ax_middle.annotate('总分', (exams[-1], totals[-1]), textcoords="offset points", 
                                       xytext=(5, 0), ha='left', fontsize=9, color='#1f77b4', fontweight='bold')
                if exams and ranks and max(ranks) > 0:
                    ax_middle_twin.annotate('级排', (exams[-1], ranks[-1]), textcoords="offset points", 
                                            xytext=(5, 0), ha='left', fontsize=9, color='#ff7f0e', fontweight='bold')
                
                # 下部：各科折线图
                ax_bottom = fig.add_axes([0.1, 0.02, 0.8, 0.25])
                
                colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', 
                          '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
                
                for idx, (subj_name, _) in enumerate(subjects_list):
                    data = [int(safe_float(d.get(subj_name), 0)) for d in exam_data]
                    color = colors[idx % len(colors)]
                    ax_bottom.plot(exams, data, marker='o', linewidth=2, markersize=6, color=color, label=subj_name)
                    if exams and data:
                        ax_bottom.annotate(subj_name, (exams[-1], data[-1]), textcoords="offset points", 
                                         xytext=(5, 0), ha='left', fontsize=9, color=color, fontweight='bold')
                
                ax_bottom.set_ylabel('分数', fontsize=11)
                ax_bottom.set_title('各科成绩趋势', fontsize=12, fontweight='bold', pad=5)
                ax_bottom.set_ylim(0, 150)
                ax_bottom.tick_params(axis='x', labelsize=9, rotation=30)
                ax_bottom.grid(True, alpha=0.3)
                
                img_buffer = io.BytesIO()
                plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight', facecolor='white')
                plt.close()
                img_buffer.seek(0)
                
                return img_buffer
            except Exception as e:
                st.error(f"生成图片时出错: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
                return None
        
        def create_pdf(images):
            pdf_buffer = io.BytesIO()
            pil_images = []
            for img_data in images:
                img = Image.open(img_data)
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                pil_images.append(img)
            if pil_images:
                pil_images[0].save(pdf_buffer, save_all=True, append_images=pil_images[1:], format='PDF')
            pdf_buffer.seek(0)
            return pdf_buffer
        
        if option == "生成全部":
            if st.button("生成全部学生成绩图"):
                with st.spinner("正在生成..."):
                    images = []
                    progress_bar = st.progress(0)
                    
                    for i, name in enumerate(students):
                        img_buffer = generate_student_image(name, all_students[name], subjects)
                        if img_buffer:
                            images.append((name, img_buffer))
                        progress_bar.progress((i + 1) / len(students))
                    
                    if images:
                        pdf_buffer = create_pdf([img for _, img in images])
                        st.success(f"生成完成！共 {len(images)} 位学生")
                        
                        st.download_button(
                            label="📥 下载合并PDF",
                            data=pdf_buffer,
                            file_name="历次成绩_合并.pdf",
                            mime="application/pdf"
                        )
                        
                        st.markdown("### 预览")
                        for name, img_buffer in images[:3]:
                            st.image(img_buffer, caption=name, width=600)
                        if len(images) > 3:
                            st.info(f"还有 {len(images) - 3} 位学生...")
        else:
            selected = st.selectbox("选择学生", students)
            
            if st.button("生成"):
                with st.spinner("正在生成..."):
                    img_buffer = generate_student_image(selected, all_students[selected], subjects)
                    
                    if img_buffer:
                        st.success("生成完成！")
                        st.image(img_buffer, caption=selected, width=600)
                        
                        st.download_button(
                            label="📥 下载图片",
                            data=img_buffer,
                            file_name=f"{selected} 历次成绩.png",
                            mime="image/png"
                        )
                    else:
                        st.error("生成失败")
                        
    except Exception as e:
        st.error(f"出错: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
