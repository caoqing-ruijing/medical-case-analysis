import pandas as pd
import os
import re
import glob

# 定义输出文件路径
OUTPUT_FILE = "matched_prompts_output.xlsx"

# 读取Excel文件
def load_excel_data():
    try:
        df = pd.read_excel('END.xlsx')
        print(f"成功加载Excel文件，共{len(df)}条数据")
        return df
    except Exception as e:
        print(f"读取Excel文件失败: {str(e)}")
        return None

# 读取Python文件中的prompt_template
def extract_prompt_template(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            # 使用正则表达式提取prompt_template
            match = re.search(r"prompt_template\s*=\s*'''([\s\S]*?)'''\s*", content)
            if match:
                return match.group(1).strip()
            else:
                print(f"在{file_path}中未找到prompt_template")
                return None
    except Exception as e:
        print(f"读取{file_path}失败: {str(e)}")
        return None

# 分析prompt_template中的占位符
def analyze_template(template):
    if not template:
        return []
    # 查找所有{{XXX}}格式的占位符
    placeholders = re.findall(r'\{\{([^\}]+)\}\}', template)
    return placeholders

# 检查Excel数据中是否包含所需的字段
def check_excel_fields(df, required_fields):
    missing_fields = []
    for field in required_fields:
        if field not in df.columns and field != 'GENDER' and field != 'ANGIOGRAPHY_RESULTS' and field != 'CORONARY_CTA':
            # 特殊处理一些字段映射
            if field == 'GENDER' and 'SEX' in df.columns:
                continue
            if field == 'ANGIOGRAPHY_RESULTS' and 'CAG' in df.columns:
                continue
            if field == 'CORONARY_CTA' and 'CTA' in df.columns:
                continue
            missing_fields.append(field)
    return missing_fields

# 将Excel数据填充到prompt模板中
def fill_template(template, row):
    filled_template = template
    
    # 处理常见的字段映射
    mappings = {
        'GENDER': 'SEX',
        'ANGIOGRAPHY_RESULTS': 'CAG',
        'CORONARY_CTA': 'CTA'
    }
    
    # 替换所有占位符
    for placeholder in analyze_template(template):
        field = placeholder
        # 检查是否需要映射字段名
        if field in mappings and mappings[field] in row:
            field = mappings[field]
        
        # 获取字段值，如果不存在则使用'未知'
        value = str(row.get(field, '未知'))
        # 对于长文本字段，限制长度
        if field in ['CAG', 'CTA', 'TTE', 'ECG']:
            max_length = 2000 if field in ['CAG', 'CTA'] else 1000
            value = value[:max_length]
        
        # 替换占位符
        filled_template = filled_template.replace(f'{{{{{placeholder}}}}}', value)
    
    return filled_template

# 主函数
def main():
    # 加载Excel数据
    df = load_excel_data()
    if df is None:
        return
    
    # 获取所有Python文件
    py_files = glob.glob("*.py")
    py_files = [f for f in py_files if f != "match_excel_to_prompts.py" and f != "read_excel.py"]
    
    if not py_files:
        print("未找到Python文件")
        return
    
    print(f"找到{len(py_files)}个Python文件: {', '.join(py_files)}")
    
    # 存储结果
    results = []
    
    # 处理每个Python文件
    for i, py_file in enumerate(py_files):
        print(f"\n处理文件 {i+1}/{len(py_files)}: {py_file}")
        
        # 提取prompt模板
        template = extract_prompt_template(py_file)
        if not template:
            continue
        
        # 分析模板中的占位符
        placeholders = analyze_template(template)
        print(f"模板中的占位符: {', '.join(placeholders)}")
        
        # 检查Excel数据中是否包含所需的字段
        missing_fields = check_excel_fields(df, placeholders)
        if missing_fields:
            print(f"警告: Excel中缺少字段: {', '.join(missing_fields)}")
        
        # 为每行Excel数据填充模板
        for idx, row in df.iterrows():
            filled_template = fill_template(template, row)
            
            # 将结果添加到列表中
            results.append({
                "文件名": py_file,
                "序号": row.get("序号", idx),
                "patientID": row.get("patientID", ""),
                "性别": row.get("SEX", ""),
                "年龄": row.get("AGE", ""),
                "手术日期": row.get("DAY", ""),
                "填充后的Prompt": filled_template
            })
    
    # 将结果保存到Excel文件
    if results:
        result_df = pd.DataFrame(results)
        result_df.to_excel(OUTPUT_FILE, index=False)
        print(f"\n处理完成，结果已保存到 {OUTPUT_FILE}")
    else:
        print("没有生成任何结果")

if __name__ == "__main__":
    main()