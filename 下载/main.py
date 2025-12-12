import os
import pandas as pd
import re

# ================= 配置区域 =================
folder_path = r"C:\Users\Administrator\Downloads\微信公众号批量下载工具4.0\下载\开源内核安全修炼"
csv_path = r"C:\Users\Administrator\Downloads\微信公众号批量下载工具4.0\下载\2025-12-12-143930.csv"
output_path = r"output.xlsx"

MATCH_LEN = 30  # 截取前30字符
# ===========================================

def clean_text_for_match(text):
    if not isinstance(text, str):
        return ""
    text = re.sub(r'^\[.*?\]', '', text) # 去日期前缀
    text = re.sub(r'\.md$', '', text, flags=re.IGNORECASE) # 去后缀
    text = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]', '', text) # 去符号
    return text

def main():
    print("=== 步骤 1: 读取 CSV 并去重 ===")
    try:
        try:
            df_csv = pd.read_csv(csv_path, header=0, encoding='utf-8')
        except UnicodeDecodeError:
            df_csv = pd.read_csv(csv_path, header=0, encoding='gbk')

        # 确定标题列
        target_col = None
        if '标题' in df_csv.columns: target_col = '标题'
        elif 'Title' in df_csv.columns: target_col = 'Title'
        else: target_col = df_csv.columns[1]

        # CSV 去重
        df_csv.drop_duplicates(subset=[target_col], keep='first', inplace=True)
        print(f"CSV 准备就绪，共 {len(df_csv)} 行。")

    except Exception as e:
        print(f"CSV 读取失败: {e}")
        return

    print(f"\n=== 步骤 2: 扫描 Markdown 文件 (Key冲突检测) ===")
    
    # 逻辑修改核心：
    # 1. all_md_files: 记录所有存在的文件名 (Set)
    # 2. key_map: { '截断后的Key': ['文件名A', '文件名B'] } -> 处理Key冲突
    # 3. file_content_map: { '文件名': '内容' } -> 根据文件名取内容
    
    all_md_files = set()
    key_map = {} 
    file_content_map = {}

    if os.path.exists(folder_path):
        files = [f for f in os.listdir(folder_path) if f.lower().endswith('.md')]
        
        for filename in files:
            all_md_files.add(filename)
            
            # 读取内容
            try:
                with open(os.path.join(folder_path, filename), 'r', encoding='utf-8') as f:
                    content = f.read()
                file_content_map[filename] = content
            except Exception as e:
                print(f"读取失败: {filename}")
                continue

            # 生成 Key
            clean_name = clean_text_for_match(filename)
            short_key = clean_name[:MATCH_LEN]
            
            # 存入 Key Map (注意：这里用列表存储，防止覆盖！)
            if short_key not in key_map:
                key_map[short_key] = []
            key_map[short_key].append(filename)
            
    else:
        print("文件夹不存在")
        return

    print(f"物理文件共: {len(all_md_files)} 个")
    print(f"生成的唯一 Key 共: {len(key_map)} 个 (如果有文件Key重复，Key数量会少于文件数)")

    print("\n=== 步骤 3: 严格匹配 ===")
    
    final_contents = []
    final_filenames = []
    
    matched_files_set = set() # 记录哪几个具体的文件被用掉了
    csv_unmatched_indices = []

    for index, row in df_csv.iterrows():
        original_title = str(row[target_col])
        csv_key = clean_text_for_match(original_title)[:MATCH_LEN]
        
        # 查找匹配
        if csv_key in key_map:
            # 获取对应的文件列表
            candidate_files = key_map[csv_key]
            
            # 默认取第一个匹配到的文件
            chosen_file = candidate_files[0]
            
            # 记录数据
            final_contents.append(file_content_map[chosen_file])
            final_filenames.append(chosen_file)
            
            # 标记该文件已被使用
            matched_files_set.add(chosen_file)
            
            # 如果一个 Key 对应多个文件，且这是第一次发现，打印警告
            if len(candidate_files) > 1:
                # 这里可以加个逻辑，比如打印出来“警告：CSV标题对应了多个MD文件”
                pass
        else:
            final_contents.append("")
            final_filenames.append("")
            csv_unmatched_indices.append(f"行[{index+2}]标题: {original_title}")

    # 将数据写入 DataFrame
    df_csv['Markdown文件名'] = final_filenames
    df_csv['Markdown内容'] = final_contents

    # === 统计计算 ===
    # 真正的未使用 = 所有物理文件 - 匹配成功被记录的文件
    unused_files = all_md_files - matched_files_set
    
    print("\n" + "="*40)
    print(f" CSV 行数: {len(df_csv)}")
    print(f" MD 文件数: {len(all_md_files)}")
    print(f" 成功匹配: {len(matched_files_set)} 个 MD 文件")
    print("="*40)

    # 1. 打印 CSV 未匹配
    if csv_unmatched_indices:
        print(f"\n[CSV 未匹配] 共有 {len(csv_unmatched_indices)} 行没有找到 MD 文件:")
        for info in csv_unmatched_indices[:10]:
            print(info)
        if len(csv_unmatched_indices) > 10: print("...")
    else:
        print("\n[CSV 匹配] 所有 CSV 行都找到了对应的文件。")

    # 2. 打印 MD 未被使用 (这才是你要的！)
    if unused_files:
        print(f"\n[MD 未使用] 共有 {len(unused_files)} 个 MD 文件未被 CSV 引用:")
        print("(原因可能是：CSV里没有这篇，或者由于 MATCH_LEN 太短，它被同名的另一个文件抢先匹配了)")
        for f in list(unused_files):
            print(f" - {f}")
    else:
        print("\n[MD 匹配] 所有 MD 文件都被用上了。")

    # 保存
    df_csv.to_excel(output_path, index=False)
    print(f"\n结果已保存: {output_path}")

if __name__ == "__main__":
    main()