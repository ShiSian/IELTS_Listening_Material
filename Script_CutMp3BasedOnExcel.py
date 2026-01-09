import os
import glob
import openpyxl
from pydub import AudioSegment, silence
from openpyxl.utils import column_index_from_string
import win32com.client as win32  # 用于调用 Excel 引擎



# ================= 导出指定Sheet的B列数据（仅导出未隐藏的行） ====================
def export_specific_sheets(excel_path, target_sheets):
    # 1. 检查文件是否存在
    if not os.path.exists(excel_path):
        print(f"错误：找不到文件 '{excel_path}'")
        return

    print(f"正在加载 Excel 文件: {excel_path} ... (文件较大时可能需要几秒)")

    # 2. 加载工作簿
    # data_only=True 确保读取的是计算后的数值而不是公式
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception as e:
        print(f"无法打开Excel文件: {e}")
        return

    # 3. 遍历指定的 Sheet
    for sheet_name in target_sheets:
        if sheet_name not in wb.sheetnames:
            print(f"⚠️  警告：Excel中找不到名为 '{sheet_name}' 的Sheet，已跳过。")
            continue

        ws = wb[sheet_name]
        exported_data = []

        print(f"正在处理 Sheet: {sheet_name} ...")

        # 4. 遍历行 (从第3行开始)
        # openpyxl 的 max_row 会获取已使用的最大行数
        for row_idx in range(3, ws.max_row + 1):

            # --- 核心逻辑：检查行是否被隐藏 ---
            # ws.row_dimensions 存储了行的属性
            # 如果 row_idx 在 row_dimensions 中且 hidden 为 True，则跳过
            if row_idx in ws.row_dimensions and ws.row_dimensions[row_idx].hidden:
                # 这一行是隐藏的，跳过
                continue

            # --- 读取 B 列数据 (Column 2) ---
            cell_value = ws.cell(row=row_idx, column=2).value

            # 如果单元格不为空，则加入列表
            # 这里使用了 str() 确保数字也能被写入 txt
            if cell_value is not None:
                exported_data.append(str(cell_value))

        # 5. 写入 TXT 文件
        if exported_data:
            txt_filename = f"Keep_{sheet_name}.txt"
            try:
                with open(txt_filename, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(exported_data))
                print(f"✅ 成功导出: {txt_filename} (共 {len(exported_data)} 行数据)")
            except Exception as e:
                print(f"❌ 写入文件 {txt_filename} 失败: {e}")
        else:
            print(f"⚠️  Sheet '{sheet_name}' 没有可导出的有效数据（A3之后为空或全为隐藏行）。")

    print("-" * 30)
    print("所有任务完成。")
     
def load_list(file_path):
    if not os.path.exists(file_path):
        return None
    with open(file_path, 'r', encoding='utf-8') as f:
        # 使用 strip() 去除换行符和首尾空格，并过滤空行
        return [line.strip() for line in f if line.strip()]

def process_single_unit(mp3_path):
    base_name = os.path.splitext(mp3_path)[0]
    full_txt =   f"Origin_{base_name}.txt"
    keep_txt =   f"Keep_{base_name}.txt"
    output_mp3 = f"Cutted_{base_name}.mp3"

    if not os.path.exists(full_txt):
        print(f"'{full_txt}' 不存在，跳过音频切割")
        return

    if not os.path.exists(keep_txt):
        print(f"'{keep_txt}' 不存在，跳过音频切割")
        return

    print(f"\n>>> 正在处理: {mp3_path}")
    full_list = load_list(full_txt)
    keep_list = load_list(keep_txt)

    audio = AudioSegment.from_mp3(mp3_path)

    # 检测静音
    nonsilent_ranges = silence.detect_nonsilent(
        audio,
        min_silence_len=800, # 保持较大值防止切断词组
        silence_thresh=-45,
        seek_step=5
    )

    print(f"  - 单词总数: {len(full_list)}")
    print(f"  - 检测片段: {len(nonsilent_ranges)}")

    start_idx = 0
    word_map = {}
    word_ranges = nonsilent_ranges[start_idx:]

    for i in range(len(full_list)):
        if i >= len(word_ranges):
            print(f"  - ❌ 警告: 音频片段耗尽，单词 '{full_list[i]}' 及其后续单词无法匹配。")
            break

        current_word = full_list[i]
        start = word_ranges[i][0]

        # 确定切片终点：默认为下一个有效片段的起点（包含停顿）
        # 如果是列表最后一个单词，或者由于杂音导致 ranges 后面还有多余的，
        # 我们这里做一个保护：如果是full_list的最后一个词，直接取到文件末尾
        if i == len(full_list) - 1:
            chunk = audio[start:]
        else:
            # 正常情况：切到下一个对应的 ranges 起点
            # 注意：这里我们用 word_ranges[i+1] 是安全的，因为我们是按 full_list 遍历的
            if i + 1 < len(word_ranges):
                next_start = word_ranges[i+1][0]
                chunk = audio[start:next_start]
            else:
                # 极端情况：单词表还没完，但音频段用完了（前面情况C已经拦截，这里是双重保险）
                chunk = audio[start:]

        word_map[current_word] = chunk

    # 合成导出
    combined = AudioSegment.empty()
    success_count = 0
    missing_words = []

    for word in keep_list:
        if word in word_map:
            combined += word_map[word]
            success_count += 1
        else:
            missing_words.append(word)

    if missing_words:
        print(f"  - ⚠️ 以下单词未找到音频: {missing_words}")

    if success_count > 0:
        combined.export(output_mp3, format="mp3")
        print(f"  - ✅ 成功导出: {output_mp3}")

def refresh_excel_formulas(file_path):
    """使用 Excel 软件打开并保存，以强制更新公式缓存"""
    abs_path = os.path.abspath(file_path)
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(abs_path)
        wb.Save()
        wb.Close()
        print(f">>> 公式缓存已刷新")
    except Exception as e:
        print(f"[Error] 刷新公式失败: {e}")
    finally:
        excel.Quit()

"""
根据规则隐藏指定Sheet的指定行
"""
def hide_completed_rows(file_name, target_sheets, target_columns):
    if not os.path.exists(file_name):
        return

    # ---------- 第一阶段：重置 ----------
    try:
        wb_reset = openpyxl.load_workbook(file_name)
        for sheet_name in target_sheets:
            if sheet_name in wb_reset.sheetnames:
                ws = wb_reset[sheet_name]
                for r in range(1, ws.max_row + 1):
                    ws.row_dimensions[r].hidden = False
                ws.auto_filter.ref = None
        wb_reset.save(file_name)
        wb_reset.close()
    except PermissionError:
        print("[Error] 文件被占用，无法重置。")
        return

    # ---------- 第二阶段：判定 ----------
    try:
        wb_reader = openpyxl.load_workbook(file_name, data_only=True)
        wb_writer = openpyxl.load_workbook(file_name)
    except Exception as e:
        print(f"重新加载失败: {e}")
        return

    target_col_indices = [column_index_from_string(c) for c in target_columns]

    # === 等价于 Excel 公式的 Python 计算 ===
    def calc_formula_equivalent(ws, row):
        """
        等价于：
        =IF(Gx="", "", IF(OR(Gx=Bx, Gx=Cx), "√", Bx & "|" & Cx & ">" & Dx))
        """
        def norm(v):
            return "" if v is None else str(v).strip()

        b = norm(ws[f"B{row}"].value)
        c = norm(ws[f"C{row}"].value)
        d = norm(ws[f"D{row}"].value)
        g = norm(ws[f"G{row}"].value)

        if g == "":
            return ""
        if g == b or g == c:
            return "√"
        return f"{b}|{c}>{d}"

    for sheet_name in target_sheets:
        if sheet_name not in wb_reader.sheetnames:
            continue

        ws_read = wb_reader[sheet_name]
        ws_write = wb_writer[sheet_name]

        rows_hidden_count = 0
        max_row = ws_read.max_row

        for row_idx in range(3, max_row + 1):
            match_counter = 0

            for col_idx in target_col_indices:
                cell = ws_read.cell(row=row_idx, column=col_idx)
                cell_value = cell.value

                # 如果 Excel 没有缓存公式结果 → 自己算
                if cell_value is None:
                    cell_value = calc_formula_equivalent(ws_read, row_idx)

                str_value = str(cell_value).strip() if cell_value is not None else ""

                if str_value == "√" or str_value == "":
                    match_counter += 1

            if match_counter == len(target_col_indices):
                ws_write.row_dimensions[row_idx].hidden = True
                rows_hidden_count += 1

        print(f"  -> Sheet '{sheet_name}' 处理完毕，隐藏了 {rows_hidden_count} 行。")

    try:
        wb_writer.save(file_name)
        print(">>> 任务最终完成！")
    except PermissionError:
        print("[Error] 保存最终结果失败。")


 
        
    


def main():
    # 1、隐藏不需要keep的行
    # 告诉程序操作哪个Excel文件
    my_excel_file = "王陆听力语料库.xlsx"
    # 内置的按照章节的单元表名称
    my_target_sheet03 = ["3.2","3.3-1", "3.3-2", "3.3-3", "3.3-4", "3.3-5", "3.3-6", "3.3-7", "3.3-8", "3.3-9"]
    my_target_sheet04 = ["4.2", "4.3-1", "4.3-2", "4.3-3", "4.4"]
    my_target_sheet05 = ["5.2", "5.3-1", "5.3-2", "5.3-3", "5.3-4", "5.3-5", "5.3-6", "5.3-7", "5.3-8", "5.3-9", "5.3-10", "5.3-11", "5.3-12"]
    my_target_sheet08 = ["8.2", "8.3-1", "8.3-2", "8.3-3", "8.3-4", "8.3-5", "8.4-1", "8.4-2", "8.4-3", "8.5", "8.6-1", "8.6-2", "8.6-3", "8.7-1", "8.7-2", "8.7-3", "8.8"]
    my_target_sheet11 = ["11.1", "11.2", "11.3", "11.4"]
    # 告诉程序需要处理哪些单元表【这里可能需要修改，当然如果这里把所有章节都加起来就会全量处理整个表格】
    my_target_sheets  = my_target_sheet03 + my_target_sheet04
    # 告诉程序基于哪几列的值来判断是否需要隐藏对应行（比如FGH表示如果在一行中FGH列的值都是√或者为空<表示本次不需要听写>，则隐藏该行）
    # 【这里可能需要修改】
    my_target_columns = ["F", "H"]
    # 执行隐藏操作
    hide_completed_rows(my_excel_file, my_target_sheets, my_target_columns)
    
    # 2、从Excel中导出需要keep的单词列表
    export_specific_sheets(my_excel_file, my_target_sheets)

    # 3、基于keep单词列表切割mp3
    mp3_files = glob.glob("*.mp3")
    for mp3 in mp3_files:
        # 跳过已经处理过的文件
        if "Cutted_" in mp3:
            continue
        # 检查文件名是否包含 my_target_sheets 中的任意一个章节号
        # 例如：如果 mp3 是 "3.2.mp3" 且 "3.2" 在列表中，则处理
        if not any(sheet in mp3 for sheet in my_target_sheets):
            continue
        # 进行音频文件的切割    
        process_single_unit(mp3)

if __name__ == "__main__":
    main()