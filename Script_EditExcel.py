import os
import glob
import whisper
import torch
from pydub import AudioSegment
import re

# ================= 配置区域 =================
# 判定相似度的阈值（完全匹配不需要改）
# 模型选择：base, small, medium, large-v3
# 4090显存足够，建议用 medium 或 large-v3 以保证极高准确率
MODEL_SIZE = "medium"
# ===========================================

def normalize_text(text):
    """简单的文本清洗，去除标点符号，转小写"""
    return re.sub(r'[^\w\s]', '', text).lower().strip()

def find_start_timestamp(model, audio_path, target_word):
    """
    使用 Whisper 识别音频，寻找目标单词第一次出现的时间点
    """
    print(f"正在分析音频内容: {os.path.basename(audio_path)} ...")

    # 启用 word_timestamps=True 来获取单词级的时间戳
    result = model.transcribe(audio_path, word_timestamps=True, language="en")

    target_clean = normalize_text(target_word)

    for segment in result['segments']:
        for word_info in segment['words']:
            # word_info 结构: {'word': ' Hello', 'start': 0.5, 'end': 0.9, ...}
            recognized_word = normalize_text(word_info['word'])

            # 简单的完全匹配。如果单词很短或容易识别错，可能需要模糊匹配，但通常这里够用了
            if recognized_word == target_clean:
                return word_info['start']

    return None

def process_folder():
    # 1. 检查 GPU 是否可用
    device = "cuda" if torch.cuda.is_available() else "cpu"
    print(f"正在加载 Whisper 模型 ({MODEL_SIZE}) 到 {device}...")
    try:
        model = whisper.load_model(MODEL_SIZE, device=device)
    except Exception as e:
        print(f"模型加载失败: {e}")
        return

    # 2. 获取当前目录下所有 Origin_*.txt 文件
    txt_files = glob.glob("Origin_*.txt")

    if not txt_files:
        print("未找到符合格式 Origin_*.txt 的单词表文件。")
        return

    print(f"找到 {len(txt_files)} 组任务，开始处理...")
    print("-" * 50)

    for txt_file in txt_files:
        # 解析对应的 mp3 文件名
        # 假设 txt 是 Origin_Filename.txt -> mp3 是 Filename.mp3
        base_name = txt_file.replace("Origin_", "").replace(".txt", "")
        mp3_file = f"{base_name}.mp3"

        if not os.path.exists(mp3_file):
            print(f"[跳过] 未找到对应的音频文件: {mp3_file}")
            continue

        # 读取单词表第一个单词
        first_word = ""
        try:
            with open(txt_file, 'r', encoding='utf-8') as f:
                # 过滤空行，找到第一行有效文本
                for line in f:
                    if line.strip():
                        first_word = line.strip().split()[0] # 取该行第一个词（防止有音标等干扰）
                        break
        except UnicodeDecodeError:
            # 尝试 GBK 编码读取
            with open(txt_file, 'r', encoding='gbk') as f:
                for line in f:
                    if line.strip():
                        first_word = line.strip().split()[0]
                        break

        if not first_word:
            print(f"[跳过] 单词表为空: {txt_file}")
            continue

        print(f"正在处理: {mp3_file}")
        print(f"  -> 目标首词: [{first_word}]")

        # 3. 寻找切割点
        start_time_sec = find_start_timestamp(model, mp3_file, first_word)

        if start_time_sec is None:
            print(f"  -> [警告] 在音频中未识别到单词 '{first_word}'，跳过此文件。")
            continue

        # 4. 执行切割
        print(f"  -> 定位成功，切割点: {start_time_sec:.2f}秒")

        try:
            audio = AudioSegment.from_mp3(mp3_file)

            # Pydub 单位是毫秒
            cut_point_ms = start_time_sec * 1000

            # 稍微往前留一点点缓冲 (比如 0.1秒)，避免切掉单词的辅音头
            # 如果不需要缓冲，设为 0
            buffer_ms = 0
            start_ms = max(0, cut_point_ms - buffer_ms)

            new_audio = audio[start_ms:]

            # 导出并覆盖原文件 (先写入临时文件，成功后再重命名，防止数据损坏)
            temp_output = f"temp_{mp3_file}"
            new_audio.export(temp_output, format="mp3")

            os.replace(temp_output, mp3_file)
            print(f"  -> [成功] 已覆盖原文件。")

        except Exception as e:
            print(f"  -> [错误] 音频处理失败: {e}")
            if os.path.exists(f"temp_{mp3_file}"):
                os.remove(f"temp_{mp3_file}")

    print("-" * 50)
    print("所有任务处理完成。")

if __name__ == "__main__":
    process_folder()