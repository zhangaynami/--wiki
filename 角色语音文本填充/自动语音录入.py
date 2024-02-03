import re
import pandas as pd

dataframe = pd.read_excel(
    "D:/星铁wiki/角色语音文本填充/语音.xlsx",
    usecols=["AvataeName", "VoiceTitle", "VoiceCHS", "VoiceEN", "VoiceJP", "VoiceKR"],
)

with open("D:/星铁wiki/角色语音文本填充/语音模板.txt", "r", encoding="utf-8") as file:
    content = file.readlines()
i = 0
# 将数据写入工作表
with open("D:/星铁wiki/角色语音文本填充/output.txt", "w", encoding="utf-8") as output_file:
    for row in content:
        lines = row.strip().split("\n")  # 将数据按行拆分，并去除首尾空白

        for line in lines:
            # 判断是否以 = 结尾
            print(line)
            if re.match(r".*=$", line):
                if line[-3:-1] == "类型":
                    line = line + dataframe.iloc[i, 1]
                elif line[-3:-1] == "文件":
                    line = line + dataframe.iloc[i, 0] + "-" + dataframe.iloc[i, 1]
                elif line[-3:-1] == "内容":
                    line = line + dataframe.iloc[i, 2]
                elif line[-3:-1] == "日语":
                    line = line + dataframe.iloc[i, 4]
                elif line[-3:-1] == "英语":
                    line = line + dataframe.iloc[i, 3]
                elif line[-3:-1] == "韩语":
                    line = line + dataframe.iloc[i, 5]
                print(line)
                output_file.write(line + "\n")
            else:
                # print(line)
                output_file.write(line + "\n")

            if line == "}}":
                i = 1 + i
