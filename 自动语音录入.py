import re
import pandas as pd

dataframe = pd.read_excel(
    "语音.xlsx", usecols=["VoiceCHS", "VoiceJP", "VoiceEN", "VoiceKR"]
)
print(dataframe)
with open("01.txt", "r", encoding="utf-8") as file:
    content = file.readlines()
i = 0
# 将数据写入工作表
with open("output.txt", "w", encoding="utf-8") as output_file:
    for row in content:
        lines = row.strip().split("\n")  # 将数据按行拆分，并去除首尾空白

        # sheet.append(row)
        # print(row)

        # if row[-1] == "=":
        #     row & "hi"
        # print(row)

        for line in lines:
            # 判断是否以 = 结尾
            if re.match(r".*=$", line):
                if line[-3:-1] == "英语":
                    line = line + dataframe.iloc[i, 1]
                elif line[-3:-1] == "韩语":
                    line = line + dataframe.iloc[i, 3]
                elif line[-3:-1] == "日语":
                    line = line + dataframe.iloc[i, 2]
                print(f"{line}")
                output_file.write(line + "\n")
            else:
                print(f"{line}")
                output_file.write(line + "\n")

            if line == "}}":
                i = 1 + i

