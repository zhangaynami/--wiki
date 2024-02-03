import re
import pandas as pd
import sys
from PyQt6.QtWidgets import QApplication, QWidget

username = "杰帕德"
dataframe = pd.read_excel(
    "D:/星铁wiki/角色语音文本填充/avatarvoice(粗排).xlsx",
    usecols=["AvataeName", "VoiceTitle", "VoiceCHS", "VoiceEN", "VoiceJP", "VoiceKR"],
)

# 行号
num = 0
titletext = (
    "{{角色语音表头|"
    + username
    + "|未完善=是}}\n"
    + "{{切换板|开始}}\n"
    + "{{切换板|默认显示|互动语音}}\n"
    + "{{切换板|默认折叠|战斗语音}}\n"
    + "{{切换板|显示内容}}"
)
# 页面切换
switchover = "{{切换板|内容结束}}\n" + "{{切换板|折叠内容}}"

for name in dataframe["AvataeName"]:
    # 匹配角色姓名
    if name == username:
        VoiceTitle = dataframe["VoiceTitle"][num]
        VoiceCHS = dataframe["VoiceCHS"][num]
        VoiceEN = dataframe["VoiceEN"][num]
        VoiceJP = dataframe["VoiceJP"][num]
        VoiceKR = dataframe["VoiceKR"][num]
        text = (
            "{{"
            + f"角色语音\n"
            + f"|语音类型={VoiceTitle}\n"
            + f"|语音文件={name}-{VoiceTitle}\n"
            + f"|语音内容={VoiceCHS}\n"
            + f"|语音内容日语={VoiceJP}\n"
            + f"|语音内容英语={VoiceEN}\n"
            + f"|语音内容韩语={VoiceKR}\n"
            + "}}\n"
        )
        titletext += text
        nextnum = num + 1
        if dataframe["VoiceTitle"][nextnum] == "战斗开始•弱点击破":
            titletext += switchover
    num += 1
# 结尾部分
end = """
{{切换板|内容结束}}
{{切换板|结束}}
"""

titletext = titletext + end
print(titletext)
with open("D:/星铁wiki/角色语音文本填充/output.txt", "w", encoding="utf-8") as file:
    file.write(titletext)
