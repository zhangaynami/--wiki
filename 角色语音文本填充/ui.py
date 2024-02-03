import sys
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QTextEdit,
    QLabel,
    QVBoxLayout,
    QPushButton,
    QFrame,
    QLineEdit,
)
from PyQt5.QtGui import QIcon, QFont

import pandas as pd


class main(QWidget):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):

        # 顶部栏
        self.setWindowTitle("星穹铁道WIKI批量模板编辑器")
        self.setWindowIcon(QIcon("D:/星铁wiki/角色语音文本填充/titleicon.jpg"))

        self.setFixedSize(600, 400)

        # 创建子区域

        # 右下区域
        rightbottom_zone = QFrame(self)
        rightbottom_zone.setStyleSheet("border-top:2px solid rgba(149, 165, 166,0.3)")
        rightbottom_zone.setGeometry(150, 350, 450, 80)

        # 组件
        self.getname = QLineEdit(rightbottom_zone)
        self.getname.setStyleSheet(
            "border:0;color:rgb(44, 62, 80);border-radius:5px;font-size:14px"
        )
        self.getname.setGeometry(100, 10, 120, 30)

        # 提交按钮
        submit = QPushButton("提交", rightbottom_zone)
        submit.setStyleSheet(
            "border:0;color:rgb(236, 240, 241);background:rgb(46, 204, 113);border-radius:5px"
        )
        submit.setGeometry(350, 10, 60, 30)

        # 点击事件
        submit.clicked.connect(self.createmodel)

        # 左侧区域
        left_zone = QFrame(self)
        left_zone.setStyleSheet("background:rgba(149, 165, 166,0.3)")
        left_zone.setGeometry(0, 0, 150, 400)

        self.show()

    # def get_name(self):
    #     name = self.getname.text()
    #     print(name)

    # 生成模板的函数
    def createmodel(self):
        username = self.getname.text()
        dataframe = pd.read_excel(
            "D:/星铁wiki/角色语音文本填充/avatarvoice(粗排).xlsx",
            usecols=[
                "AvataeName",
                "VoiceTitle",
                "VoiceCHS",
                "VoiceEN",
                "VoiceJP",
                "VoiceKR",
            ],
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
        end = "{{切换板|内容结束}}\n" + "{{切换板|结束}}"

        titletext = titletext + end
        print(titletext)
        with open(
            "D:/星铁wiki/角色语音文本填充/output.txt", "w", encoding="utf-8"
        ) as file:
            file.write(titletext)


if __name__ == "__main__":

    app = QApplication(sys.argv)
    ex = main()
    sys.exit(app.exec_())
