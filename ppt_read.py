import math
import os
import shutil

import win32com.client
import xlrd
import xlwt
from aip import AipSpeech
from moviepy.editor import AudioFileClip
from pptx import Presentation

""" 你的 APPID AK SK """
APP_ID = '22663952'
API_KEY = '5bMW9nKQ6Oxqlvif31mr07h2'
SECRET_KEY = 'GzabHXI5mgFRwhKT2rhsTVNb3xBDtaPb'
#
client = AipSpeech(APP_ID, API_KEY, SECRET_KEY)


def reading(text, name):
    '''
    朗读程序
    :param text: 朗读内容
    :param name: 生成文件名
    :return:
    '''
    # print(text, name)
    result = client.synthesis(text, 'zh', 1, {
        'spd': 5,
        'vol': 5,
        'pit': 5,
        'per': 0,
    })
    print(result)

    if not isinstance(result, dict):
        with open('./audios/%s.mp3' % (str(name)), 'wb') as f:
            f.write(result)


def ppt2jpg():
    '''
    PPT转为JPG图片
    :return:
    '''
    ppt_app = win32com.client.Dispatch('PowerPoint.Application')
    ppt_app.Visible = 1
    print(os.getcwd() + '\朗读.pptx')
    ppt = ppt_app.Presentations.Open(os.getcwd() + '\朗读.pptx')  # 打开 ppt

    ppt.SaveAs(os.getcwd() + "\images", 17)  # 17数字是转为 ppt 转为图片
    # # ppt.SaveAs("D:\PythonTest\ppt_read\images", 17)  # 17数字是转为 ppt 转为图片
    print(len(ppt.Slides))
    ppt_app.Quit()  # 关闭资源，退出


def ppt2xls():
    '''
    PPT文字转为EXCEL文件
    :return:
    '''
    prs = Presentation('朗读.pptx')
    # 获取slide幻灯片
    picture_num = 0
    # 创建一个workbook 设置编码
    workbook = xlwt.Workbook(encoding='utf-8')
    # 创建一个worksheet
    worksheet = workbook.add_sheet('My Worksheet')

    # 写入excel
    # 参数对应 行, 列, 值
    worksheet.write(0, 0, label='this is test')
    worksheet.write(1, 0, label='页号')
    worksheet.write(1, 1, label='行号')
    worksheet.write(1, 2, label='文本内容')
    for slide in prs.slides:
        picture_num = picture_num + 1
        text_all = ''
        # 获取形状shape
        for shape in slide.shapes:

            if shape.has_text_frame:  # 判断是否有文字
                text_frame = shape.text_frame.text  # 获取文字框
                # print(text_frame)
                # print('----->', len(text_frame), '   ', text_frame)
                if len(text_frame) != 0:
                    text_all = text_all + text_frame

        print(picture_num, '     ', text_all)
        text_all = text_all.replace(" ", "").replace(" ", "")
        text_all = text_all.replace("\n", "  ").replace("\r", "  ")
        worksheet.write(picture_num + 1, 0, label=str(picture_num))
        worksheet.write(picture_num + 1, 1, label=text_all)
    workbook.save('朗读.xls')


def check_directory():
    '''
    检查文件目录结构
    :return:
    '''
    if not os.path.isfile('朗读.pptx'):
        print('请另存待处理文件为‘朗读.pptx’然后重新运行此程序。')
        input('按任意键退出此程序')
    if os.path.isdir('./images'):
        shutil.rmtree('images')
    if os.path.isdir('./audios'):
        shutil.rmtree('audios')
    if os.path.isdir('./temp'):
        shutil.rmtree('temp')
    os.mkdir('images')
    os.mkdir('audios')
    os.mkdir('temp')


def xls2mp3():
    '''
    朗读xls文件，转为MP3
    :return:
    '''
    book = xlrd.open_workbook('朗读.xls')

    sheet1 = book.sheets()[0]

    nrows = sheet1.nrows
    t_p_sn = 0
    print('表格总行数', nrows)
    for i in range(2, nrows):
        print(sheet1.cell(i, 0).value)
        print((sheet1.cell(i, 1).value))
        reading(sheet1.cell(i, 1).value, sheet1.cell(i, 0).value)


def audio_duration(file):
    print(file)
    audioclip = AudioFileClip(file)
    s = audioclip.duration
    audioclip.close()
    print(s)
    s = math.ceil(s)
    if s < 5:
        return 5
    return s


def merge():
    '''
    合并所有音视频
    :return:
    '''
    all_file = os.listdir('images')
    # print(len(all_file))
    for i in range(1, len(all_file) + 1):
        auido = './audios/%i.mp3' % i
        image = './images/幻灯片%i.JPG' % i
        print(auido, image)
        duration = audio_duration(auido)
        print(duration)


def stop():
    input('现在是修正读音暂停，请校验朗读.xls后按任意键继续')


def main():
    check_directory()
    # ppt2jpg()
    # ppt2xls()
    # stop()
    # xls2mp3()
    # merge()


if __name__ == '__main__':
    main()
