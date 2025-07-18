from pyautocad import Autocad
from pyautocad import utils
import math, re
from itertools import chain
import datetime, time
import win32api, win32con, os
from itertools import chain

# from pypinyin import pinyin, Style
import statistics
import tkinter as tk
from tkinter.constants import CURRENT, W, E, N, S
import tkinter.font as tkFont
from collections import Counter
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from ttkbootstrap import Style
import json
import qrcode
from PIL import ImageTk
from ttkbootstrap.icons import Icon
import os
import sys


# print('主题切换以及悬浮提示功能由grok和腾讯元宝的deepseek模型实现和优化')
def get_config_path():
    if getattr(sys, 'frozen', False):
        # 打包后使用用户文档目录
        base_path = os.path.join(os.path.expanduser('~'), 'Documents')
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, 'theme_pref.json')

#不考虑列表中元素的位置排序，仅依据内容来检查两个列表是否相等
def is_list_equal_ignore_order(list1, list2):
    return Counter(list1) == Counter(list2)

# 时代成因、剖面层号列判断文本对象的组合
def concatenate_text_in_list(text_list):
    # text_list = [text_range_id, text_content, center_point_coordinate, text_object_id,text_bounding_box]
    concatenated_text_list = []
    used_text_list = []
    for text in text_list:
        if text not in used_text_list:
            used_text_list.append(text)
            range_id = text[0]
            content = text[1]
            center_point_coordinate = text[2]
            center_point_y = center_point_coordinate[1]
            obj_id = text[3]
            bounding_box = text[4]
            text_min_y = bounding_box[0][1]
            text_max_y = bounding_box[1][1]
            brother_text_list = []
            for text1 in text_list:
                if text1 not in used_text_list:
                    range_id_1 = text1[0]
                    content_1 = text1[1]
                    center_point_coordinate_1 = text1[2]
                    center_point_y_1 = center_point_coordinate_1[1]
                    obj_id_1 = text1[3]
                    bounding_box_1 = text1[4]
                    text_min_y_1 = bounding_box_1[0][1]
                    text_max_y_1 = bounding_box_1[1][1]
                    if (
                        (text_min_y >= text_min_y_1 and text_min_y <= text_max_y_1)
                        or (text_max_y >= text_min_y_1 and text_max_y <= text_max_y_1)
                        or (text_min_y_1 >= text_min_y and text_min_y_1 <= text_max_y)
                        or (text_max_y_1 >= text_min_y and text_max_y_1 <= text_max_y)
                    ):
                        brother_text_list.append(text1)
                        used_text_list.append(text1)
                        for text2 in text_list:
                            if text2 not in used_text_list:
                                range_id_2 = text2[0]
                                content_2 = text2[1]
                                center_point_coordinate_2 = text2[2]
                                center_point_y_2 = center_point_coordinate_2[1]
                                obj_id_2 = text2[3]
                                bounding_box_2 = text2[4]
                                text_min_y_2 = bounding_box_2[0][1]
                                text_max_y_2 = bounding_box_2[1][1]
                                if (
                                    (
                                        text_min_y_1 >= text_min_y_2
                                        and text_min_y_1 <= text_max_y_2
                                    )
                                    or (
                                        text_max_y_1 >= text_min_y_2
                                        and text_max_y_1 <= text_max_y_2
                                    )
                                    or (
                                        text_min_y_2 >= text_min_y_1
                                        and text_min_y_2 <= text_max_y_1
                                    )
                                    or (
                                        text_max_y_2 >= text_min_y_1
                                        and text_max_y_2 <= text_max_y_1
                                    )
                                ):
                                    brother_text_list.append(text2)
                                    used_text_list.append(text2)

            # print(text,' its Brothers:')
            brother_text_list.insert(0, text)
            brother_text_list = sorted(
                brother_text_list, key=lambda x: [(x[2][0], x[2][1])]
            )  # 先x升序再y升序
            # for i in brother_text_list:
            #     print('bbbbb: ',i)

            concatenated_text_min_x = min([x[4][0][0] for x in brother_text_list])
            concatenated_text_min_y = min([y[4][0][1] for y in brother_text_list])
            concatenated_text_max_x = max([x[4][1][0] for x in brother_text_list])
            concatenated_text_max_y = max([y[4][1][1] for y in brother_text_list])
            concatenated_text_bounding_box = (
                (concatenated_text_min_x, concatenated_text_min_y, 0.0),
                (concatenated_text_max_x, concatenated_text_max_y, 0.0),
            )
            concatenated_text_center_point_coordinate = (
                (concatenated_text_min_x + concatenated_text_max_x) / 2,
                (concatenated_text_min_y + concatenated_text_max_y) / 2,
                0,
            )
            concatenated_text_content = " @ ".join(
                [text[1] for text in brother_text_list]
            )
            concatenated_text = (
                range_id,
                concatenated_text_content,
                concatenated_text_center_point_coordinate,
                obj_id,
                concatenated_text_bounding_box,
            )

            concatenated_text_list.append(concatenated_text)

    # for item in concatenated_text_list:
    #     print(item)
    # print('+++++++++++++++++++++++++++++')
    return concatenated_text_list


def send_command_to_cad(cad_doc, command):  # 发送命令到cad
    cad_doc.ActiveDocument.SendCommand(command)


def get_string_in_range(text_list, min_x, max_x, min_y, max_y, sort_type):
    string_list = []
    # print('在它里面吗：',min_x, max_x, min_y,
    #                     max_y,'\n')
    for text in text_list:
        # print('最原始的文本对象：',text)

        if coor_inside_range(text[2][0], text[2][1], min_x, max_x, min_y, max_y):
            string_list.append(text)
            # print('这些是符合要求的：',text[1])
    if len(string_list) == 0:
        return "【空】"
    # for item in string_list:
    #     print('最原始的文本对象text：',item)
    # print(string_list)
    if sort_type == 1:  # 岩土描述
        string_list = [
            (item[0], item[1], (round(item[2][0]), round(item[2][1])), item[3])
            for item in string_list
        ]  # 这个是x，y坐标四舍五入，因为岩土描述里面同一行的标点符号和普通文字插入点的y坐标不是完全一样的
        string_text_list = [
            "Obj_ID $ " + str(text[3]) + " $ " + str(text[1])
            for text in list(
                sorted(string_list, key=lambda x: (-x[2][1], x[2][0]))
            )  # 按先y降序再x升序排列(岩土描述)   (-x[2][1], x[2][0])
        ]
        # for text in string_text_list:
        #     print('你不对劲：143行 ', text)
        # string_text_list.insert(0,' @ ')
        # 只有一个文本对象
        if len(string_text_list) == 1:
            combin_string = [
                item + " @ Obj_ID $ 999999999999999 $ " for item in string_text_list
            ][0]
            # print('为什么啊？？？',item)
        else:
            # 多个文本对象
            combin_string = " @ ".join(string_text_list)
        # return combin_string
        # print('------------------')
    else:
        string_text_list = [
            text[1]
            for text in list(
                sorted(string_list, key=lambda x: (x[2][0], x[2][1]))
            )  # 按先x升序再y升序排列(时代成因)
        ]
        # string_text_list.insert(0,' @ ')
        combin_string = " @ ".join(string_text_list)
    return combin_string


def get_hor_line_nearest_up_text(
    hor_line_y, min_x, max_x, line_list, text_list, sort_type, frame_bottom_y
):  # 返回横线上方格子最近的文本对象(以字段对象插入点辅助)
    # for text in text_list:
    #     print(text)
    mid_field_x = (min_x + max_x) / 2
    nearest_y_up_line_list = []
    nearest_y_down_line_list = []
    for line in line_list:
        line_start_point_x = line[0]
        line_start_point_y = line[1]
        line_end_point_x = line[2]
        line_end_point_y = line[3]
        line_range_id = line[4]
        line_max_x = max(line_start_point_x, line_end_point_x)
        line_min_x = min(line_start_point_x, line_end_point_x)
        line_max_y = max(line_start_point_y, line_end_point_y)
        line_min_y = min(line_start_point_y, line_end_point_y)
        # print(mid_field_x,line_min_x,line_max_x,hor_line_y,line_min_y)
        if (
            mid_field_x >= line_min_x
            and mid_field_x <= line_max_x
            and hor_line_y < line_min_y
        ):
            nearest_y_up_line_list.append(line)
    # for l in nearest_y_up_line_list:
    #     print(l)

    # exit()
    nearest_y_up_line = list(sorted(nearest_y_up_line_list, key=lambda x: (x[1])))[
        0
    ]  # 按坐标y升序后选第一条(y最小者)
    # print('top',nearest_y_up_line)
    max_y = nearest_y_up_line[1]  # y坐标
    # print('最大Y', max_y)
    # print('---------')
    for line in line_list:
        line_start_point_x = line[0]
        line_start_point_y = line[1]
        line_end_point_x = line[2]
        line_end_point_y = line[3]
        line_range_id = line[4]
        line_max_x = max(line_start_point_x, line_end_point_x)
        line_min_x = min(line_start_point_x, line_end_point_x)
        line_max_y = max(line_start_point_y, line_end_point_y)
        line_min_y = min(line_start_point_y, line_end_point_y)
        # print(mid_field_x,line_min_x,line_max_x,hor_line_y,nearest_y_up_line[1],line_min_y)
        if (
            mid_field_x >= line_min_x
            and mid_field_x <= line_max_x
            and nearest_y_up_line[1] > line_min_y
            and round(max_y, 5) > round(line_min_y, 5)
        ):
            nearest_y_down_line_list.append(line)
    # for l in nearest_y_down_line_list:
    #     print('底部的直线：', l)
    nearest_y_down_line = list(sorted(nearest_y_down_line_list, key=lambda x: (-x[1])))[
        0
    ]  # 按坐标y降序后选第一条(y最大者)
    # if round(nearest_y_down_line[1],2) == round(max_y,2):#太坑了，有的水平线原来不是完美水平，到小数点后十几位就不同了，上面的if已经用round改了
    #     nearest_y_down_line = list(
    #         sorted(nearest_y_down_line_list, key=lambda x:
    #             (-x[1])))[1]  #按坐标y降序后选第e二条(y次大者，排除最大者)

    # print(nearest_y_down_line)
    # print('bottom',nearest_y_down_line)
    # frame_bottom_y
    min_y = nearest_y_down_line[1]  # y坐标
    # print('最小Y', min_y)

    # print('y最小值：',min_y)
    # print('岩土描述范围：',min_x, max_x, min_y, max_y,'边框底部y：',frame_bottom_y)
    if "bottom:" in str(frame_bottom_y):
        min_y = float(str(frame_bottom_y).split(":")[1])
    inner_string = get_string_in_range(
        text_list, min_x, max_x, min_y, max_y, sort_type
    )  # 格子内文本
    # print(min_x, max_x, min_y, max_y, '内容:', inner_string)
    # print('######################################')
    # exit()
    return min_x, max_x, min_y, max_y, inner_string


def get_text_nearest_one_line(text_x, text_y, line_list, navigation):
    nearest_line_list = (
        []
    )  # (start_point[0], start_point[1], end_point[0],end_point[1], line_range_id)
    for line in line_list:
        line_start_point_x = line[0]
        line_start_point_y = line[1]
        line_end_point_x = line[2]
        line_end_point_y = line[3]
        line_range_id = line[4]
        line_max_x = max(line_start_point_x, line_end_point_x)
        line_min_x = min(line_start_point_x, line_end_point_x)
        line_max_y = max(line_start_point_y, line_end_point_y)
        line_min_y = min(line_start_point_y, line_end_point_y)
        if navigation == "下":
            # 默认向下就是找水平线
            if text_x > line_min_x and text_x < line_max_x and text_y > line_min_y:
                # 下方水平线
                nearest_line_list.append(line)
    # for line in nearest_line_list:
    #     print('蛤？')
    #     print(line)
    if navigation == "下":
        nearest_line_list = list(
            sorted(nearest_line_list, key=lambda x: (-x[1]))
        )  # 按坐标y降序后选第一条(y最大者)
    nearest_line = nearest_line_list[0]
    return nearest_line  # (line_start_point_x,line_start_point_y,line_end_point_x,line_end_point_y,line_range_id)

    return nearest_line_list


def if_text_content_part_in_string(text_content, string):
    text_content = text_content.strip()
    print("【", string, "】", " not in ", text_content, " ?")
    if text_content in string:
        print(string, "in ", text_content, " !")
        return True


def if_str_all_in_list(text_str, key_list):
    if all(key_char in text_str for key_char in key_list):
        return True


# def to_pinyin(s):  #中文转拼音
#     return ''.join(chain.from_iterable(pinyin(s, style=Style.TONE3)))


def get_max_and_second_max_list(l, approximate):
    l.sort(reverse=True)
    max_list = [l[0]]
    for num1, num2 in zip(l[:], l[1:]):
        # print(num1,num2)
        if num2 / num1 > approximate:
            max_list.append(num1)
            max_list.append(num2)
        else:
            break
    max_list = list(set(max_list))  # 最大值列表（0.99相似）
    max_list.sort(reverse=True)
    for num in max_list:
        l.remove(num)  # 删除最大值

    second_max_list = [l[0]]
    for num1, num2 in zip(l[:], l[1:]):
        if num2 / num1 > 0.99:
            second_max_list.append(num1)
            second_max_list.append(num2)
        else:
            break
    second_max_list = list(set(second_max_list))  # 次大值列表（0.99相似）
    second_max_list.sort()
    return max_list, second_max_list


def get_current_time():
    dtime = datetime.datetime.now()
    untime = time.mktime(dtime.timetuple())
    times = datetime.datetime.fromtimestamp(untime)
    return str(times)


def seperate_Chinese(strings):  # 拆分中文字符串为单个字
    chn_pattern = re.compile(r"([\u4e00-\u9fff])")  # 中文字符正则表达式
    chars = chn_pattern.split(strings)
    chars = [c for c in chars if len(c.strip()) > 0]
    return chars


def coor_inside_range(targetX, targetY, minX, maxX, minY, maxY):
    if targetX >= minX and targetX <= maxX and targetY >= minY and targetY <= maxY:
        return True


def point_adscription(point_coor, range_list):  # 判断点在哪个矩形范围之内，返回矩形id
    point_coor_x = point_coor[0]
    point_coor_y = point_coor[1]
    point_in_range_list = []
    # print(point_coor_x)
    # print(point_coor_y)
    # print("=======================")
    for range_coor in range_list:
        range_id = range_coor[0]
        range_min_x = range_coor[1]
        range_min_y = range_coor[2]
        range_max_x = range_coor[3]
        range_max_y = range_coor[4]
        if (
            (point_coor_x > range_min_x)
            and (point_coor_x < range_max_x)
            and (point_coor_y > range_min_y)
            and (point_coor_y < range_max_y)
        ):
            point_in_range_list.append(range_id)
    return point_in_range_list


def line_adscription(
    line_start_point,
    line_end_point,
    range_list,  # 判断直线在哪个矩形范围之内，返回矩形id
):
    line_start_coor_x = round(line_start_point[0], 2)
    line_start_coor_y = round(line_start_point[1], 2)
    line_end_coor_x = round(line_end_point[0], 2)
    line_end_coor_y = round(line_end_point[1], 2)
    line_in_range_list = []
    # print(point_coor_x)
    # print(point_coor_y)
    # print("=======================")
    for range_coor in range_list:
        range_id = range_coor[0]
        range_min_x = round(range_coor[1], 2)
        range_min_y = round(range_coor[2], 2)
        range_max_x = round(range_coor[3], 2)
        range_max_y = round(range_coor[4], 2)
        if (
            (line_start_coor_x >= range_min_x)
            and (line_start_coor_x <= range_max_x)
            and (line_start_coor_y >= range_min_y)
            and (line_start_coor_y <= range_max_y)
            and (line_end_coor_x >= range_min_x)
            and (line_end_coor_x <= range_max_x)
            and (line_end_coor_y >= range_min_y)
            and (line_end_coor_y <= range_max_y)
        ):
            line_in_range_list.append(range_id)
    return line_in_range_list


def get_string_list(
    target_string, text_content, text_insert_coordinate, text_range_id, target_list
):
    if target_string in text_content:  # 包含柱状图标题所有字符的文本对象
        target_list.append((text_insert_coordinate, text_range_id))


def text_contains_str(text_obj):
    global title_name
    str = title_name
    return str in text_obj.TextString.replace(" ", "")


def get_neraby_text(
    field_name,
    field_coordinates,
    range_id,
    range_list,
    vertical_line_list,
    horizon_line_list,
    horizon_polyline_with_range_list,
    vertical_polyline_with_range_list,
    navigation,
    text_list,
    txt,
):
    global title_text_height
    part_str = ""
    if "@" in navigation:
        temp = navigation.split("@")[0]
        part_str = navigation.split("@")[1]  # 目标文本中要求包含某字符串
        navigation = temp
    range = [x for x in range_list if x[0] == range_id][0]  # 柱状图范围
    range_minX = range[1]
    range_minY = range[2]
    range_maxX = range[3]
    range_maxY = range[4] - title_text_height  # 减去标题高度，变回多段线maxY
    if navigation != "标题":  # 找标题附近文本对象不依赖四周横竖线
        ver_line_list = [x for x in vertical_line_list if x[4] == range_id]
        ver_polyline_list = [
            x for x in vertical_polyline_with_range_list if x[4] == range_id
        ]
        ver_line_polyline_list = (
            ver_line_list + ver_polyline_list
        )  # 直线多段线竖线列表合并
        hor_line_list = [x for x in horizon_line_list if x[4] == range_id]
        hor_polyline_list = [
            x for x in horizon_polyline_with_range_list if x[4] == range_id
        ]
        hor_line_polyline_list = (
            hor_line_list + hor_polyline_list
        )  # 直线多段线横线列表合并

        field_coordinates_x = field_coordinates[0]
        field_coordinates_y = field_coordinates[1]

        ver_line_around_field_x = [
            l[0]
            for l in ver_line_polyline_list
            if field_coordinates_y > min(l[1], l[3])
            and field_coordinates_y < max(l[1], l[3])
        ]
        ver_line_around_field_x.append(
            field_coordinates_x
        )  # 将字段x坐标存入候选x坐标列表
        ver_line_around_field_x.append(range_minX)  # 柱状图x范围
        ver_line_around_field_x.append(range_maxX)  # 柱状图x范围
        ver_line_around_field_x.sort()

        # ver_line_around_field_x = list(set(ver_line_around_field_x))#不知道为什么会有重复，damn
        # print('++++++++++++++')
        # for line in ver_line_around_field_x:
        #     print(line)
        right_line_x = ver_line_around_field_x[
            ver_line_around_field_x.index(field_coordinates_x) + 1
        ]  # 右侧第一条竖线x
        try:
            right_second_line_x = ver_line_around_field_x[
                ver_line_around_field_x.index(field_coordinates_x) + 2
            ]  # 右侧第二条竖线x
            right_second_line_x_substitude = ver_line_around_field_x[
                ver_line_around_field_x.index(field_coordinates_x) + 3
            ]  # 右侧第三？条竖线x(唔知点解有重复，目前默认只会重复一次先)
            if right_second_line_x == right_line_x:
                right_second_line_x = right_second_line_x_substitude

        except:
            pass
        # print(range_id,'右侧x',right_line_x,' -> ',right_second_line_x)
        left_line_x = ver_line_around_field_x[
            ver_line_around_field_x.index(field_coordinates_x) - 1
        ]  # 左侧第一条竖线x

        # print(range_id,'左侧x',left_line_x)

        hor_line_around_field_y = [
            l[1]
            for l in hor_line_polyline_list
            if field_coordinates_x > min(l[0], l[2])
            and field_coordinates_x < max(l[0], l[2])
        ]

        hor_line_around_field_y.append(
            field_coordinates_y
        )  # 将目标x坐标存入候选x坐标列表
        hor_line_around_field_y.append(range_minY)  # 柱状图y范围
        hor_line_around_field_y.append(range_maxY)  # 柱状图y范围(减去标题高度)
        hor_line_around_field_y.sort()
        up_line_y = hor_line_around_field_y[
            hor_line_around_field_y.index(field_coordinates_y) + 1
        ]  # 上方第一条横线y
        try:
            up_second_line_y = hor_line_around_field_y[
                hor_line_around_field_y.index(field_coordinates_y) + 2
            ]  # 上方第二条横线y
        except:
            pass
        # print(range_id,'上方y',up_line_y)
        down_line_y = hor_line_around_field_y[
            hor_line_around_field_y.index(field_coordinates_y) - 1
        ]  # 下方第一条横线y
        try:
            down_second_line_y = hor_line_around_field_y[
                hor_line_around_field_y.index(field_coordinates_y) - 2
            ]  # 下方第二条横线y
        except:
            pass
    # print(range_id,'下方y',down_line_y)

    target_text_list = []
    print_add = ""
    if part_str != "":
        print_add = "包含(" + str(part_str) + ")字符串"
    content = "格子内为【空】"
    if navigation == "上":
        # print('柱状图标识：', range_id, '字段名：', field_name, '右侧格子坐标范围:x：',
        #       right_line_x, '->', right_second_line_x, ' y：', down_line_y,
        #       '->', up_line_y)
        try:
            target_text_list = [
                text
                for text in text_list
                if text[0] == range_id
                and coor_inside_range(
                    text[2][0],
                    text[2][1],
                    left_line_x,
                    right_line_x,
                    up_line_y,
                    up_second_line_y,
                )
            ]
            target_text_list = sorted(
                target_text_list, key=lambda x: [x[2][0]]
            )  # 同一格子内文本对象按x升序
        except:
            pass
            # print('字段【' + field_name + '】上方内容为空，请检查对应目标项所填方向是否错误【改成"下"？"右"？】')
            # txt.write('柱状图标识：' + str(range_id) + ' 字段名：' + field_name +
            #           ' 上方文本对象内容为：空，请检查对应目标项所填方向是否错误【改成"下"？"右"？】\n')
            return
        if len(target_text_list) == 0:
            pass
            # print('柱状图标识：', range_id, '字段名：', field_name, ' 上方为空')
            # txt.write('柱状图标识：' + str(range_id) + ' 字段名：' + field_name +
            #           ' 上方为空' + '\n')
        else:
            target_text_content_list = [
                target[1] for target in target_text_list if part_str in target[1]
            ]
            up_content = " $ ".join(target_text_content_list)
            content = up_content
            # print('柱状图标识：', range_id, '字段名：' + field_name + ' 上方文本对象内容为:【',
            #       up_content + '】'+ print_add )
            # txt.write('柱状图标识：' + str(range_id) + ' 字段名：' + field_name +
            #           ' 上方文本对象内容为:【' + str(up_content) + '】'+ print_add +'\n')
            # for index, target, in enumerate(target_text_list):
            # print('柱状图标识：', range_id, '字段名：', field_name, '序号：', index,
            #       '右侧文本对象内容为：', target[1], '坐标为', target[2])

    if navigation == "右":
        # print('柱状图标识：', range_id, '字段名：', field_name, '右侧格子坐标范围:x：',
        #       right_line_x, '->', right_second_line_x, ' y：', down_line_y,
        #       '->', up_line_y)
        # if  "稳定水位(标高)" in field_name:
        #     print('[',field_name,']')
        #     exit()
        try:
            target_text_list = [
                text
                for text in text_list
                if text[0] == range_id
                and coor_inside_range(
                    text[2][0],
                    text[2][1],
                    right_line_x,
                    right_second_line_x,
                    down_line_y,
                    up_line_y,
                )
            ]
            target_text_list = sorted(
                target_text_list, key=lambda x: [x[2][0]]
            )  # 同一格子内文本对象按x升序
        except:
            pass
            # print('字段【' + field_name + '】右侧内容为空，请检查对应目标项所填方向是否错误【改成"下"？】')
            # txt.write('柱状图标识：' + str(range_id) + ' 字段名：' + field_name +
            #           ' 右侧文本对象内容为：空，请检查对应目标项所填方向是否错误【改成"下"？】\n')
            return
        if len(target_text_list) == 0:
            pass
            # print('柱状图标识：', range_id, '字段名：', field_name, '【】')
            # txt.write('柱状图标识：' + str(range_id) + ' 字段名：' + field_name + ' 【】' +
            #           '\n')
        else:
            target_text_content_list = [
                target[1] for target in target_text_list if part_str in target[1]
            ]
            right_content = " $ ".join(target_text_content_list)
            content = right_content
            # print('柱状图标识：', range_id, '字段名：' + field_name + ' 右侧文本对象内容为:【',
            #       right_content + '】' + print_add)
            # txt.write('柱状图标识：' + str(range_id) + ' 字段名：' + field_name +
            #           ' 右侧文本对象内容为:【' + str(right_content) + '】' + print_add +
            #           '\n')
            # for index, target, in enumerate(target_text_list):
            # print('柱状图标识：', range_id, '字段名：', field_name, '序号：', index,
            #       '右侧文本对象内容为：', target[1], '坐标为', target[2])

    if navigation == "下":
        # print('柱状图标识：', range_id, '字段名：', field_name, '下方格子坐标范围:x：',
        #       left_line_x, '->', right_line_x, ' y：', down_second_line_y, '->',
        #       down_line_y)
        try:
            target_text_list = [
                text
                for text in text_list
                if text[0] == range_id
                and coor_inside_range(
                    text[2][0],
                    text[2][1],
                    left_line_x,
                    right_line_x,
                    down_second_line_y,
                    down_line_y,
                )
            ]
            target_text_list = sorted(
                target_text_list, key=lambda x: [x[2][0]]
            )  # 同一格子内文本对象按x升序
        except:
            # print('字段【' + field_name + '】下方内容为空，请检查对应目标项所填方向是否错误')
            # txt.write('柱状图标识：' + str(range_id) + ' 字段名：' + field_name +
            #           ' 下方文本对象内容为:【空，请检查对应目标项所填方向是否错误】\n')
            pass
            return
        if len(target_text_list) == 0:
            pass
            # print('柱状图标识：', range_id, '字段名：', field_name, '【】')
            # txt.write('柱状图标识：' + str(range_id) + ' 字段名：' + field_name + ' 【】' +
            #           '\n')
        else:
            target_text_content_list = [
                target[1] for target in target_text_list if part_str in target[1]
            ]
            down_content = " $ ".join(target_text_content_list)
            content = down_content
            # print('柱状图标识：', range_id, '字段名：', field_name, ' 下方文本对象内容为:【',
            #       down_content + '】' + print_add)
            # txt.write('柱状图标识：' + str(range_id) + ' 字段名：' + field_name +
            #           ' 下方文本对象内容为:【' + str(down_content) + '】' + print_add +
            #           '\n')

    if navigation == "标题":
        target_text_list = [
            text
            for text in text_list
            if text[0] == range_id
            and coor_inside_range(
                text[2][0],
                text[2][1],
                range_minX,
                range_maxX,
                range_maxY,
                range_maxY + title_text_height,
            )
        ]
        target_text_list = sorted(
            target_text_list, key=lambda x: [x[2][0]]
        )  # 同一格子内文本对象按x升序
        if len(target_text_list) == 0:
            pass
            # print('柱状图标识：', range_id, '字段名：', field_name,
            #       '标题附近不存在：【' + part_str + '】')
            # txt.write('柱状图标识：' + str(range_id) + ' 字段名：' + field_name +
            #           ' 标题附近不存在：【' + part_str + '】' + '\n')
        else:
            target_text_content_list = [
                target[1] for target in target_text_list if part_str in target[1]
            ]
            title_part_content = " $ ".join(target_text_content_list)
            content = title_part_content
            # print(
            #     '柱状图标识：', range_id, '字段名：', field_name, ' 标题附近包含字符串【' +
            #     str(part_str) + '】的对象内容为:【' + str(title_part_content) + '】')
            # txt.write('柱状图标识：' + str(range_id) + ' 字段名：' + field_name +
            #           ' 标题附近包含字符串【' + str(part_str) + '】的对象内容为:【' +
            #           str(title_part_content) + '】\n')
    return range_id, field_name, navigation, content, print_add

#寻找同一格子内的文字对象
def get_partner_in_the_same_cell(
    text_obj,
    text_obj_list,
    range_list,
    vertical_line_list,
    vertical_polyline_list,
    horizon_line_list,
    horizon_polyline_list,
):
    range = range_list[0]  # 柱状图范围
    range_minX = range[1]
    range_minY = range[2]
    range_maxX = range[3]
    range_maxY = range[4] - title_text_height  # 减去标题高度，变回多段线maxY

    ver_line_list = vertical_line_list
    ver_polyline_list = vertical_polyline_list
    ver_line_polyline_list = ver_line_list + ver_polyline_list  # 直线多段线竖线列表合并
    hor_line_list = horizon_line_list
    hor_polyline_list = horizon_polyline_list
    hor_line_polyline_list = hor_line_list + hor_polyline_list  # 直线多段线横线列表合并
    field_coordinates_x = text_obj[2][0]
    field_coordinates_y = text_obj[2][1]

    # 左右竖线
    ver_line_around_field_x = [  # 筛选x
        l[0]
        for l in ver_line_polyline_list
        if field_coordinates_y > min(l[1], l[3])
        and field_coordinates_y < max(l[1], l[3])
    ]
    ver_line_around_field_x.append(field_coordinates_x)  # 将字段x坐标存入候选x坐标列表
    ver_line_around_field_x.append(range_minX)  # 柱状图x范围
    ver_line_around_field_x.append(range_maxX)  # 柱状图x范围
    ver_line_around_field_x.sort()
    right_line_x = ver_line_around_field_x[
        ver_line_around_field_x.index(field_coordinates_x) + 1
    ]  # 【四周定位线】右侧第一条竖线x
    left_line_x = ver_line_around_field_x[
        ver_line_around_field_x.index(field_coordinates_x) - 1
    ]  # 【四周定位线】左侧第一条竖线x

    print(text_obj)
    print("\n", "左竖线：", left_line_x)
    print("右竖线：", right_line_x, "\n")

    hor_line_around_field_y = [  # 筛选y
        l[1]
        for l in hor_line_polyline_list
        # ① 如果表格线有轻微错位的情况，如：
        # 存在出头、悬挂节点、或者一个单元格中用一根不接触左右两侧竖线的横线
        # 分隔内容时，就不能严格限制横线坐标，就用下面的两行）
        if field_coordinates_x > min(l[0], l[2])
        and field_coordinates_x < max(l[0], l[2])
        # ② 理想的表格的纵横线应该是完美相接，用下面两行（但是往往世事难料）
        # if left_line_x >= min(l[0], l[2])
        # and right_line_x <= max(l[0], l[2])
    ]

    hor_line_around_field_y.append(field_coordinates_y)  # 将目标x坐标存入候选x坐标列表
    hor_line_around_field_y.append(range_minY)  # 柱状图y范围
    hor_line_around_field_y.append(range_maxY)  # 柱状图y范围(减去标题高度)
    hor_line_around_field_y.sort()
    up_line_y = hor_line_around_field_y[
        hor_line_around_field_y.index(field_coordinates_y) + 1
    ]  # 【四周定位线】上方第一条横线y
    down_line_y = hor_line_around_field_y[
        hor_line_around_field_y.index(field_coordinates_y) - 1
    ]  # 【四周定位线】下方第一条横线y

    print("\n", "上横线：", up_line_y)
    print("下横线：", down_line_y, "\n")
    # exit()

    target_text_content_list = [
        (text[1], text[2][0], text[2][1])
        for text in text_obj_list
        if coor_inside_range(
            text[2][0], text[2][1], left_line_x, right_line_x, down_line_y, up_line_y
        )
    ]
    sort_by_coors_target_text_content_list = sorted(
        target_text_content_list, key=lambda x: (-x[2], x[1])
    )  # 先按坐标y降序，再按x升序
    all_str_in_same_cell_combine = "".join(
        [text[0] for text in sort_by_coors_target_text_content_list]
    )

    print("\n", all_str_in_same_cell_combine)
    print("字段名下方第一条横线y坐标：", down_line_y, "\n")

    # print("x范围：",left_line_x, right_line_x)
    return (
        all_str_in_same_cell_combine,
        (left_line_x, right_line_x, down_line_y, up_line_y),
    )  # 返回同一个格子里面的内容拼接在一起,还有格子坐标范围


def findSubStrIndex(substr, str, time):
    times = str.count(substr)
    if (times == 0) or (times < time):
        pass
    else:
        i = 0
        index = -1
        while i < time:
            index = str.find(substr, index + 1)
            i += 1
        return index


class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self.widget.bind("<Enter>", self.show)
        self.widget.bind("<Leave>", self.hide)

    def show(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 50
        y += self.widget.winfo_rooty() + 30
        
        # 使用ttkbootstrap样式创建提示窗口
        self.tip_window = ttkb.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry(f"+{x}+{y}")
        
        # 使用ttkbootstrap标签样式
        label = ttkb.Label(
            self.tip_window, 
            text=self.text, 
            bootstyle="inverse-secondary",
            padding=(4, 2),
            relief="solid",
            font=("微软雅黑", 10, "bold")
        )
        label.pack()

    def hide(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


class QRCodeTooltip:
    def __init__(self, widget, qr_data):
        self.widget = widget
        self.qr_data = qr_data
        self.tip_window = None
        self.widget.bind("<Enter>", self.show)
        self.widget.bind("<Leave>", self.hide)

    def generate_qr(self):
        """动态生成二维码图像[5,8](@ref)"""
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=4,
            border=1,
        )
        qr.add_data(self.qr_data)
        qr.make(fit=True)
        return qr.make_image(fill_color="#7E42D1", back_color="white")

    def show(self, event=None):
        """显示二维码窗口[1,4](@ref)"""
        if self.tip_window:
            return

        # 生成二维码
        qr_image = self.generate_qr()
        self.tk_image = ImageTk.PhotoImage(qr_image)

        # 创建悬浮窗口
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() - 150  # 向上偏移避免超出屏幕

        self.tip_window = ttkb.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)
        self.tip_window.wm_geometry(f"+{x}+{y}")

        # 添加动画效果[4](@ref)
        self.tip_window.attributes("-alpha", 0.0)
        self.tip_window.after(10, lambda: self._fade_in())

        label = ttkb.Label(
            self.tip_window, 
            image=self.tk_image,
            bootstyle="secondary",
            padding=2
        )
        label.pack()

    def _fade_in(self):
        """渐入动画效果[4](@ref)"""
        alpha = self.tip_window.attributes("-alpha")
        if alpha < 1.0:
            alpha += 0.1
            self.tip_window.attributes("-alpha", alpha)
            self.tip_window.after(10, self._fade_in)

    def hide(self, event=None):
        """隐藏二维码窗口[3](@ref)"""
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


class SampleApp(ttkb.Window):
    def toggle_console(self):
        if self.is_console_visible:
            self.console_frame.grid_forget()  # 隐藏控制台
            self.toggle_console_button.config(text="☰")
            global window_width, window_height
            self.geometry(f"{window_width}x{self.winfo_height()}")  # 恢复原始宽度
        else:
            self.console_frame.grid(row=0, column=4, rowspan=19, sticky=W+E+N+S, padx=10, pady=10)  # 显示控制台
            self.toggle_console_button.config(text="☱")
            self.update_idletasks()  # 更新布局
            new_width = self.winfo_width() + self.console_frame.winfo_width()  # 计算新宽度
            self.geometry(f"{new_width}x{self.winfo_height()}")  # 调整窗口宽度
        self.is_console_visible = not self.is_console_visible


    def write(self, message):

        if "Error" in message or "Exception" in message or "Traceback" in message:
                self.console_output.insert(tk.END, message)
                self.console_output.see(tk.END)
                self.update_idletasks()
                return
        
        self.buffer += message
            # 每隔一定数量的字符或行数批量插入
        if "正在识别每个柱状图范围" in self.buffer or "程序运行结束" in self.buffer or "正在遍历" in self.buffer or "完成柱状图外框识别时间" in self.buffer or len(self.buffer) > 5000:  # 例如每1000字符插入一次
            self.console_output.insert(tk.END, self.buffer)
            self.buffer = ""
            self.console_output.see(tk.END)
            self.update_idletasks()

    def flush(self):
        """实现flush方法"""
        self.original_stdout.flush()
        
    def clear_console(self):
        """清空输出控制台"""
        self.console_output.delete(1.0, tk.END)


    def new_entry(self, default_value, row, column, rowspan, columnspan):
        entry = ttkb.Entry(self, width=18)
        entry.insert(0, str(default_value))  # 显式设置默认值
        entry.grid(
            row=row, column=column, 
            rowspan=rowspan, columnspan=columnspan, 
            padx=5, pady=5, sticky=W+E
        )
        return entry
    def new_combobox(self, default_value_tuple, row, column, rowspan, columnspan):
        combobox = ttkb.Combobox(self, values=default_value_tuple, width=15)
        combobox.current(0)
        combobox.grid(
            row=row, column=column, rowspan=rowspan, columnspan=columnspan, padx=5, pady=5, sticky=W+E
        )

        return combobox

    def new_combobox_with_width(self, default_value_tuple, row, column, rowspan, columnspan):
        combobox = ttkb.Combobox(self, values=default_value_tuple, width=5,state="readonly")
        combobox.current(0)
        combobox.grid(
            row=row, column=column, rowspan=rowspan, columnspan=columnspan, padx=5, pady=5, sticky=W+E
        )
        return combobox

    def new_label(self, string, row, column, rowspan, columnspan, style='TLabel', tooltip_text=None):
        label = ttkb.Label(
            self,
            text=string,
            style=style,  # 使用传入的样式参数
            foreground='#34A853',
            font=("幼圆", 11, "bold"),
            width=15,
        )
        label.grid(
            row=row, column=column, 
            rowspan=rowspan, columnspan=columnspan, 
            padx=5, pady=5, sticky=W
        )
        # 添加悬浮提示功能
        if tooltip_text:
            Tooltip(label, tooltip_text)
        return label

    def new_label_with_width(self, string, row, column, rowspan, columnspan, width,tooltip_text=None):
        label = ttkb.Label(
            self,
            text=string,
            foreground='#EA4335',
            font=("微软雅黑", 13, "bold"),
            width=width,
        )
        label.grid(
            row=row, column=column, rowspan=rowspan, columnspan=columnspan, padx=5, pady=5, sticky=W
        )
        if tooltip_text:
            Tooltip(label, tooltip_text)
        return label

    def new_one2one_comboboxes(self, default_value_tuple_1, row, column, rowspan, columnspan):
        combobox1 = ttkb.Combobox(self, values=default_value_tuple_1, width=15)
        combobox1.current(0)
        combobox1.grid(
            row=row, column=column, rowspan=rowspan, columnspan=columnspan, padx=5, pady=5, sticky=W+E
        )
        combobox2 = ttkb.Combobox(self, values=("右", "下", "上", "标题"), width=5)
        combobox2.current(0)
        combobox2.grid(
            row=row, column=column + 1, rowspan=rowspan, columnspan=columnspan, padx=5, pady=5, sticky=W+E
        )
        return combobox1, combobox2

    def new_one2one_entry(self, default_value, row, column, rowspan, columnspan):
        entry = ttkb.Entry(self, textvariable=tk.StringVar(value=default_value), width=18)
        entry.grid(
            row=row, column=column, rowspan=rowspan, columnspan=columnspan, padx=5, pady=5, sticky=W+E
        )
        combobox = ttkb.Combobox(self, values=("右", "下", "上", "标题"), width=5)
        combobox.current(0)
        combobox.grid(
            row=row, column=column + 1, rowspan=rowspan, columnspan=columnspan, padx=5, pady=5, sticky=W+E
        )
        return entry, combobox

    def __init__(self):
        super().__init__(themename=self.load_theme())
        self.buffer = ""
        self.is_console_visible = False  # 添加一个标志位，用于控制控制台的显示状态
        # 配置必填项样式
        self.style.configure('Required.TLabel', 
                           foreground='#4285F4', 
                           font=('微软雅黑', 10, 'bold'))
        
        # 配置选填项样式
        self.style.configure('Optional.TLabel', 
                           foreground='#00A876',
                           font=('微软雅黑', 10))
        
        # 配置钻孔信息样式
        self.style.configure('Drilling.TLabel',
                           foreground='#FF7F2B',
                           font=('微软雅黑', 10, 'italic'))
        self.columnconfigure
        self.rowconfigure
        self.title("CAD柱状图识别")
        self.configure(padx=20, pady=20)  # Add padding to the window

        # Define fonts
        self.fontStyle = tkFont.Font(family="微软雅黑", size=11, weight="bold")

        # Group required fields in a LabelFrame
        # required_frame = ttkb.LabelFrame(self, text="必填项",style='TLabelframe', padding=10,relief='flat', borderwidth=0)
        # required_frame.grid(row=0, column=0, columnspan=2, sticky=W+E+N+S, padx=10, pady=5)

        required_fields_first_row = 1
        self.label00 = self.new_label_with_width("必填项字段名", 0, 0, 1, 1, 15,tooltip_text='下面绿色的前5项是程序识别柱状图的参照的基本框架，\n右侧填写的【cad中的名称】必须要跟cad中的对应上程序才能\n正常运行，否则会报错退出')
        self.label01 = self.new_label_with_width("CAD中的名称", 0, 1, 1, 1, 15,tooltip_text='在当前cad文件中左侧的目标字段叫什么？\n比如【柱状图标题】在遇到过的情况中可以叫：\n   ①钻孔柱状图\n   ②地质柱状图\n下拉列表中的项是我之前遇到过的表述，\n如果没有列表中没有对应的，\n请手动填写')
        
        # Required fields
        self.label1 = self.new_label("柱状图标题", required_fields_first_row, 0, 1, 1)
        self.combobox1 = self.new_combobox(
            ("钻孔柱状图", "地质柱状图"), required_fields_first_row, 1, 1, 1
        )
        self.label2 = self.new_label(
            "层底深度", required_fields_first_row + 1, 0, 1, 1,
            tooltip_text="对应CAD中表示地层深度的字段名\n例如：层底深度、深度等"
        )
        self.combobox2 = self.new_combobox(
            ("层底深度", "深度", "分层深度"), required_fields_first_row + 1, 1, 1, 1
        )
        self.label3 = self.new_label("时代成因", required_fields_first_row + 2, 0, 1, 1)
        self.combobox3 = self.new_combobox(
            ("时代成因", "成因时代", "时代与成因", "成因与时代", "地层时代"),
            required_fields_first_row + 2, 1, 1, 1
        )
        self.label4 = self.new_label("剖面层号", required_fields_first_row + 3, 0, 1, 1)
        self.combobox4 = self.new_combobox(
            ("地层编号", "剖面层号"), required_fields_first_row + 3, 1, 1, 1
        )
        self.label5 = self.new_label("岩土描述", required_fields_first_row + 4, 0, 1, 1)
        self.combobox5 = self.new_combobox(
            (
                "岩土名称及其特征",
                "岩土名称及描述",
                "岩土描述",
                "岩土特征描述",
                "岩性描述",
                "岩土名称及其描述",
                "地层名称及其特征",
                "地层描述",
                "地层特征描述",
                "地层名称及描述",
                "地层名称及其描述",
            ),
            required_fields_first_row + 4, 1, 1, 1
        )
        self.label6 = self.new_label("两侧宽度增加值", required_fields_first_row + 5, 0, 1, 1,tooltip_text='每个柱状图最外侧边框坐标范围为\n这一钻孔的搜索范围，若有些文字超出外框一点，\n搜索目标属性时可能会被排除掉\n填写增加值可以在原外框范围\n基础上扩大搜索范围，\ncad中移动鼠标估算坐标增量')
        self.entry11 = self.new_entry(0, required_fields_first_row + 5, 1, 1, 1)
        self.label7 = self.new_label("底部高度增加值", required_fields_first_row + 6, 0, 1, 1,'')
        self.entry12 = self.new_entry(0, required_fields_first_row + 6, 1, 1, 1)
        self.label8 = self.new_label("插入点或中心点", required_fields_first_row + 7, 0, 1, 1,tooltip_text='两种判断文字对象所处位置的方式，详见GitHub说明')
        self.combobox6 = self.new_combobox(
            ("中心点", "插入点"), required_fields_first_row + 7, 1, 1, 1
        )
        self.label9 = self.new_label("岩土描述排序方式", required_fields_first_row + 8, 0, 1, 1,tooltip_text='岩土描述的拼接方式，测试\n时在结果文件中查看拼接效果，\n如发生文字乱序可以切换另外\n两种方式再试')
        self.combobox7 = self.new_combobox(
            ("A", "B", "C"), required_fields_first_row + 8, 1, 1, 1
        )
        self.label10 = self.new_label("土工标贯形式", required_fields_first_row + 9, 0, 1, 1)
        self.combobox8 = self.new_combobox(
            ("分数", "单行"), required_fields_first_row + 9, 1, 1, 1
        )

        # Group optional fields in a LabelFrame
        optional_frame = ttkb.LabelFrame(self, text="选填项", padding=10,relief='flat', borderwidth=0)
        optional_frame.grid(row=11, column=0, columnspan=2, sticky=W+E+N+S, padx=10, pady=5)

        self.label120 = self.new_label_with_width("选填项字段名", required_fields_first_row + 10, 0, 1, 1, 15,tooltip_text='下面是读取土工标贯力学数据等无横线分隔的\n字符串（一般位于柱状图右下区域）\n这一列是自定义的名称，\n在输出的结果文件中会成为首行的字段名称')
        self.label121 = self.new_label_with_width("CAD中的名称", required_fields_first_row + 10, 1, 1, 1, 15,tooltip_text='填写cad中对应左边字段名称的表述（可局部匹配字符串）')
        self.entry1 = self.new_entry("标贯", required_fields_first_row + 11, 0, 1, 1)
        self.entry2 = self.new_entry("标贯", required_fields_first_row + 11, 1, 1, 1)
        self.entry3 = self.new_entry("取样", required_fields_first_row + 12, 0, 1, 1)
        self.entry4 = self.new_entry("取样", required_fields_first_row + 12, 1, 1, 1)
        self.entry5 = self.new_entry("力学数据", required_fields_first_row + 13, 0, 1, 1)
        self.entry6 = self.new_entry("力学数据", required_fields_first_row + 13, 1, 1, 1)
        self.entry7 = self.new_entry("稳定水位", required_fields_first_row + 14, 0, 1, 1)
        self.entry8 = self.new_entry("稳定水位", required_fields_first_row + 14, 1, 1, 1)
        self.entry9 = self.new_entry("", required_fields_first_row + 15, 0, 1, 1)
        self.entry10 = self.new_entry("", required_fields_first_row + 15, 1, 1, 1)

        # Group drilling information in a LabelFrame
        # drilling_frame = ttkb.LabelFrame(self, text="钻孔信息", padding=10)
        # drilling_frame.grid(row=0, column=2, columnspan=2, sticky=W+E+N+S, padx=10, pady=5, rowspan=11)
        # drilling_frame.columnconfigure(2, weight=3)  # Wider column for field names
        # drilling_frame.columnconfigure(3, weight=1)  # Narrower column for navigation
        self.label02 = self.new_label_with_width("钻孔信息字段(CAD)", 0, 2, 1, 1, 16,tooltip_text='下面的项对应钻孔柱状图顶部的钻孔基本信息部分，\n包括钻孔编号、孔口标高等，\n按照当前cad文件中的表述填写即可，\n同时在右侧要为目标属性选择对应的方向\n\n这个部分不会影响程序运行\n没有对应填写只是会读取不出目标属性，\n但程序依然能够成功运行到底\n\n❗可局部匹配字符串，比如cad中叫【孔口高程（m）】，\n这里可以填写：\n【孔口高程（m）】（完整表述）\n【孔口高】\n【口高】\n【口高程】\n【口高程（m】\n......')
        self.label03 = self.new_label_with_width("方向", 0, 3, 1, 1, 5,tooltip_text='下面是一个钻孔信息区域的常见布局\n👇👇👇👇👇👇👇👇👇\n ---------------------\n | 钻孔深度 | 13.5m |\n ---------------------\n |    100m  |    坐标  | \n---------------------\n在上方的表格中，钻孔深度为13.5m。\n所以左侧字段名填写【钻孔深度】，\n由于目标属性【13.5m】位于它的右侧，\n所以这里的方向就选择【右】。\n\n💡😀【13.5m】位于【坐标】的上方，\n所以字段名填写【坐标】，方向选择【上】，也\n可以读取到“13.5m”\n\n同理，【钻孔深度】+【下】可以读取到“100m”\n\n如果选择【标题】，则会在【柱状图标题】对应\n目标字符对象与柱状图外框之间搜索，\n主要针对一些靠近标题但没有被外框框住的字符串\n\n【左】这个方向没有实现是因为感觉用不着。（懒）')
        self.one2one_field_name_1 = self.new_combobox(("钻孔编号", "孔号", "勘探孔编号"), 1, 2, 1, 1)
        self.one2one_navigation_1 = self.new_combobox_with_width(("右", "下", "上", "标题"), 1, 3, 1, 1)
        self.one2one_field_name_2 = self.new_combobox(("孔口高程", "孔口标高", "勘探孔标高", "勘探孔高程", "标高"), 2, 2, 1, 1)
        self.one2one_navigation_2 = self.new_combobox_with_width(("右", "下", "上", "标题"), 2, 3, 1, 1)
        self.one2one_field_name_3 = self.new_combobox(("开孔日期", "开工日期"), 3, 2, 1, 1)
        self.one2one_navigation_3 = self.new_combobox_with_width(("右", "下", "上", "标题"), 3, 3, 1, 1)
        self.one2one_field_name_4 = self.new_combobox(("终孔日期", "竣工日期"), 4, 2, 1, 1)
        self.one2one_navigation_4 = self.new_combobox_with_width(("右", "下", "上", "标题"), 4, 3, 1, 1)
        self.one2one_field_name_5 = self.new_combobox(("钻孔深度", "勘探深度", "深度"), 5, 2, 1, 1)
        self.one2one_navigation_5 = self.new_combobox_with_width(("右", "下", "上", "标题"), 5, 3, 1, 1)
        self.one2one_field_name_6 = self.new_combobox(("坐", "坐标", "标"), 6, 2, 1, 1)
        self.one2one_navigation_6 = self.new_combobox_with_width(("右", "下", "上", "标题"), 6, 3, 1, 1)
        self.one2one_field_name_7 = self.new_combobox(("初见水位深度", "初见水位"), 7, 2, 1, 1)
        self.one2one_navigation_7 = self.new_combobox_with_width(("右", "下", "上", "标题"), 7, 3, 1, 1)
        self.one2one_field_name_8 = self.new_combobox(("稳定水位深度", "稳定水位", "静止水位深度", "静止水位", "静止深度"), 8, 2, 1, 1)
        self.one2one_navigation_8 = self.new_combobox_with_width(("右", "下", "上", "标题"), 8, 3, 1, 1)
        self.one2one_field_name_9 = self.new_combobox(("里程"), 9, 2, 1, 1)
        self.one2one_navigation_9 = self.new_combobox_with_width(("右", "下", "上", "标题"), 9, 3, 1, 1)
        self.one2one_field_name_10 = self.new_combobox(("工程名称"), 10, 2, 1, 1)
        self.one2one_navigation_10 = self.new_combobox_with_width(("右", "下", "上", "标题"), 10, 3, 1, 1)
        self.one2one_field_name_11 = self.new_combobox(("工点名称"), 11, 2, 1, 1)
        self.one2one_navigation_11 = self.new_combobox_with_width(("右", "下", "上", "标题"), 11, 3, 1, 1)
        self.one2one_field_name_12 = self.new_entry("", 12, 2, 1, 1)
        self.one2one_navigation_12 = self.new_combobox_with_width(("右", "下", "上", "标题"), 12, 3, 1, 1)
        self.one2one_field_name_13 = self.new_entry("", 13, 2, 1, 1)
        self.one2one_navigation_13 = self.new_combobox_with_width(("右", "下", "上", "标题"), 13, 3, 1, 1)
        self.one2one_field_name_14 = self.new_entry("", 14, 2, 1, 1)
        self.one2one_navigation_14 = self.new_combobox_with_width(("右", "下", "上", "标题"), 14, 3, 1, 1)
        self.one2one_field_name_15 = self.new_entry("", 15, 2, 1, 1)
        self.one2one_navigation_15 = self.new_combobox_with_width(("右", "下", "上", "标题"), 15, 3, 1, 1)
        self.one2one_field_name_16 = self.new_entry("", 16, 2, 1, 1)
        self.one2one_navigation_16 = self.new_combobox_with_width(("右", "下", "上", "标题"), 16, 3, 1, 1)
        # self.one2one_field_name_17 = self.new_entry("", 17, 2, 1, 1)
        # self.one2one_navigation_17 = self.new_combobox_with_width(("右", "下", "上", "标题"), 17, 3, 1, 1)

        # Buttons
        button_frame = ttkb.Frame(self)
        button_frame.grid(row=18, column=0, columnspan=4, pady=10, sticky=W+E)


        
        self.button3 = ttkb.Button(
            button_frame, text="GO", command=self.execute, style="primary.TButton", width=7
        )
        self.button3.pack(side=RIGHT, padx=5)

        # 创建“显示控制台”按钮
        self.toggle_console_button = ttkb.Button(
            button_frame,
            text="☰",
            command=self.toggle_console,
            bootstyle="info"
        )
        self.toggle_console_button.pack(side=RIGHT, padx=5)  # 将其放在“GO”按钮的右侧
        
        self.button4 = ttkb.Button(
            button_frame, text="退出", command=self.suicide, style="secondary.TButton", width=7
        )
        self.button4.pack(side=LEFT, padx=5)

        # 新增主题标签
        theme_label = ttkb.Label(
            button_frame,
            text="👚",
            font=("微软雅黑", 10),
            bootstyle="secondary"
        )
        theme_label.pack(side=LEFT, padx=(20,5))
        # ==== 新增主题切换下拉框（中） ====
        self.theme_cb = ttkb.Combobox(
            button_frame,
            values=Style().theme_names(),
            width=10,
            state="readonly"
        )
        self.theme_cb.set(self.style.theme_use())
        self.theme_cb.pack(side=tk.LEFT, padx=10, expand=True, fill=tk.X)  # 关键布局参数
        self.theme_cb.bind("<<ComboboxSelected>>", self.change_theme)

        # 中间组件组（包含主题切换和帮助图标）
        middle_components = ttkb.Frame(button_frame)
        middle_components.pack(side=LEFT, expand=True, fill=X, padx=2)
        # 帮助图标（紧跟在主题下拉框后）
        self.help_icon = ttkb.Label(
            middle_components,
            text="更多信息可访问👉 🔒GitHub",
            font=("Segoe UI Symbol", 10),
            cursor="hand2",
            bootstyle="secondary"
        )
        self.help_icon.pack(side=LEFT, padx=2, fill='y')
        QRCodeTooltip(self.help_icon, "https://github.com/he007209/Borehole-logs-Extractor-AutoCAD-")  # 绑定二维码  

        self.help_icon1 = ttkb.Label(
            middle_components,
            text="| 🌐Gitee",
            font=("Segoe UI Symbol", 10),
            cursor="hand2",
            bootstyle="secondary"
        )
        self.help_icon1.pack(side=LEFT, padx=2, fill='y')
        QRCodeTooltip(self.help_icon1, "https://gitee.com/he007209/Borehole-logs-Extractor-AutoCAD")  # 绑定二维码  




        # ---------- 1. 创建控制台区域，但先隐藏 ----------
        self.console_frame = ttkb.LabelFrame(self, text="程序输出", padding=10)
        # 暂时不 grid，等执行时再显示
        # self.console_frame.grid(row=0, column=4, rowspan=19, sticky=W+E+N+S, padx=10, pady=10)

        # 保证第 4 列有足够宽度（像素可调）
        # self.columnconfigure(4, minsize=650)
        # 创建带滚动条的文本框
        self.console_output = ttkb.ScrolledText(
            self.console_frame, 
            height=15,
    
            wrap=tk.WORD,

        )
        self.console_output.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 创建清空按钮
        self.clear_btn = ttkb.Button(
            self.console_frame,
            text="清空输出",
            command=self.clear_console,
            bootstyle=(OUTLINE, SECONDARY),
            width=10
        )
        self.clear_btn.pack(side=tk.RIGHT, padx=5, pady=5)
        
        # 保存原始标准输出
        self.original_stdout = sys.stdout
        self.original_stderr = sys.stderr  # 新增：保存原始stderr[11](@ref)
        # 重定向标准输出
        sys.stderr = self  # 新增：重定向标准错误
        sys.stdout = self
        
        # 添加输出区域提示
        Tooltip(self.clear_btn, "清空输出控制台内容")


   # 添加一个按钮，用于切换控制台的显示和隐藏
        # self.toggle_console_button = ttkb.Button(
        #     self,
        #     text="显示控制台",
        #     command=self.toggle_console,
        #     bootstyle="info"
        # )
        # self.toggle_console_button.grid(row=19, column=3, padx=10, pady=10, sticky=E)  # 放在右下角

        # global origin_width
        # origin_width = self.winfo_width()
        # print(origin_width)
        self.putmiddle()

    def create_theme_switcher(self):
        """创建主题切换组件"""
        # 主题切换面板容器
        theme_frame = ttkb.Frame(self)
        theme_frame.grid(row=0, column=4, padx=10, pady=5, sticky=NE)
        
        # 标签
        ttkb.Label(theme_frame, text="主题切换").grid(row=0, column=0, padx=5)
        
        # 主题选择下拉框
        self.theme_cb = ttkb.Combobox(
            theme_frame,
            values=Style().theme_names(),  # 获取所有可用主题
            width=15,
            state="readonly"
        )
        self.theme_cb.set(self.style.theme_use())  # 显示当前主题
        self.theme_cb.grid(row=0, column=1, padx=5)
        self.theme_cb.bind("<<ComboboxSelected>>", self.change_theme)
        
        # 主题预览按钮
        ttkb.Button(
            theme_frame,
            text="预览",
            command=self.preview_theme,
            bootstyle=PRIMARY
        ).grid(row=0, column=2, padx=5)

    def change_theme(self, event=None):
        """切换主题"""
        selected_theme = self.theme_cb.get()
        self.style.theme_use(selected_theme)
        
    def preview_theme(self):
        """预览主题时保持下拉框状态"""
        self.theme_cb.set(self.style.theme_use())
    def putmiddle(self):
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        global window_width,window_height
        window_width = self.winfo_width()
        window_height = self.winfo_height()
        self.geometry(
            f"{window_width}x{window_height}+{(screen_width - window_width) // 2}+{(screen_height - window_height) // 4}"
        )
    def load_theme(self):
        """读取保存的主题设置"""
        config_path = get_config_path()
        try:
            with open(config_path) as f:
                return json.load(f).get("theme")
        except FileNotFoundError:
            return "flatly"

    def save_theme(self):
        """保存当前主题"""
        config_path = get_config_path()
        print(config_path)
        try:
            with open(config_path, 'w') as f:
                json.dump({"theme": self.style.theme_use()}, f)
        except PermissionError:
            # 降级到临时目录写入
            temp_path = os.path.join(sys._MEIPASS, 'theme_pref.json')
            with open(temp_path, 'w') as f:
                json.dump({"theme": self.style.theme_use()}, f)

    def suicide(self):
        """程序退出时恢复标准输出"""
    # 恢复标准输出和错误
        if hasattr(self, 'original_stdout'):
            sys.stdout = self.original_stdout
        if hasattr(self, 'original_stderr'):  # 新增：检查属性是否存在[11](@ref)
            sys.stderr = self.original_stderr
        self.save_theme()
        self.destroy()
        exit()

    def execute(self):
        # 执行前显示控制台
        self.console_frame.grid(row=0, column=4, rowspan=19, sticky=W+E+N+S, padx=10, pady=10)
        self.toggle_console_button.config(text="☱")
        self.is_console_visible = True
        self.update_idletasks()  # 更新布局 
        global window_width, window_height
        # 1. 先扩展
        new_width = window_width + self.console_frame.winfo_width()  # 计算新宽度
        self.geometry(f"{new_width}x{self.winfo_height()}")  # 调整窗口宽度
        # self.geometry(f"{max(window_width, new_width)}x{max(window_height, 700)}")
        self.update_idletasks()
        # 2. 再显示
        self.console_frame.grid(row=0, column=4, rowspan=19,
                                sticky=W + E + N + S, padx=10, pady=10)
        self.update_idletasks()

        # 3. 强制全窗口重绘，去掉黑边
        self.after_idle(self.update)
        global required_list, optional_list
        required_list = []
        optional_list = []
        drilling_list = []
        common_setting_list = []
        # 必填项列表
        common_setting_list.append((self.label1.cget("text"), self.combobox1.get()))
        required_list.append((self.label2.cget("text"), self.combobox2.get()))
        required_list.append((self.label3.cget("text"), self.combobox3.get()))
        required_list.append((self.label4.cget("text"), self.combobox4.get()))
        required_list.append((self.label5.cget("text"), self.combobox5.get()))
        # 设置列表
        common_setting_list.append((self.label6.cget("text"), self.entry11.get()))
        common_setting_list.append((self.label7.cget("text"), self.entry12.get()))
        common_setting_list.append((self.label8.cget("text"), self.combobox6.get()))
        common_setting_list.append((self.label9.cget("text"), self.combobox7.get()))
        common_setting_list.append((self.label10.cget("text"), self.combobox8.get()))
        # 选填项列表
        optional_list.append((self.entry1.get(), self.entry2.get()))
        optional_list.append((self.entry3.get(), self.entry4.get()))
        optional_list.append((self.entry5.get(), self.entry6.get()))
        optional_list.append((self.entry7.get(), self.entry8.get()))
        optional_list.append((self.entry9.get(), self.entry10.get()))
        # 钻孔信息表列表
        drilling_list.append(
            (self.one2one_field_name_1.get(), self.one2one_navigation_1.get())
        )
        drilling_list.append(
            (self.one2one_field_name_1.get(), self.one2one_navigation_1.get())
        )
        drilling_list.append(
            (self.one2one_field_name_2.get(), self.one2one_navigation_2.get())
        )
        drilling_list.append(
            (self.one2one_field_name_3.get(), self.one2one_navigation_3.get())
        )
        drilling_list.append(
            (self.one2one_field_name_4.get(), self.one2one_navigation_4.get())
        )
        drilling_list.append(
            (self.one2one_field_name_5.get(), self.one2one_navigation_5.get())
        )
        drilling_list.append(
            (self.one2one_field_name_6.get(), self.one2one_navigation_6.get())
        )
        drilling_list.append(
            (self.one2one_field_name_7.get(), self.one2one_navigation_7.get())
        )
        drilling_list.append(
            (self.one2one_field_name_8.get(), self.one2one_navigation_8.get())
        )
        drilling_list.append(
            (self.one2one_field_name_9.get(), self.one2one_navigation_9.get())
        )
        drilling_list.append(
            (self.one2one_field_name_10.get(), self.one2one_navigation_10.get())
        )
        drilling_list.append(
            (self.one2one_field_name_11.get(), self.one2one_navigation_11.get())
        )
        drilling_list.append(
            (self.one2one_field_name_12.get(), self.one2one_navigation_12.get())
        )
        drilling_list.append(
            (self.one2one_field_name_13.get(), self.one2one_navigation_13.get())
        )
        drilling_list.append(
            (self.one2one_field_name_14.get(), self.one2one_navigation_14.get())
        )
        drilling_list.append(
            (self.one2one_field_name_15.get(), self.one2one_navigation_15.get())
        )
        drilling_list.append(
            (self.one2one_field_name_16.get(), self.one2one_navigation_16.get())
        )
        # drilling_list.append(
        #     (self.one2one_field_name_17.get(), self.one2one_navigation_17.get())
        # )
        # return required_list,optional_list,drilling_list
        optional_list = [item for item in optional_list if item[0] != ""]
        required_list = [item for item in required_list if item[0] != ""]
        drilling_list = [item for item in drilling_list if item[0] != ""]
        global title_name
        title_name = required_list[0][1]  # 柱状图标题
        required_reverse_list = [(item[1], item[0]) for item in required_list]
        optional_reverse_list = [(item[1], item[0]) for item in optional_list]
        required_reverse_dict = dict(required_reverse_list)
        optional_reverse_dict = dict(optional_reverse_list)
        drilling_dict = dict(drilling_list)

        global common_setting_dict, list_target_text_dict, target_text_dict, extend_width, extend_bottom_height, use_insertion_point, YTMS_sort_type, single_or_multiple_column
        common_setting_dict = dict(common_setting_list)
        list_target_text_dict = {**required_reverse_dict, **optional_reverse_dict}
        target_text_dict = drilling_dict
        extend_width = common_setting_dict["两侧宽度增加值"]
        extend_bottom_height = common_setting_dict["底部高度增加值"]
        use_insertion_point = common_setting_dict["插入点或中心点"]
        YTMS_sort_type = common_setting_dict["岩土描述排序方式"]
        single_or_multiple_column = common_setting_dict["土工标贯形式"]

        self.go()
        # self.destroy()

    def go(self):
        try:
            py_name = os.path.basename(__file__)  # 当前运行的py文件名
            py_folder = os.path.dirname(__file__)  # py文件所在路径
            py_path = __file__  # py文件完整路径
            # print(py_folder)
            # exit()
            #####################开始#######################
            #####################开始#######################
            #####################开始#######################
            #####################开始#######################
            #####################开始#######################
            if common_setting_dict["插入点或中心点"] == "中心点":
                new_dict = {"插入点或中心点": 0}
                common_setting_dict.update(new_dict)
            elif common_setting_dict["插入点或中心点"] == "插入点":
                new_dict = {"插入点或中心点": 1}
                common_setting_dict.update(new_dict)
            # print(common_setting_dict['插入点或中心点'])

            if common_setting_dict["岩土描述排序方式"] == "A":
                new_dict = {"岩土描述排序方式": 999}
                common_setting_dict.update(new_dict)
            elif common_setting_dict["岩土描述排序方式"] == "B":
                new_dict = {"岩土描述排序方式": 0}
                common_setting_dict.update(new_dict)
            elif common_setting_dict["岩土描述排序方式"] == "C":
                new_dict = {"岩土描述排序方式": 1}
                common_setting_dict.update(new_dict)
            # print(common_setting_dict['岩土描述排序方式'])

            if common_setting_dict["土工标贯形式"] == "分数":
                new_dict = {"土工标贯形式": 0}
                common_setting_dict.update(new_dict)
            elif common_setting_dict["土工标贯形式"] == "单行":
                new_dict = {"土工标贯形式": 1}
                common_setting_dict.update(new_dict)
            # print(common_setting_dict['土工标贯形式'])

            print(
                "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
            )
            print(
                "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
            )
            print(
                "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
            )
            start_time = time.time()
            start_times = get_current_time()
            print("开始时间：", start_times)
            try:
                acad = Autocad()
            except Exception as e:
                error_msg = f"CAD连接错误：{str(e)}\n请检查CAD是否打开"
                sys.stderr.write(error_msg)
                return
            print("开始前最好删除较短的直线和多段线:")
            print("     1、提高识别速度")
            print(
                "     2、有时可以避免影响识别结果,比如剖面层号如果用了直线作为连接线'-',就可能将它识别为【空】"
            )
            print("-------------------------------------------------------------")
            try:
                acad.prompt("Hello, Autocad from Python")
                # acad.SendStringToExecute(Chr(27)+Chr(27))
                # send_command_to_cad(acad,"ESC")
                cad_name = acad.doc.Name  # cad文件名
                cad_folder_path = acad.doc.path  # cad文件所在目录路径
            except Exception as e:
                print(
                    str(e),
                    "\n出错喇！原因可能是：\n   1、CAD文件没打开\n   2、CAD现在被占用了：不要编辑这个CAD文件，不要选中任何对象或工具，\n   返回CAD按个ESC应该就🆗(光标应变回中间带个小正方形的十字线)\n   还不行就试试重启CAD",
                )
                return
                # exit()
            print("Cad name: " + cad_name)
            # exit()
            # acad.prompt("删除重复对象...")
            # send_command_to_cad(acad,"-overkill all  p n d ")#删除重复项(不打断多段线)
            # time.sleep(3)
            # print("重复对象已删除(不打断多段线)")
            # time.sleep(3)
            #####################公共对象########################
            extend_width = float(
                common_setting_dict["两侧宽度增加值"]
            )  # max_frame_list,vertical_line_list和vertical_polyline_list里面的坐标x相应改动，有些坑爹的图文本对象插入点在外框外面，可以给外框适当加宽一点
            extend_bottom_height = float(
                common_setting_dict["底部高度增加值"]
            )  # 外框底部加高
            global title_name
            use_insertion_point = common_setting_dict[
                "插入点或中心点"
            ]  # 文本对象(含块参照名称)用插入点坐标(填 1 ),默认用中心点( 填 0 )
            YTMS_sort_type = common_setting_dict[
                "岩土描述排序方式"
            ]  # (默认为999)【岩土描述】重叠排序选项，如岩土描述重叠，看测试结果顺序来填: 0: 按先text_object_id先升序再depth_order_number降序排列(岩土描述列) ; 1 :按先text_object_id先降序再depth_order_number降序排列(岩土描述列)
            single_or_multiple_column = common_setting_dict[
                "土工标贯形式"
            ]  # 【标贯】【土工样】多列单行(1)或单列分数(0)
            title_name = common_setting_dict["柱状图标题"].replace(
                " ", ""
            )  # 柱状图标题(去空格)
            # list_target_text_dict = {
            #     "层底深度": "层底深度",  #必填项
            #     "地层编号": "剖面层号",  #必填项
            #     "时代成因": "时代成因",  #必填项
            #     "岩土名称及其特征": "岩土描述",  #必填项
            #     #下面为选填项(标贯、土工等)
            #     # "=深度": "深度",
            #     "取样位置": "取样位置",
            #     "标贯动探击数击": "标贯击数",
            #     "=柱":"取样标贯符号",
            #     "力学数据":""
            # }  #层底深度等多个文本对象参考字典:{CAD中的字段名，目标字段名}（CAD中的字段名分块会自动连接）字典键不要有空格
            # target_text_dict = {
            #     # "审  核": "下",
            #     "工程名称": "右",
            #     "工点名称": "右",
            #     "勘探孔编号": "右",
            #     "钻孔编号": "右",
            #     "静止水位": "下",
            #     # "孔口高程": "右",
            #     "勘探深度": "右",
            #     "开孔日期": "右",
            #     "终孔日期": "右",
            #     "坐": "右",  #此处如果是分开的，填单个字(上面的字典不用，有空改改)
            #     # "标": "右",
            #     "初见水位深度": "右",
            #     "稳定水位深度": "右",
            #     "里程": "右",
            #     "孔口高程(m)": "右",
            #     "X =": "下",
            #     "Y =": "上",
            #     # "孔  深": "下",
            #     # "开工日期": "右",
            #     # "竣工日期": "右",
            #     # "编  录": "上",
            #     # "制  图": "下@黄",  #下方格子中包含某个字符串的对象
            #     "勘察单位": "上",
            #     # "制  图": "上",
            #     '勘察单位': "标题@X"  #框外标题附近包含某个字符串的对象(键需为框内任意对象)   如【'勘察单位': "标题@X"】
            # }

            title_list = []  # 柱状图标题文本对象列表
            vertical_line_list = []  # 竖直线列表
            horizon_line_list = []  # 横直线列表
            vertical_polyline_with_range_list = []  # 竖多段线列表
            horizon_polyline_with_range_list = []  # 横多段线列表
            max_frame_list = []  # 外框范围列表
            drilling_number_list = []  # 钻孔编号列表
            range_id_drilling_number_chart_dict = {}  # 柱状图标识和钻孔编号对照字典
            out_frame_list = []
            in_frame_list = []
            text_list = []  # y=text对象列表
            # print(py_path)
            dwg_name = "ZZT_" + str(cad_name).split(".")[0]
            # result_path = str(dwg_path).replace('\\', '/') + '/' + dwg_name  #新建文件夹的路径
            # print(result_path)
            result_path = cad_folder_path  # 结果txt文件与CAD文件放在同级目录
            # print(result_path)
            # exit()
            result_path = result_path + "/" + dwg_name
            if not os.path.exists(result_path):
                os.makedirs(result_path)
            txt_path = result_path + "/" + dwg_name + ".txt"
            addition_txt_4_depth_description_path = (
                result_path + "/" + "层底深度_岩土描述数量对照表" + ".txt"
            )
            txt_add = open(addition_txt_4_depth_description_path, "a")
            txt_add.write("手动检查下面"+"【层底深度数量 > 岩土描述数量？：True】的钻孔，可以使用给出的定位命令直接跳转过去\n")
            # print(txt_path)
            # exit()
            txt = open(txt_path, "w")
            #############################################
            current_times = get_current_time()
            if use_insertion_point == 1:
                print("文本对象(含块参照)用【插入点】坐标")
            else:
                print("文本对象(含块参照)用【中心点】坐标")
            print("--------------------------------------------")
            print(str(current_times) + "  正在识别每个柱状图范围...")
            try:
                title_text = acad.find_one("Text", predicate=text_contains_str)
            except:
                print(
                    "这个文件读不了，可能是因为：1、您卸载了AutoDesk360，这样似乎需要重装CAD \n2、CAD版本太新了，反正2021不行，2014的可以，换个吧"
                )
                return
                exit()

            try:
                global title_text_height
                title_text_x = title_text.InsertionPoint[0]  # 标题对象x
                title_text_y = title_text.InsertionPoint[1]  # 标题对象y
                title_text_height = title_text.height  # 标题对象文字高度
            except:
                print(
                    '···找了这么久，都没找到："'
                    + title_name
                    + '",原因可能是:\n  A.这个CAD文件需要【全选(Ctrl+A)->"[ 分解 ]"】\n  B.这个CAD文件的柱状图标题项它就不叫"'
                    + title_name
                    + '"\n'
                )
                return
                # exit()
            ################################################
            # 遍历多段线对象，寻找面积最大的，得到单个 柱状图的坐标范围
            vertical_polyline_list = []  # 竖多段线列表
            horizon_polyline_list = []  # 横多段线列表
            max_frame_max_coor_polyline = (
                []
            )  # 每个柱状图的外框坐标最值列表（代表单个柱状图的范围；由面积最大的多段线确定）
            polyline_list = []  # 多段线面积、坐标列表
            polyline_area_list = []  # 多段线面积列表（用于找面积最大值）
            max_frame_polyline = (
                []
            )  # 每个柱状图的外框列表（代表单个柱状图的范围；由面积最大的多段线确定）
            for polyline in acad.iter_objects("Polyline"):
                polyline_visibility = polyline.Visible
                polyline_area = polyline.area  # 多段线面积
                polyline_coordinates = polyline.Coordinates  # 多段线顶点坐标
                if polyline.objectName == "AcDb2dPolyline":
                    x_list = polyline_coordinates[::3]  # (x)
                    y_list = polyline_coordinates[1::3]  # (y)
                    c = list(chain.from_iterable(zip(x_list, y_list)))
                    polyline_coordinates = c
                if polyline_visibility == True:
                    polyline_list.append(
                        (polyline_area, polyline_coordinates)
                    )  # 多段线面积、坐标列表
                else:
                    print("隐藏多段线对象：", polyline_coordinates)
                # if polyline_area != 0:
                #     print("面积："+str(polyline_area)+" 坐标："+str(polyline_coordinates[0])+str(polyline_coordinates[1]))

                polyline_area_list.append(int(polyline_area))  # 多段线面积
                #########################寻找用多段线来画的竖直和水平直线
                if polyline_area == 0:  # 面积为零，说明不封闭
                    # print('这他妈是条多段线',polyline_coordinates,len(polyline_coordinates))
                    # polyline_x = polyline_coordinates[::2]#坐标组奇数位(x坐标)
                    # polyline_y = polyline_coordinates[1::2]#坐标组偶数位(y坐标)
                    point_num = len(polyline_coordinates)  # 多段线节点数
                    start_point_x = polyline_coordinates[0]  # 端点x坐标
                    start_point_y = polyline_coordinates[1]  # 端点y坐标
                    end_point_x = polyline_coordinates[point_num - 2]  # 另一个端点x坐标
                    end_point_y = polyline_coordinates[point_num - 1]  # 另一个端点y坐标
                    if round(start_point_x, 2) == round(end_point_x, 2):
                        vertical_polyline_list.append(
                            (start_point_x, start_point_y, end_point_x, end_point_y)
                        )
                    if round(start_point_y, 2) == round(end_point_y, 2):
                        horizon_polyline_list.append(
                            (start_point_x, start_point_y, end_point_x, end_point_y)
                        )
                        # print('竖多段线：',point_num,start_point_x,start_point_y,end_point_x,end_point_y)
            if (
                len(polyline_area_list) != 0
            ):  # 多段线面积不为零，外框用多段线绘制，直线另外判断
                print("-------------------柱状图外框由多段线绘制-----------------------")
                polyline_area_list = list(set(polyline_area_list))  # 面积唯一值
                polyline_area_list.sort(reverse=True)  # 面积降序排列
                if len(polyline_area_list) > 1:
                    area_list = get_max_and_second_max_list(polyline_area_list, 0.9)
                    max_polyline_area_list = area_list[
                        0
                    ]  # 1.多段线面积最大值列表(整数)(相近值视为同种多段线)
                    second_max_polyline_area_list = area_list[
                        1
                    ]  # 2.多段线面积次大值列表(整数)(相近值视为同种多段线)
                else:
                    max_polyline_area_list = polyline_area_list

                if extend_width != 0:
                    print(
                        "注意！为了读取部分插入点在框外的文本对象，现在外框左右分别增加了宽度："
                        + str(extend_width),
                        "\n如柱状图文本对象插入点均在外框内，为了更准确的处理，应将变量extend_width值改为0！",
                    )
                if extend_bottom_height != 0:
                    print(
                        "注意！为了读取部分插入点在框外的文本对象，现在外框底部分别增加了高度："
                        + str(extend_bottom_height),
                        "\n这样可能导致底部审核、编制等内容插入岩土描述中！如柱状图文本对象插入点均在外框内，为了更准确的处理，应将变量 extend_bottom_height 值改为0！",
                    )
                print("^^^^^^^^^^^^^^^^^^^^")

                out_frame = 0
                for polyline in polyline_list:
                    x_list = polyline[1][::2]  # 奇数项(x)
                    y_list = polyline[1][1::2]  # 偶数项(y)
                    min_x = min(x_list)
                    max_x = max(x_list)
                    min_y = min(y_list)
                    max_y = max(y_list)
                    if (
                        int(polyline[0]) in max_polyline_area_list
                        and title_text_y < max_y
                        and title_text_y > min_y
                        and title_text_x > min_x
                        and title_text_x < max_x
                    ):
                        print("【柱状图有外框】")
                        # 后续用second_max_polyline_area筛选多段线
                        out_frame = 1
                        break
                if out_frame == 0:
                    print("【柱状图无外框】")
                    # 后续用max_polyline_area筛选多段线
                print("^^^^^^^^^^^^^^^^^^^^")
                if out_frame == 1:
                    max_polyline_area_list = second_max_polyline_area_list  # 柱状图如果有外框，就用面积次大值筛选多段线

                for polyline in polyline_list:
                    if int(polyline[0]) in max_polyline_area_list:
                        max_frame_polyline.append(polyline)  # (面积,(各顶点坐标))
                #         print(polyline)
                # exit()
                frame_x_range = []  # 每个柱状图x范围
                frame_y_range = []  # 每个柱状图y范围
                for frame in max_frame_polyline:  # 每个柱状图坐标最值
                    x_list = frame[1][::2]  # 奇数项(x)
                    y_list = frame[1][1::2]  # 偶数项(y)
                    min_x = min(x_list)
                    max_x = max(x_list)
                    min_y = min(y_list)
                    max_y = max(y_list)
                    max_frame_max_coor_polyline.append(
                        (min_x, min_y, max_x, max_y)
                    )  # 对角坐标(x最小值,y最小值,x最大值,y最大值)
                #     print(min_x, min_y, max_x, max_y)
                # exit()
                max_frame_max_coor_polyline = set(max_frame_max_coor_polyline)
                sorted_by_min_x_max_frame_max_coor_polyline = enumerate(
                    sorted(max_frame_max_coor_polyline, key=lambda x: x[0]), 1
                )  # 外框列表按x坐标升序排序，并添加序号
                for coor in sorted_by_min_x_max_frame_max_coor_polyline:
                    range_id = coor[0]
                    range_min_x = coor[1][0]
                    range_min_y = coor[1][1]
                    range_max_x = coor[1][2]
                    range_max_y = coor[1][3]
                    max_frame_list.append(
                        (
                            range_id,
                            range_min_x - extend_width,
                            range_min_y - extend_bottom_height,
                            range_max_x + extend_width,
                            range_max_y + title_text_height,
                        )
                    )  # 加了坑爹的外扩范围
                max_frame_list = set(max_frame_list)  # 防止相同重叠多段线，删除重复项
                max_frame_list = sorted(
                    max_frame_list, key=lambda x: x[0]
                )  # 按range_ID升序排序
            else:
                print(
                    "-------------------柱状图外框由直线绘制-----------------------"
                )  # 也有可能是多段线画的直线，先不写了
                print("正在遍历直线对象，识别每个柱状图范围...")
                current_times = get_current_time()
                # 遍历直线对象，得到竖线和横线列表
                jr_ver_line_list = []  # 不带range_id的竖线列表
                jr_hor_line_list = []  # 不带range_id的横线列表
                for line in acad.iter_objects("AcDbLine"):
                    # print(line.objectName)
                    line_visibility = line.Visible
                    start_point = line.StartPoint  # 起点坐标
                    end_point = line.EndPoint  # 终点坐标
                    rad_line_angle = line.Angle  # 直线角度(弧度)
                    line_angle = rad_line_angle * 180 / math.pi
                    angle_180_remainder = abs(
                        (line_angle + 180) % 180
                    )  # 角度除以180取余数(绝对值)
                    if angle_180_remainder == 0 and line_visibility == True:
                        # 横线
                        # start_x = start_point[0]

                        jr_hor_line_list.append(
                            (start_point[0], start_point[1], end_point[0], end_point[1])
                        )
                        # print(start_point[0], start_point[1], end_point[0], end_point[1])
                    if angle_180_remainder == 90 and line_visibility == True:
                        # 竖线
                        jr_ver_line_list.append(
                            (start_point[0], start_point[1], end_point[0], end_point[1])
                        )
                # exit()
                # 遍历柱状图标题，从标题找下方最近一条直线
                for text in acad.iter_objects("AcDbText"):
                    text_visibility = text.Visible
                    text_content = text.TextString  # 文本对象内容
                    text_insert_coordinate = text.InsertionPoint  # 文本对象插入点坐标
                    if (
                        text_content.replace(" ", "") == title_name
                        and text_visibility == True
                    ):  # 寻找标题对象
                        title_list.append(text_insert_coordinate)

                for title in enumerate(title_list):
                    range_id = title[0] + 1  # 从1开始计数
                    title_x = title[1][0]
                    title_y = title[1][1]
                    # print(range_id,title_x,title_y)
                    # exit()
                    jr_out_frame_hor_line_list = []  # 柱状图外框横线列表
                    for hor_line in jr_hor_line_list:
                        # print("合格西南：：：", hor_line)
                        hor_line_start_point_x = hor_line[0]
                        hor_line_start_point_y = hor_line[1]
                        hor_line_end_point_x = hor_line[2]
                        hor_line_end_point_y = hor_line[3]

                        # #奇葩啊，下面改成小的x为起点，大的x为终点先
                        # if hor_line_start_point_x > hor_line_end_point_x:
                        #     temp = hor_line_end_point_x
                        #     hor_line_end_point_x = hor_line_start_point_x
                        #     hor_line_start_point_x = hor_line_end_point_x

                        min_x = min(hor_line_start_point_x, hor_line_end_point_x)
                        max_x = max(hor_line_start_point_x, hor_line_end_point_x)
                        if (
                            title_x > min_x
                            and title_x < max_x
                            and title_y > hor_line_start_point_y
                        ):
                            jr_out_frame_hor_line_list.append(
                                hor_line
                            )  # 柱状图标题下方横线集合
                    # 外框顶部横线
                    for hor_line in jr_out_frame_hor_line_list:
                        print("horizon line:  " + str(hor_line))
                    top_hor_line = sorted(
                        jr_out_frame_hor_line_list, key=lambda x: [-x[1], x[0]]
                    )[
                        0
                    ]  # 按横线y坐标降序，再按横线x坐标升序(排除重叠的较短的横线)取首位的横线

                    # print("top horizon line: "+str(top_hor_line))

                    range_min_x = min(top_hor_line[0], top_hor_line[2])  # 顶部横线左端点x
                    range_max_x = max(top_hor_line[0], top_hor_line[2])  # 顶部横线右端点x
                    range_max_y = top_hor_line[1]  # 顶部横线y
                    # 外框左侧竖线
                    print("x值最小值：" + str(range_min_x))
                    print("y值最大值：" + str(range_max_y))

                    for line in jr_ver_line_list:
                        print(line)
                    # exit()
                    left_ver_line = [
                        line
                        for line in jr_ver_line_list
                        if line[0] == range_min_x
                        and (line[1] == range_max_y or line[3] == range_max_y)
                    ][0]
                    # 外框右侧竖线
                    left_ver_line = [
                        line
                        for line in jr_ver_line_list
                        if line[0] == range_max_x
                        and (line[1] == range_max_y or line[3] == range_max_y)
                    ][0]
                    range_min_y = min(left_ver_line[1], left_ver_line[3])  # 底部横线y
                    # print('title coors:', (title_x, title_y), 'range:',
                    #       (range_min_x,range_min_y,range_max_x,range_max_y))

                    max_frame_list.append(
                        (
                            range_id,
                            range_min_x - extend_width,
                            range_min_y - extend_bottom_height,
                            range_max_x + extend_width,
                            range_max_y + title_text_height,
                        )
                    )  # 加了坑爹的外扩范围 #底部范围未增加
                    max_frame_list = set(max_frame_list)  # 防止相同重叠多段线，删除重复项
                    max_frame_list = sorted(
                        max_frame_list, key=lambda x: x[0]
                    )  # 按range_ID升序排序
                # for range in max_frame_list:
                #     print(range)
            range_id_list = [id[0] for id in max_frame_list]
            range_id_drilling_number_chart_dict = dict(zip(range_id_list, range_id_list))
            # for key,value in range_id_drilling_number_chart_dict.items():
            #     print(key,value)
            # exit()

            txt.write(
                "******************************柱状图识别*********************************"
                + "\n"
            )
            max_frame_count = len(max_frame_list)  # 柱状图数量
            print(
                "在CAD文件【"
                + str(cad_name)
                + "】中识别出柱状图"
                + str(max_frame_count)
                + "个"
            )
            print(
                "************************************************************************"
            )
            txt.write(
                "在CAD文件【"
                + str(cad_name)
                + "】中识别出柱状图"
                + str(max_frame_count)
                + "个"
                + "\n"
            )
            txt.write(
                "************************************************************************"
                + "\n"
            )
            for frame in max_frame_list:
                print(
                    "【柱状图标识】",
                    frame[0],
                    '定位命令："zoom '
                    + str(frame[1])
                    + ","
                    + str(frame[2])
                    + " "
                    + str(frame[3])
                    + ","
                    + str(frame[4])
                    + ' "(是的，最后有个空格)',
                )
                txt.write(
                    "柱状图标识："
                    + str(frame[0])
                    + ' 定位命令："zoom '
                    + str(frame[1])
                    + ","
                    + str(frame[2])
                    + " "
                    + str(frame[3])
                    + ","
                    + str(frame[4])
                    + ' "(复制双引号内的字符串，在CAD中粘贴，可以定位到对应柱状图)'
                    + "\n"
                )
            print(
                "************************************************************************"
            )

            outframe_time = time.time()
            outframe_times = get_current_time()
            time_cnsumption = format(outframe_time - start_time, ".2f")
            print(
                "开始时间："
                + start_times
                + "\n"
                + "完成柱状图外框识别时间："
                + outframe_times
                + "\n"
                + "耗时："
                + str(time_cnsumption)
                + "秒"
            )
            txt.write(
                "开始时间："
                + start_times
                + "\n"
                + "完成柱状图外框识别时间："
                + outframe_times
                + "\n"
                + "耗时："
                + str(time_cnsumption)
                + "秒"
                + "\n"
            )
            print(
                "************************************************************************"
            )
            txt.write(
                "************************************************************************"
                + "\n"
            )
            #############################################
            current_times = get_current_time()
            print(str(current_times) + "  正在遍历多段线对象，寻找水平和竖直多段线...")
            # 遍历多段线画的直线对象，得到竖多段线和横多段线列表(包含范围id)
            for ver_line in vertical_polyline_list:
                start_point = (ver_line[0], ver_line[1])
                end_point = (ver_line[2], ver_line[3])
                line_range_id_list = line_adscription(
                    start_point, end_point, max_frame_list
                )
                if len(line_range_id_list) == 1:
                    line_range_id = line_range_id_list[0]
                else:
                    continue

                out_frame_min_x = [
                    frame[1] for frame in max_frame_list if frame[0] == line_range_id
                ][0] + extend_width
                out_frame_max_x = [
                    frame[3] for frame in max_frame_list if frame[0] == line_range_id
                ][0] - extend_width

                if start_point[0] != out_frame_min_x and start_point[0] != out_frame_max_x:
                    vertical_polyline_with_range_list.append(
                        (
                            start_point[0],
                            start_point[1],
                            end_point[0],
                            end_point[1],
                            line_range_id,
                        )
                    )

            for hor_line in horizon_polyline_list:
                start_point = (hor_line[0], hor_line[1])
                end_point = (hor_line[2], hor_line[3])
                line_range_id_list = line_adscription(
                    start_point, end_point, max_frame_list
                )
                if len(line_range_id_list) == 1:
                    line_range_id = line_range_id_list[0]
                else:
                    continue
                horizon_polyline_with_range_list.append(
                    (
                        start_point[0],
                        start_point[1],
                        end_point[0],
                        end_point[1],
                        line_range_id,
                    )
                )
            #############################################
            current_times = get_current_time()
            print(str(current_times) + "  正在遍历直线对象，寻找水平和竖直直线...")
            # 遍历直线对象，得到竖线和横线列表
            for line in acad.iter_objects("AcDbLine"):
                # print(line.objectName)
                line_visibility = line.Visible
                start_point = line.StartPoint  # 起点坐标
                end_point = line.EndPoint  # 终点坐标
                rad_line_angle = line.Angle  # 直线角度(弧度)
                line_angle = rad_line_angle * 180 / math.pi
                # print(start_point,end_point,line_angle)
                if round(start_point[0], 5) == round(
                    end_point[0], 5
                ):  # line_angle % 90 == 0:##竖线判断(202105250933本想改成用角度来判断，但是)
                    line_range_id_list = line_adscription(
                        start_point, end_point, max_frame_list
                    )
                    if len(line_range_id_list) == 1:
                        line_range_id = line_range_id_list[0]
                    else:
                        continue
                    # 去掉与外框重合的竖线对象
                    out_frame_min_x = [
                        frame[1] for frame in max_frame_list if frame[0] == line_range_id
                    ][0] + extend_width
                    out_frame_max_x = [
                        frame[3] for frame in max_frame_list if frame[0] == line_range_id
                    ][0] - extend_width

                    if (
                        start_point[0] != out_frame_min_x
                        and start_point[0] != out_frame_max_x
                        and line_visibility == True
                    ):
                        vertical_line_list.append(
                            (
                                start_point[0],
                                start_point[1],
                                end_point[0],
                                end_point[1],
                                line_range_id,
                            )
                        )
                if round(start_point[1], 5) == round(
                    end_point[1], 5
                ):  # line_angle == 0 or line_angle % 360 == 0:##横线判断(202105250933本想改成用角度来判断，但是)
                    # print(start_point[0], start_point[1], end_point[0],
                    #                           end_point[1], line_range_id)
                    line_range_id_list = line_adscription(
                        start_point, end_point, max_frame_list
                    )
                    if len(line_range_id_list) == 1:
                        line_range_id = line_range_id_list[0]
                    else:
                        continue
                    # 去掉与外框重合的横线对象
                    out_frame_min_y = [
                        frame[2] for frame in max_frame_list if frame[0] == line_range_id
                    ][0] + extend_bottom_height
                    out_frame_max_y = [
                        frame[4] for frame in max_frame_list if frame[0] == line_range_id
                    ][0] - title_text_height

                    if (
                        start_point[1] != out_frame_min_y
                        and start_point[1] != out_frame_max_y
                        and line_visibility == True
                    ):
                        horizon_line_list.append(
                            (
                                start_point[0],
                                start_point[1],
                                end_point[0],
                                end_point[1],
                                line_range_id,
                            )
                        )
                    # print(start_point[0], start_point[1], end_point[0],
                    #                           end_point[1], line_range_id)
            #############################################
            current_times = get_current_time()
            print(
                str(current_times)
                + "  正在遍历块参照对象，将会把块参照对象的名称合并到文本对象列表中一并处理..."
            )
            # 遍历块参照对象，以对象名称为内容，将块参照合并到下面的文本对象列表中
            BlockReference_list = []
            for obj in acad.iter_objects("AcDbBlockReference"):
                # print(dir(obj))
                ref_block_visibility = obj.Visible
                ref_block_content = (
                    "BlockReference_" + obj.Name
                )  # 将块参照对象名称视作其内容，加个前缀标记下
                ref_block_insertion_coordinates = obj.InsertionPoint
                ref_block_objectid = obj.ObjectID

                try:
                    ref_bounding_box = (
                        obj.GetBoundingBox()
                    )  # 包络矩形对角坐标,有些没有这个属性
                except:
                    continue
                min_x = ref_bounding_box[0][0]
                min_y = ref_bounding_box[0][1]
                max_x = ref_bounding_box[1][0]
                max_y = ref_bounding_box[1][1]
                center_point_coordinate = (
                    (min_x + max_x) / 2.0,
                    (min_y + max_y) / 2.0,
                    0.0,
                )
                point_coordinate = center_point_coordinate
                if use_insertion_point == 1:
                    point_coordinate = ref_block_insertion_coordinates

                ref_block_range_id_list = point_adscription(
                    point_coordinate, max_frame_list
                )
                if len(ref_block_range_id_list) == 1:
                    ref_block_range_id = ref_block_range_id_list[0]
                else:
                    if len(ref_block_range_id_list) != 0:
                        print(len(ref_block_range_id_list))
                        print(
                            "该文本对象在多个范围中！ ["
                            + str(ref_block_content)
                            + "]"
                            + str(point_coordinate[0])
                            + ","
                            + str(point_coordinate[1])
                        )
                    continue
                if ref_block_visibility == True:
                    BlockReference_list.append(
                        (
                            ref_block_range_id,
                            ref_block_content,
                            point_coordinate,
                            ref_block_objectid,
                            ref_bounding_box,
                        )
                    )

            #############################################
            current_times = get_current_time()
            print(str(current_times) + "  正在遍历文本对象(包括多行文本对象MText)...")
            # 遍历cad文本对象，寻找目标字符串对象
            for text in acad.iter_objects(
                ["AcDbText", "AcDbMText"]
            ):  # ,'AcDbMText']):  #遍历'Text'类型对象并输出内容和坐标
                text_content = text.TextString.strip()  # 文本对象内容
                # print("原版content", text_content)
                # print(dir(text))
                # exit()
                text_type = text.EntityType
                text_visibility = (
                    text.Visible
                )  # 可见性，用来排除隐藏的文字对象（SD/START/END/CXH...）
                if text_type == 21:
                    text_content = utils.unformat_mtext(text_content)
                    text_content = text_content.replace("\\P", "").strip()
                # if text_type == 21 and (";" in text_content or "\\P" in text_content):#Mtext
                #     try:
                #         text_content = text_content.replace("\\P", "").strip()
                #         text_content = re.sub(r"\{.*?\}", "", text_content)
                #         print("得左？")
                #     except:
                #         print("咩事啊")
                #         continue
                # print("替换后content", text_content, text_type)
                # exit()
                text_insert_coordinate = text.InsertionPoint  # 文本对象插入点坐标
                try:
                    text_bounding_box = text.GetBoundingBox()  # 包络矩形对角坐标
                except:
                    continue
                min_x = text_bounding_box[0][0]
                min_y = text_bounding_box[0][1]
                max_x = text_bounding_box[1][0]
                max_y = text_bounding_box[1][1]
                center_point_coordinate = (
                    (min_x + max_x) / 2.0,
                    (min_y + max_y) / 2.0,
                    0.0,
                )
                point_coordinate = center_point_coordinate
                if use_insertion_point == 1:
                    point_coordinate = text_insert_coordinate
                text_object_id = text.ObjectID
                text_range_id_list = point_adscription(point_coordinate, max_frame_list)
                if len(text_range_id_list) == 1:
                    text_range_id = text_range_id_list[0]
                else:
                    if len(text_range_id_list) != 0:
                        print(len(text_range_id_list))
                        print(
                            "该文本对象在多个范围中！ ["
                            + str(text_content)
                            + "]"
                            + str(point_coordinate[0])
                            + ","
                            + str(point_coordinate[1])
                        )
                    continue
                # 根据可见性排除隐藏对象
                if text_visibility == True:
                    text_list.append(
                        (
                            text_range_id,
                            text_content,
                            point_coordinate,
                            text_object_id,
                            text_bounding_box,
                        )
                    )
                else:
                    print("隐藏文字对象：", text_content, " 包络矩形: ", text_bounding_box)

                # get_string_list('工点名称', text_content, text_insert_coordinate,
                #                 text_range_id, title_list)  #获取工点名称列表

            ########################################################
            text_list = text_list + BlockReference_list  # 将块参照名称列表并入文本对象列表
            # for text in text_list:
            #     print(text)
            #########################################################
            target_text_dict_keys_list = [
                no_space_str.replace(" ", "")
                for no_space_str in list(target_text_dict.keys())
            ]  # 表头单一对象字典键列表
            target_text_dict_value_list = list(
                target_text_dict.values()
            )  # 表头单一对象字典内容列表
            no_space_target_text_dict = dict(
                zip(target_text_dict_keys_list, target_text_dict_value_list)
            )  # 表头单一对象去字典空格字典
            for key in no_space_target_text_dict:
                print('这是目标字段名称（自己填写的）和目标属性相对方向键值对：',key,no_space_target_text_dict[key])
            no_space_target_text_dict_keys_list = list(
                no_space_target_text_dict.keys()
            )  # 表头单一对象键列表(去空格)
            # no_space_target_text_dict_keys_list = [no_space_str.replace(' ','') for no_space_str in target_text_dict_keys_list]#表头单一对象字典键列表（去空格 ）
            # print(target_text_dict_keys_list)
            multi_target_text_dict_keys_list = [
                no_space_str.replace(" ", "").strip()
                for no_space_str in list(list_target_text_dict.keys())
            ]  # 深度、时代等一列多对象字典键列表
            multi_target_text_dict_value_list = list(
                list_target_text_dict.values()
            )  # 深度、时代等一列多对象字典内容列表
            no_space_multi_target_text_dict = dict(
                zip(multi_target_text_dict_keys_list, multi_target_text_dict_value_list)
            )  # 深度、时代等一列多对象去字典空格字典
            no_space_multi_target_text_dict_keys_list = list(
                no_space_multi_target_text_dict
            )  # 深度、时代等一列多对象键列表(去空格)

            #########################################################
            #########################################################
            #########################################################
            current_times = get_current_time()
            print(str(current_times) + "  完成各类目标对象的分类整理...")
            print("")
            print("================================================")
            print("==============开始识别每个柱状图内容==============")
            print("================================================")
            print("")
            current_times = get_current_time()
            print(str(current_times) + "  开始逐一按每个柱状图柱状图的范围提取目标信息...")
            print("")
            # 按range_id区分柱状图，分别识别分层信息表、钻孔信息表等的内容
            # for line in horizon_line_list:
            #     print('all:',line)
            depth_count = 0  # 深度如果为0，就是个空表，跳过
            # 按识别到的柱状图范围开始识别目标内容
            DrillHole_info_list = []
            Separation_info_list = []
            TG_BG_info_list = []
            for id in [id[0] for id in max_frame_list]:
                text_with_target_id_list = [text for text in text_list if text[0] == id]
                max_frame_list_with_target_id_list = [
                    frame for frame in max_frame_list if frame[0] == id
                ]
                vertical_line_list_with_target_id_list = [
                    line for line in vertical_line_list if line[4] == id
                ]
                horizon_line_list_with_target_id_list = [
                    line for line in horizon_line_list if line[4] == id
                ]
                # for line in horizon_line_list_with_target_id_list:
                #     print('with id:',line)
                max_frame_bottom_line_list = [
                    line for line in max_frame_list if line[0] == id
                ]
                # print(max_frame_bottom_line_list)
                max_frame_bottom_line = max_frame_bottom_line_list[
                    0
                ]  # 将外框底边加入横线列表
                # print(max_frame_bottom_line)
                max_frame_bottom_line_minX = max_frame_bottom_line[1]
                max_frame_bottom_line_minY = max_frame_bottom_line[2]
                max_frame_bottom_line_maxX = max_frame_bottom_line[3]
                max_frame_bottom_line_maxY = max_frame_bottom_line[
                    2
                ]  # 底边为水平线，端点y相同

                # for line in horizon_line_list_with_target_id_list:
                #     print(line)
                # print('我还有按时散发出胜多负少的v发')
                # print(str(max_frame_bottom_line_minX)+'   asdasdadacawefdc c')
                horizon_line_list_with_target_id_list = [
                    line
                    for line in horizon_line_list_with_target_id_list
                    if not line[1] == max_frame_bottom_line_minY
                ]  # 删除与边框底边重合的水平直线
                # for line in horizon_line_list_with_target_id_list:
                #     print(line)

                horizon_line_list_with_target_id_list.append(
                    (
                        max_frame_bottom_line_minX,
                        max_frame_bottom_line_minY,
                        max_frame_bottom_line_maxX,
                        max_frame_bottom_line_maxY,
                        id,
                    )
                )
                # for line in horizon_line_list_with_target_id_list:
                #     print(line)
                horizon_polyline_with_target_id_list = [
                    line for line in horizon_polyline_with_range_list if line[4] == id
                ]
                horizon_line_polyline_list = (
                    horizon_line_list_with_target_id_list
                    + horizon_polyline_with_target_id_list
                )  # 水平直线、多段线合并列表
                vertical_polyline_with_target_id_list = [
                    line for line in vertical_polyline_with_range_list if line[4] == id
                ]

                Drilling_information_list = []  # 钻孔信息表
                Stratification_infomation_list = []  # 分层信息表
                print("-------------------------钻孔信息表----------------------------")
                # txt.write('\n')
                # txt.write('========================================================================================'+ '\n')
                # txt.write('========================================================================================'+ '\n')
                txt.write(
                    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
                    + "\n"
                )
                txt.write(
                    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
                    + "\n"
                )
                # txt.write(
                #     '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━'
                #     + '\n')
                for frame in max_frame_list:
                    if frame[0] == id:
                        txt.write(
                            "【柱状图标识】"
                            + str(frame[0])
                            + ' CAD中的定位命令："zoom '
                            + str(frame[1])
                            + ","
                            + str(frame[2])
                            + " "
                            + str(frame[3])
                            + ","
                            + str(frame[4])
                            + ' "(是的，最后有个空格)'
                            + "\n"
                        )
                txt.write(
                    "**************************钻孔信息表**************************" + "\n"
                )
                ##A.钻孔信息表内容识别
                # 1、cad中的文本对象是否在输入的目标字段列表中（字符串精确匹配）
                drill_hole_info_list = []
                match_target_text_keys_list = []#填写的字段名称与cad中的文本对象已匹配上列表
                for text in text_with_target_id_list:
                    text_range_id = text[0]
                    text_content = text[1].strip()
                    # print('这是cad中的一个文字对象，我应该找到与它处于同一格内的其他文字对象连接起来，并加入这个列表：',text_content)
                    text_insert_coordinate = text[2]
                    if text_content.replace(" ", "") in no_space_target_text_dict_keys_list:
                        match_target_text_keys_list.append(text_content)
                        target_value = get_neraby_text(
                            text_content,
                            text_insert_coordinate,
                            text_range_id,
                            max_frame_list_with_target_id_list,
                            vertical_line_list_with_target_id_list,
                            horizon_line_list_with_target_id_list,
                            horizon_polyline_with_target_id_list,
                            vertical_polyline_with_target_id_list,
                            no_space_target_text_dict[text_content.replace(" ", "")],
                            text_with_target_id_list,
                            txt,
                        )
                        if target_value != None:
                            print('cad中的字段和cad中的目标属性：',text_content,target_value)
                            drill_hole_info_list.append(target_value)
                            if "编号" in text_content or "孔号" in text_content.replace(
                                " ", ""
                            ):
                                if not "工程" in text_content.replace(" ", ""):
                                    drilling_number = target_value[3]
                                    key_value_in_range_drilling_dict = (
                                        range_id_drilling_number_chart_dict[text_range_id]
                                    )
                                    to_be_update_dict = {text_range_id: drilling_number}
                                    # if drilling_number != key_value_in_range_drilling_dict:
                                    range_id_drilling_number_chart_dict.update(
                                        to_be_update_dict
                                    )
                print('------------------------------------------------------')
                print('填写的字段名称与cad中的文本对象已匹配上列表:')
                for key in match_target_text_keys_list:
                    print(key)
                # 找出只存在于no_space_target_text_dict_keys_list中的元素(也就是cad中没匹配上的但用户已填写字段名称列表)
                unique_to_no_space_target_text_dict_keys_list = [item for item in no_space_target_text_dict_keys_list if item not in match_target_text_keys_list]
                print('------------------------------------------------------')
                print('填写的字段名称与cad中的文本对象匹配不上列表:')
                for key in unique_to_no_space_target_text_dict_keys_list:
                    print(key)

                # 2、精确匹配不上的目标字段，把它拆开成单个字的列表，用每个字单独与cad中的文本对象做匹配
                value_to_exclude_list = []#存放这一步匹配到的新目标字段，以便后续在未匹配列表中排除
                for key in unique_to_no_space_target_text_dict_keys_list:
                    single_char_list = list(key)#拆开字符串每个字，放入列表
                    print('------------------------------------------------------')
                    for char in single_char_list:
                        print('待匹配单个字：',char)
                    print('------------------------------------------------------')
                    maybe_in_the_same_cell_list = []#潜在的可能是目标字段的单个文本对象列表
                    for text in text_with_target_id_list:
                        text_range_id = text[0]
                        text_content = text[1].strip()
                        # print(text_content)
                        # print('这是cad中的一个文字对象，我应该找到与它处于同一格内的其他文字对象连接起来，并加入这个列表：',text_content)
                        text_insert_coordinate = text[2]
                        if text_content.replace(" ", "") in single_char_list:
                            print('开始对字符串【',key,'】中的【',text_content,'】字进行匹配：')
                            # match_target_text_keys_list.append(text_content)
                            target_value = get_neraby_text(
                                text_content,
                                text_insert_coordinate,
                                text_range_id,
                                max_frame_list_with_target_id_list,
                                vertical_line_list_with_target_id_list,
                                horizon_line_list_with_target_id_list,
                                horizon_polyline_with_target_id_list,
                                vertical_polyline_with_target_id_list,
                                no_space_target_text_dict[key.replace(" ", "")],
                                text_with_target_id_list,
                                txt,
                            )
                            if target_value != None:
                                print('上一步未匹配的字段经过单个字拆分后，匹配到的cad中的字段和cad中的目标属性：',char,target_value,'目标字段是：',key)
                                maybe_in_the_same_cell_list.append(target_value)

                    print('------------------------------------------------------')
                    print('目标字段【',key,'】拆分单字后匹配到的文本对象列表：')
                    for to_concatenate_char in maybe_in_the_same_cell_list:
                        print(to_concatenate_char)

                    print('------------------------------------------------------')
                    print('目标字段【',key,'】拆分单字后匹配到的潜在待合并字段文本对象列表：')
                    # 统计每个从第三个元素开始的子元组出现的次数
                    counts = Counter(t[2:] for t in maybe_in_the_same_cell_list)

                    # 筛选出从第三个元素开始相同且出现次数大于 1 的元组
                    result = [t for t in maybe_in_the_same_cell_list if counts[t[2:]] > 1]
                    print(result)

                    result_char_list = [char[1] for char in result] 
                    print('cad中潜在待合并字段单个文本对象列表',result_char_list)
                    key_list = list(key)#输入的目标字段拆分单字列表
                    print('输入的目标字段单个文字列表：',key_list)
                    is_the_same = is_list_equal_ignore_order(result_char_list,key_list)
                    print('是目标字段吗：',is_the_same)#不考虑列表中元素的位置排序，仅依据内容来检查两个上面列表是否相等
                    if is_the_same == True:
                        value_to_exclude_list.append(key)#将新匹配到的目标字段加入待排除列表
                        new_target_tuple = (result[0][0], key) + result[0][2:]
                        print('待合并到结果元组列表的目标元组：',new_target_tuple)
                        target_value = new_target_tuple
                        drill_hole_info_list.append(target_value)
                        if "编号" in text_content or "孔号" in text_content.replace(
                            " ", ""
                        ):
                            if not "工程" in text_content.replace(" ", ""):
                                drilling_number = target_value[3]
                                key_value_in_range_drilling_dict = (
                                    range_id_drilling_number_chart_dict[text_range_id]
                                )
                                to_be_update_dict = {text_range_id: drilling_number}
                                # if drilling_number != key_value_in_range_drilling_dict:
                                range_id_drilling_number_chart_dict.update(
                                    to_be_update_dict
                                )

                print('------------------------------------------------------')
                print('通过对输入的目标字段拆字匹配到的新字段列表：',value_to_exclude_list)

                # 3、还是匹配不上的目标字段（key），进行局部匹配，看它是否完全包含于cad中的某个文本对象
                print('------------------------------------------------------')
                print('排除新字段前填写的字段名称与cad中的文本对象匹配不上列表:')
                for key in unique_to_no_space_target_text_dict_keys_list:
                    print(key)
                print('------------------------------------------------------')
                print('排除新字段后填写的字段名称与cad中的文本对象匹配不上列表:')
                unique_to_no_space_target_text_dict_keys_list = [i for i in unique_to_no_space_target_text_dict_keys_list if i not in value_to_exclude_list]
                for key in unique_to_no_space_target_text_dict_keys_list:
                    print(key)      
                    for text in text_with_target_id_list:
                        text_range_id = text[0]
                        text_content = text[1].strip()
                        # print(text_content)
                        # print('这是cad中的一个文字对象，我应该找到与它处于同一格内的其他文字对象连接起来，并加入这个列表：',text_content)
                        text_insert_coordinate = text[2]
                        if key in text_content.replace(" ", ""):
                            print('开始对字符串【',key,'】进行局部匹配：')
                            # match_target_text_keys_list.append(text_content)
                            target_value = get_neraby_text(
                                text_content,
                                text_insert_coordinate,
                                text_range_id,
                                max_frame_list_with_target_id_list,
                                vertical_line_list_with_target_id_list,
                                horizon_line_list_with_target_id_list,
                                horizon_polyline_with_target_id_list,
                                vertical_polyline_with_target_id_list,
                                no_space_target_text_dict[key.replace(" ", "")],
                                text_with_target_id_list,
                                txt,
                            )
                            if target_value != None:
                                print('------------------------------------------------------')
                                print('目标字段与cad中文本对象进行局部匹配后，匹配到的cad中的字段和cad中的目标属性：',text_content,target_value,'目标字段是：',key)
                                new_target_tuple = (target_value[0], key) + target_value[2:]
                                print('------------------------------------------------------')
                                print('待合并到结果元组列表的目标元组：',new_target_tuple)
                                target_value = new_target_tuple
                                drill_hole_info_list.append(target_value)
                                # maybe_in_the_same_cell_list.append(target_value)



                # exit()


                for key, value in range_id_drilling_number_chart_dict.items():
                    print(key, value)#柱状图标识和钻孔编号
                    # else:
                    #     continue
                    # print(text_range_id,text_content,text_insert_coordinate,target_value)

                for item in drill_hole_info_list:
                    range_id = item[0]
                    field_name = item[1]
                    navigation = item[2]
                    content = item[3]
                    affiliate_str = item[4]
                    drilling_number = range_id_drilling_number_chart_dict[range_id]
                    if content == "格子内为【空】":
                        content = "【空】(1、是真的空 2、字段名填重复了 3、不知道)"
                    print(
                        "【柱状图标识】"
                        + str(range_id)
                        + " "
                        + "【钻孔编号】"
                        + str(drilling_number)
                        + " "
                        + "【字段名称】"
                        + field_name
                        + " "
                        + " 【"
                        + navigation
                        + "边格子的内容】"
                        + str(content)
                        + " "
                        + str(affiliate_str)
                    )
                    txt.write(
                        "【柱状图标识】"
                        + str(range_id)
                        + " "
                        + "【钻孔编号】"
                        + str(drilling_number)
                        + " "
                        + "【字段名称】"
                        + field_name
                        + " 【"
                        + navigation
                        + "边格子的内容】"
                        + str(content)
                        + " "
                        + str(affiliate_str)
                        + "\n"
                    )
                # exit()
                DrillHole_info_list = (
                    DrillHole_info_list + drill_hole_info_list
                )  # 合并到总表
                # exit()
                # for drilling_number in drilling_number_list:
                #     print(drilling_number)

                print("-------------------------分层信息表----------------------------")
                txt.write(
                    "**************************分层信息表**************************" + "\n"
                )
                ##B.分层信息、土工样、标贯及其他纵向内容识别
                TG_BG_range_list = []  # 标贯土工等字段下方范围列
                target_key_range_list = []  # 目标字段，下方目标内容范围
                for key_string in no_space_multi_target_text_dict_keys_list:
                    # print('来啊干我啊','key_string: ',key_string)
                    target_part_key_text_list = []
                    for text in text_with_target_id_list:
                        text_range_id = text[0]
                        text_content = text[1]
                        # print('All text:',text_content)
                        text_insert_coordinate = text[2]
                        if (
                            if_text_content_part_in_string(
                                text_content.replace(" ", ""), key_string
                            )
                            == True
                        ):  # and ' ' not in text_content:
                            target_part_key_text_list.append(text)
                            print('这是隐藏字符吗：',text)
                    key_cell_list = []
                    # for line in horizon_line_list_with_target_id_list:
                    #     print(line)
                    # exit()
                    for part_key_text in target_part_key_text_list:
                        # print('这里还有【Part_key】  ',part_key_text)
                        if " " not in part_key_text and part_key_text[1] != '':
                            same_cell_full_str = get_partner_in_the_same_cell(
                                part_key_text,
                                text_with_target_id_list,
                                max_frame_list_with_target_id_list,
                                vertical_line_list_with_target_id_list,
                                vertical_polyline_with_target_id_list,
                                horizon_line_list_with_target_id_list,
                                horizon_polyline_with_target_id_list,
                            )
                            same_cell_full_field_name = same_cell_full_str[
                                0
                            ]  # 同一格子内所有字符串相连
                            same_cell_full_field_range = same_cell_full_str[
                                1
                            ]  # 该格子(minX,maxX,minY,maxY)
                            print(same_cell_full_field_name, same_cell_full_field_range)
                        else:
                            continue
                        # exit()
                        same_cell_full_field_name = same_cell_full_field_name.strip()
                        print(part_key_text, '同格内容：', same_cell_full_field_name,
                            '所在格子坐标范围：', same_cell_full_field_range)
                        if "=" not in key_string:
                            # print('yeess')
                            # print('我的key呢？？？？？？这里还是对的',key_string)
                            if if_text_content_part_in_string(
                                key_string, same_cell_full_field_name.replace(" ", "")
                            ):
                                print("你这有换行符吗？", same_cell_full_field_name)
                                # print('fuckkkkkkkk!!!!这里就不行了')
                                key_cell_list.append(
                                    (
                                        text_range_id,
                                        same_cell_full_field_name,
                                        same_cell_full_field_range,
                                        part_key_text[1],
                                        part_key_text[2][0],
                                        part_key_text[2][1],
                                    )
                                )  # part_key_text是key的部分字符串
                        else:
                            first_char_key_string_with_no_space = list(
                                key_string.replace(" ", "").replace("=", "")
                            )[0]
                            first_char_same_cell_full_field_name_with_no_space = list(
                                same_cell_full_field_name.replace(" ", "")
                            )[0]
                            # print(first_char_key_string_with_no_space,first_char_same_cell_full_field_name_with_no_space)
                            if (
                                first_char_key_string_with_no_space
                                == first_char_same_cell_full_field_name_with_no_space
                            ):
                                # print('就是他了',same_cell_full_field_name)
                                # if if_text_content_part_in_string(
                                #         key_string, same_cell_full_field_name.replace(' ', '')):
                                key_cell_list.append(
                                    (
                                        text_range_id,
                                        same_cell_full_field_name,
                                        same_cell_full_field_range,
                                        "=" + part_key_text[1],
                                        part_key_text[2][0],
                                        part_key_text[2][1],
                                    )
                                )  # part_key_text是key的部分字符串
                            else:
                                # print('他不行',same_cell_full_field_name)
                                pass

                            # print(key_string)
                            # exit()
                            # part_key_text_first_char = part_key_text
                    # exit()

                    key_cell_list = sorted(
                        key_cell_list, key=lambda x: [-x[5]]
                    )  # 字段名分块按y降序

                    # print(key_cell_list)
                    key_name = "".join([key_part[3] for key_part in key_cell_list]).replace(
                        " ", ""
                    )  # 拼接出cad中的对应字段名
                    # print('尴尬',key_name)
                    key_cell_list = [
                        (item[0], item[1], item[2]) for item in key_cell_list
                    ]  # 删除列表中的part_key_text[1]
                    key_cell_list = list(
                        set(key_cell_list)
                    )  # 删除重复项(同一个字段在cad中可能分成几个文本对象，例如'层底','深度','m'实为一个字段名，位于同一个格子中)

                    for item in key_cell_list:
                        # print('这里就开始没了！！！！！！！',item)
                        range_id = item[0]  # 柱状图标识
                        field_name_in_cad = item[1]  # cad中的对应字段名
                        cell_coordinates = item[
                            2
                        ]  # (minX,maxX,minY,maxY) cad中的对应字段名所在格子坐标范围
                        min_x = cell_coordinates[0]
                        max_x = cell_coordinates[1]
                        min_y = max_frame_list_with_target_id_list[0][2]  # 外框底边y
                        max_y = cell_coordinates[2]  # 字段所在格子底边y
                        # for i in no_space_multi_target_text_dict:
                        #     print(i)
                        key_repr_field_name = no_space_multi_target_text_dict[
                            key_name
                        ]  # 从字典中取回对应目标字段名称(CAD字段：'层底深度(m)'->目标字段名：'层底深度')

                        if key_repr_field_name in [
                            "层底深度",
                            "剖面层号",
                            "时代成因",
                            "岩土描述",
                        ]:
                            target_key_range_list.append(
                                (
                                    range_id,
                                    key_repr_field_name,
                                    field_name_in_cad,
                                    (min_x, max_x, min_y, max_y),
                                )
                            )  # 分层信息表几个字段下方格子范围
                            print(
                                "分层信息表字段检查："
                                + str(range_id)
                                + " "
                                + str(key_repr_field_name)
                                + " "
                                + str(field_name_in_cad)
                                + " "
                                + str(min_x)
                                + " "
                                + str(max_x)
                                + " "
                                + str(min_y)
                                + " "
                                + str(max_y)
                            )
                        else:
                            TG_BG_range_list.append(
                                (
                                    range_id,
                                    key_repr_field_name,
                                    field_name_in_cad,
                                    (min_x, max_x, min_y, max_y),
                                )
                            )  # 标贯、土工字段等下方格子范围
                for TG_BG in TG_BG_range_list:
                    print(
                        "土工、标贯信息表字段检查："
                        + str(TG_BG[0])
                        + " "
                        + str(TG_BG[1])
                        + " "
                        + str(TG_BG[2])
                        + " "
                        + str(TG_BG[3][0])
                        + " "
                        + str(TG_BG[3][1])
                        + " "
                        + str(TG_BG[3][2])
                        + " "
                        + str(TG_BG[3][3])
                    )

                # B1.分层信息表识别
                YTMS_bottom_line_list = []  # 岩土描述列下方横线列表
                for field in target_key_range_list:
                    range_coor = field[3]  # 外框范围(minX,maxX,minY,maxY)
                    # print('What ',range_coor)
                    if field[1] == "时代成因":
                        range_id = field[0]
                        time_reason_range_id = range_id
                        key_repr_field_name = field[1]
                        field_name_in_cad = field[2]
                        # range_coor = field[3]
                        time_reason_range = range_coor  # (最小x,最大x,最小y,最大y)
                        print(
                            "这什么意思", range_id, key_repr_field_name, time_reason_range
                        )

                    if field[1] == "剖面层号":
                        range_id = field[0]
                        time_reason_range_id = range_id
                        key_repr_field_name = field[1]
                        field_name_in_cad = field[2]
                        # range_coor = field[3]
                        PM_range = range_coor  # (最小x,最大x,最小y,最大y)
                        # print('PM_range',PM_range)
                    if field[1] == "岩土描述":
                        range_id = field[0]
                        time_reason_range_id = range_id
                        key_repr_field_name = field[1]
                        field_name_in_cad = field[2]
                        # range_coor = field[3]
                        YTMS_range = range_coor  # (最小x,最大x,最小y,最大y)
                        mid_field_x = (range_coor[0] + range_coor[1]) / 2
                        for line in horizon_line_polyline_list:
                            line_start_point_x = line[0]
                            line_start_point_y = line[1]
                            line_end_point_x = line[2]
                            line_end_point_y = line[3]
                            line_range_id = line[4]
                            line_max_x = max(line_start_point_x, line_end_point_x)
                            line_min_x = min(line_start_point_x, line_end_point_x)
                            line_max_y = max(line_start_point_y, line_end_point_y)
                            line_min_y = min(line_start_point_y, line_end_point_y)
                            if (
                                mid_field_x > line_min_x
                                and mid_field_x < line_max_x
                                and range_coor[3] > line_min_y
                            ):
                                YTMS_bottom_line_list.append(line)
                                print(line)
                # exit()

                unique_y_YTMS_bottom_line_list = sorted(
                    list(set([line[1] for line in YTMS_bottom_line_list])), reverse=True
                )  # 岩土描述用底下横线来分层
                # for line in unique_y_YTMS_bottom_line_list:
                #     print(line)
                # exit()
                ytms_count = len(unique_y_YTMS_bottom_line_list)  # 以横线统计多少个岩土描述
                # print('you jige ',ytms_count)
                # 从层底深度开始，读取分层信息表的时代成因、剖面层号、岩土描述等对应内容
                for field in target_key_range_list:
                    # print(time_reason_range[0],time_reason_range[1])
                    if field[1] == "层底深度":  # 从层底深度开始搜索
                        range_id = field[0]
                        key_repr_field_name = field[1]
                        field_name_in_cad = field[2]
                        range_coor = field[3]
                        # print(range_id, '目标字段：', key_repr_field_name, 'CAD中匹配字段名：',
                        #             field_name_in_cad,'下方格子坐标范围：',range_coor)
                        bottom_depth_list = []  # 每个柱状图层底深度值列表
                        for (
                            text
                        ) in (
                            text_with_target_id_list
                        ):  # (text_range_id, text_content, text_insert_coordinate)
                            text_content = text[1]
                            text_insertion_point_x = text[2][0]  # 插入点x
                            text_insertion_point_y = text[2][1]  # 插入点y
                            if coor_inside_range(
                                text_insertion_point_x,
                                text_insertion_point_y,
                                range_coor[0],
                                range_coor[1],
                                range_coor[2],
                                range_coor[3],
                            ):
                                bottom_depth_list.append(
                                    (
                                        range_id,
                                        key_repr_field_name,
                                        field_name_in_cad,
                                        text_content,
                                        text_insertion_point_x,
                                        text_insertion_point_y,
                                    )
                                )
                                # print('柱状图标识：',range_id, '目标字段：', key_repr_field_name, 'CAD中匹配字段名：',
                                #     field_name_in_cad,'深度值：',text_content,'y坐标：',text_insertion_point_y)
                        bottom_unique_depth_list = []
                        exist_depth_list = []  # 深度值唯一值列表
                        for depth in bottom_depth_list:
                            # print(depth)
                            # exit()
                            range_id = depth[0]
                            key_repr_field_name = depth[1]
                            field_name_in_cad = depth[2]
                            try:
                                text_content = float(depth[3])  # 深度值
                            except:
                                continue
                            # print(text_content)
                            text_insertion_point_x = depth[4]
                            text_insertion_point_y = depth[5]
                            if text_content not in exist_depth_list:
                                exist_depth_list.append(text_content)
                                bottom_unique_depth_list.append(
                                    (
                                        range_id,
                                        key_repr_field_name,
                                        field_name_in_cad,
                                        text_content,
                                        text_insertion_point_x,
                                        text_insertion_point_y,
                                    )
                                )
                            for item in bottom_unique_depth_list:
                                print(item)
                        # exit()
                        unique_bottom_depth_list = list(
                            set(bottom_unique_depth_list)
                        )  # 去个重

                        sort_by_y_bottom_depth_list = sorted(
                            unique_bottom_depth_list, key=lambda x: (-x[5])
                        )  # 按坐标y降序
                        # for i in sort_by_y_bottom_depth_list:
                        #     print(i)
                        # exit()
                        depth_list = []
                        depth_order_number = 1  # 层序号
                        for depth in sort_by_y_bottom_depth_list:
                            range_id = depth[0]
                            key_repr_field_name = depth[1]
                            field_name_in_cad = depth[2]
                            text_content = depth[3]
                            text_insertion_point_x = depth[4]
                            text_insertion_point_y = depth[5]
                            nearest_bottom_line = get_text_nearest_one_line(
                                text_insertion_point_x,
                                text_insertion_point_y,
                                horizon_line_list_with_target_id_list
                                + horizon_polyline_with_target_id_list,
                                "下",
                            )  # 找出下方第一条横线
                            print(
                                "看这里",
                                key_repr_field_name,
                                text_content,
                                nearest_bottom_line,
                            )
                            depth_list.append(
                                (
                                    range_id,
                                    key_repr_field_name,
                                    field_name_in_cad,
                                    text_content,
                                    text_insertion_point_x,
                                    text_insertion_point_y,
                                    nearest_bottom_line,
                                    depth_order_number,
                                )
                            )
                            depth_order_number += 1
                        depth_count = len(depth_list)  # 多少个深度
                        if depth_count == 0:
                            print(
                                "柱状图标识", range_id, "没有搜索到层底深度，这应该是个空表"
                            )
                            txt.write(
                                "【柱状图标识】"
                                + str(range_id)
                                + " 没有搜索到层底深度，这应该是个空表"
                            )
                            continue
                        # print(depth_count)
                        if_last_combine = (
                            0  # 当层底深度和岩土描述对不上时，把最后剩余的层合并到最后一层
                        )
                        if depth_count == ytms_count:
                            print(
                                "【柱状图标识】",
                                range_id,
                                "不错，层底深度和岩土描述数量刚好对应上了，共有：",
                                depth_count,
                                "层",
                            )
                        elif depth_count < ytms_count:
                            print(
                                "【柱状图标识】",
                                range_id,
                                " 层底深度数量 < 岩土描述数量，分别有",
                                depth_count,
                                "和",
                                ytms_count,
                                "层，如有多出的文本对象将合并到第"
                                + str(depth_count)
                                + "层中，可以自行去查看，不过问题应该不大，\n有可能是因为最下面一层除了边框底边外还有条横线，而预设的情况是底部的一层下方除了边框的底边不再另外添加横线分隔。",
                            )
                        elif depth_count > ytms_count:

                            # txt_add.write(str(range_id)+str(drilling_number)+str(depth_count)+str(ytms_count)+str(depth_count==ytms_count)+ "\n")

                            print(
                                "【柱状图标识】",
                                range_id,
                                " 层底深度数量 > 岩土描述数量，分别有",
                                depth_count,
                                "和",
                                ytms_count,
                                "层，出现这种情况可以尝试两种检查方法："
                                + "\n    ①岩土描述太拥挤了，后面的层识别不出来，请自行修改。"
                                + "\n    ②出现了超出外框的线（可能后期被人为修改过，不小心移动了），可尝试适当增加两侧宽度和底部高度"
                                + "\n    ③有些地层没有填写岩土描述，这种情况可能存在各层信息错位，注意手动检查",
                            )
                        txt_add = open(addition_txt_4_depth_description_path, "a")
                        for frame in max_frame_list:
                            if frame[0] == range_id:
                                txt_add.write(
                                    "钻孔编号："
                                    + str(drilling_number)
                                    + " "
                                    + "层底深度数量 > 岩土描述数量？："
                                    + str(depth_count > ytms_count)
                                    + " "
                                    + "层底深度数量："
                                    + str(depth_count)
                                    + " "
                                    + "岩土描述数量："
                                    + str(ytms_count)
                                    + " "
                                    + ' CAD中的定位命令："zoom '
                                    + str(frame[1])
                                    + ","
                                    + str(frame[2])
                                    + " "
                                    + str(frame[3])
                                    + ","
                                    + str(frame[4])
                                    + ' "(是的，最后有个空格)'
                                    + "\n"
                                )
                        txt_add.close()
                        print("-----------------------------------------------------------")
                        # exit()
                        YT_maybe_wrong_list = []  # 待调整顺序的岩土描述列表
                        YT_MText_list = []  # 多行文本列表
                        target_value_list = []
                        i = 0

                        # time_reason_range
                        # 时代成因下方文本对象列表(绑定bounding_box有交集的对象)
                        if use_insertion_point != 1:
                            time_reason_list = [
                                time_reason
                                for time_reason in text_list
                                if coor_inside_range(
                                    time_reason[2][0],
                                    time_reason[2][1],
                                    time_reason_range[0],
                                    time_reason_range[1],
                                    time_reason_range[2],
                                    time_reason_range[3],
                                )
                            ]
                            time_reason_concatenate_text_list = concatenate_text_in_list(
                                time_reason_list
                            )
                        else:
                            time_reason_concatenate_text_list = text_with_target_id_list
                        # for k in time_reason_concatenate_text_list:
                        #     print(k)
                        # exit()
                        # 剖面层号下方文本对象列表(绑定bounding_box有交集的对象)
                        if use_insertion_point != 1:
                            PM_list = [
                                PM_order
                                for PM_order in text_list
                                if coor_inside_range(
                                    PM_order[2][0],
                                    PM_order[2][1],
                                    PM_range[0],
                                    PM_range[1],
                                    PM_range[2],
                                    PM_range[3],
                                )
                            ]
                            # print('PM: ', PM_list)
                            PM_concatenate_text_list = concatenate_text_in_list(PM_list)
                        else:
                            PM_concatenate_text_list = text_with_target_id_list
                        # for k in PM_concatenate_text_list:
                        #     print(k)
                        # exit()
                        # print('有几层啊：',len(depth_list))
                        for depth in depth_list:
                            range_id = depth[0]  # 【柱状图标识】
                            key_repr_field_name = depth[1]  # 【目标字段】
                            field_name_in_cad = depth[2]  # 【目标字段】在CAD中的完整名称
                            text_content = depth[3]  # 深度值
                            text_insertion_point_x = depth[4]  # 深度文本对象插入点x
                            text_insertion_point_y = depth[5]  # 深度文本对象插入点y
                            nearest_bottom_line = depth[
                                6
                            ]  # 深度文本对象下方最近的一条横线对象
                            depth_order_number = depth[7]  # 层数计数
                            # 时代成因列
                            frame_bottom_y = time_reason_range[2]  # 外框底边y
                            nearest_SD_bottom_line_up_line = get_hor_line_nearest_up_text(
                                nearest_bottom_line[1],
                                time_reason_range[0],
                                time_reason_range[1],
                                horizon_line_list_with_target_id_list
                                + horizon_polyline_with_target_id_list,
                                time_reason_concatenate_text_list,
                                0,
                                frame_bottom_y,
                            )  # 最后一个参数：!=1：先x升序再y降序
                            # print('SD：'+nearest_SD_bottom_line_up_line[4])
                            # 剖面层号列
                            frame_bottom_y = PM_range[2]  # 外框底边y
                            nearest_PM_bottom_line_up_line = get_hor_line_nearest_up_text(
                                nearest_bottom_line[1],
                                PM_range[0],
                                PM_range[1],
                                horizon_line_list_with_target_id_list
                                + horizon_polyline_with_target_id_list,
                                PM_concatenate_text_list,
                                0,
                                frame_bottom_y,
                            )  # 最后一个参数：!=1：先x升序再y降序

                            if i < ytms_count:
                                ytms_bottom_y = unique_y_YTMS_bottom_line_list[
                                    i
                                ]  # 隐患，深度和岩土描述数量可能对不上
                                # print()
                                i = i + 1
                            else:
                                i = i + 1
                                print(
                                    "【层底深度数量(按文本对象识别) > 岩土描述数量(按水平线识别)】\n第"
                                    + str(i)
                                    + "层及之后的层未能识别，且后面几层可能会对应错位，请自行修改\n原因可能是:\n    ①岩土描述列中的分隔线[部分]用多段线绘制-->分解多段线\n    ②有线发生偏移了（出现“悬挂节点”和一小截超出外框的情况）-->添加两侧宽度和底部高度增加值\n    ③有些地层没有填写岩土描述，这种情况可能存在各层信息错位，注意手动检查（直接使用开头定位命令'zoom xxx'可缩放到目标范围）"
                                )
                                print(
                                    "-----------------------------------------------------------"
                                )
                                txt.write(
                                    "【层底深度数量(按文本对象识别) > 岩土描述数量(按水平线识别)】\n第"
                                    + str(i)
                                    + "层及之后的层未能识别，且后面几层可能会对应错位，请自行修改\n原因可能是:\n    ①岩土描述列中的分隔线[部分]用多段线绘制-->分解多段线\n    ②有线发生偏移了（出现“悬挂节点”和一小截超出外框的情况）-->添加宽度和底部高度增加值\n    ③有些地层没有填写岩土描述，这种情况可能存在各层信息错位，注意手动检查（直接使用开头定位命令'zoom xxx'可缩放到目标范围）"
                                    + "\n"
                                )
                                txt.write(
                                    "-----------------------------------------------------------"
                                    + "\n"
                                )
                                continue
                            # print('层序：',depth_order_number,'底边:',ytms_bottom_y,'外框底边:',YTMS_range[2])

                            frame_bottom_y = YTMS_range[2]  # 外框底边y
                            if depth_order_number == len(depth_list):
                                frame_bottom_y = "bottom:" + str(frame_bottom_y)
                            nearest_YTMS_bottom_line_up_line = get_hor_line_nearest_up_text(
                                ytms_bottom_y,
                                YTMS_range[0],
                                YTMS_range[1],
                                horizon_line_list_with_target_id_list
                                + horizon_polyline_with_target_id_list,
                                text_with_target_id_list,
                                1,
                                frame_bottom_y,
                            )  # 最后一个参数：1：先y降序再x升序

                            # print('[', nearest_YTMS_bottom_line_up_line[4], ']')
                            # if nearest_YTMS_bottom_line_up_line[4] != '【空】':

                            if (
                                "@" in nearest_YTMS_bottom_line_up_line[4]
                            ):  # 岩土描述由单行文本组成
                                for YT_text in nearest_YTMS_bottom_line_up_line[4].split(
                                    "@"
                                ):

                                    YT_maybe_wrong_list.append(
                                        (range_id, depth_order_number, YT_text)
                                    )
                            else:  # 岩土描述由多行文本组成
                                # print('why here',('@' in nearest_YTMS_bottom_line_up_line[4]))
                                # print(nearest_YTMS_bottom_line_up_line[4])
                                for YT_text in nearest_YTMS_bottom_line_up_line:
                                    if "$" in str(YT_text):
                                        YT_MText_list.append(
                                            (
                                                depth_order_number,
                                                YT_text[
                                                    YT_text.index(
                                                        " ", YT_text.index(" ") + 3
                                                    )
                                                    + 3 :
                                                ],
                                            )
                                        )  # 多行文本列，后面为截取前面非描述内容
                                        # print(YT_text)
                            # else:
                            #     continue
                            target_value_list.append(
                                (
                                    range_id,
                                    depth_order_number,
                                    text_content,
                                    nearest_SD_bottom_line_up_line[4],
                                    nearest_PM_bottom_line_up_line[4],
                                )
                            )
                        # for value in YT_maybe_wrong_list:
                        #     print('岩土描述：',value)
                        # exit()
                        # for value in YT_text:
                        #     print('mT ',value)
                        # print(len(YT_maybe_wrong_list))
                        # exit()
                        # print(len(YT_MText_list))

                        if len(YT_MText_list) == 0:
                            print("【CAD文件中的岩土描述用文本(单行)表达】")
                            txt.write("【CAD文件中的岩土描述用文本(单行)表达】" + "\n")
                            print("")
                            obj_id_YT_maybe_wrong_list = []
                            used_object_id_list = []
                            for item in YT_maybe_wrong_list:
                                # print('重复？？？？？？？？？？？？？？？？？？？',item)
                                range_id = item[0]
                                depth_order_number = item[1]
                                text_object_id = int(item[2].split(" $ ")[1])
                                if text_object_id not in used_object_id_list:
                                    used_object_id_list.append(text_object_id)

                                    text_content = item[2].split(" $ ")[2]
                                    # print(text_content)
                                    obj_id_YT_maybe_wrong_list.append(
                                        (
                                            range_id,
                                            depth_order_number,
                                            text_object_id,
                                            text_content,
                                        )
                                    )
                                # print(range_id,depth_order_number,text_object_id,text_content)
                            # for item in obj_id_YT_maybe_wrong_list:
                            # print('重复？？？？？？',item)

                            sort_obj_id_YT_maybe_wrong_list = obj_id_YT_maybe_wrong_list  # 默认不排序（描述有重叠才考虑用下面的排序）

                            # for obj in sort_obj_id_YT_maybe_wrong_list:
                            #     print(obj)
                            # exit()

                            tips = "(目前有三种排序方式：0,1,999 如果拼接效果不好可以试试另外两种)"

                            # 下面的排序暂时不按坐标判断升降序(要改太麻烦了，前面的判断中坐标没传下来)
                            if YTMS_sort_type == 0:
                                print(
                                    "【岩土描述】排序方式: 0 :先按层序升序再按objectid升序"
                                    + tips
                                )
                                txt.write(
                                    "【岩土描述】排序方式: 0 :先按层序升序再按objectid升序"
                                    + tips
                                    + "\n"
                                )
                                print("")
                                sort_obj_id_YT_maybe_wrong_list = sorted(
                                    obj_id_YT_maybe_wrong_list, key=lambda x: (x[1], x[2])
                                )  # 按先text_object_id先升序再depth_order_number降序排列(岩土描述列)
                            elif YTMS_sort_type == 1:
                                print(
                                    "【岩土描述】排序方式: 1 :先objectid降序再按层序降序"
                                    + tips
                                )
                                txt.write(
                                    "【岩土描述】排序方式: 1 :先objectid降序再按层序降序"
                                    + tips
                                    + "\n"
                                )
                                print("")
                                sort_obj_id_YT_maybe_wrong_list = sorted(
                                    obj_id_YT_maybe_wrong_list, key=lambda x: (-x[2], x[1])
                                )  # 按先text_object_id先降序再depth_order_number降序排列(岩土描述列)
                            else:
                                print(
                                    "【岩土描述】排序方式: != 0 and ! = 1 :按objectid升序(原始排序)"
                                    + tips
                                )
                                txt.write(
                                    "【岩土描述】排序方式: != 0 and ! = 1 :按objectid升序(原始排序)"
                                    + tips
                                    + "\n"
                                )
                                print("")

                            # print(len(sort_obj_id_YT_maybe_wrong_list))
                            list_with_front_diff = []  # 与上一个元素objectid的差
                            list_with_next_diff = []  # 与下一个元素objectid的差
                            for i, j in zip(
                                sort_obj_id_YT_maybe_wrong_list[0::],
                                sort_obj_id_YT_maybe_wrong_list[1::],
                            ):
                                diff_width_with_front_id = abs(i[2] - j[2])
                                diff_width_with_next_id = abs(j[2] - i[2])
                                depth_order_number = i[1]
                                # print(depth_order_number,diff_width_with_next_id)
                                list_with_front_diff.append(
                                    (depth_order_number, diff_width_with_next_id)
                                )
                                list_with_next_diff.append(
                                    (depth_order_number, diff_width_with_next_id)
                                )

                            # xxzx = [order[1] for order in sort_obj_id_YT_maybe_wrong_list]
                            # aass = max(list(set(xxzx)))
                            # print(aass)

                            # print('岩土描述分层数：'+str(len(sort_obj_id_YT_maybe_wrong_list)))
                            # for order in sort_obj_id_YT_maybe_wrong_list:
                            #     print('分层描述：'+order)
                            if len(sort_obj_id_YT_maybe_wrong_list) != 0:

                                try:
                                    max_depth_order_number = max(
                                        list(
                                            set(
                                                [
                                                    order[1]
                                                    for order in sort_obj_id_YT_maybe_wrong_list
                                                ]
                                            )
                                        )
                                    )  # 层序号最大值
                                except:
                                    print(
                                        "出错了，很可能是【岩土描述】(CAD里的)字段名没填对或者【岩土描述】为空"
                                    )
                                    return

                                list_with_front_diff.insert(0, (1, 99999999))
                                list_with_next_diff.append(
                                    (max_depth_order_number, 99999999)
                                )
                                sorting_ytms_list = []
                                for obj1, obj2, content in zip(
                                    list_with_front_diff,
                                    list_with_next_diff,
                                    sort_obj_id_YT_maybe_wrong_list,
                                ):
                                    # print('有重复？？？？',content)
                                    sorting_ytms_list.append(
                                        (obj2[0], content[3])
                                    )  # (分层号，内容)
                                    # print('深度分层号：', obj2[0], '内容：', content[3], '前：后 obj_id：',
                                    #       obj1[1], ':', obj2[1])
                                # 排个序试试
                                max_ytms_order_number = max(
                                    [order[0] for order in sorting_ytms_list]
                                )  # 分层号起点
                                order_list = list(
                                    set([order[0] for order in sorting_ytms_list])
                                )  # 岩土描述序号列表
                                used_ytms_order_number_list = []
                                sorted_ytms_list = []
                                current_order = 0
                                # print(depth_count)
                                for ytms in sorting_ytms_list:
                                    # print('有重复的？',ytms)
                                    ytms_order_number = ytms[0]  # 当前岩土描述所属分层号
                                    ytms_content = ytms[1]  # 岩土描述文本对象内容
                                    # print(ytms_order_number,ytms_content)
                                    if (
                                        ytms_order_number not in used_ytms_order_number_list
                                        and ytms_order_number != current_order
                                    ):
                                        used_ytms_order_number_list.append(
                                            ytms_order_number
                                        )
                                        # if current_order < depth_count:
                                        current_order = ytms_order_number
                                    else:
                                        ytms_order_number = current_order
                                    sorted_ytms_list.append(
                                        (ytms_order_number, ytms_content)
                                    )
                                ytms_list = []
                                for order in order_list:
                                    match_order_string_list = [
                                        ytms[1]
                                        for ytms in sorted_ytms_list
                                        if ytms[0] == order
                                    ]
                                    current_order_text_combine = (
                                        "".join(match_order_string_list)
                                        .replace(" ", "")
                                        .replace("%%%", "%")
                                    )
                                    ytms_list.append(
                                        (order, current_order_text_combine)
                                    )  # 层号，连接好的岩土描述
                            else:
                                print("【CAD文件中的岩土描述用多行文本表达】")
                                txt.write("【CAD文件中的岩土描述用多行文本表达】" + "\n")
                                print("")
                                ytms_list = (
                                    YT_MText_list  # 多行文本的岩土描述列表(层号，描述内容)
                                )
                            ####################################################################################
                            ####################################################################################
                            ####################################################################################
                            # 对重叠造成的错层文本重新拼接
                            str_update_list = []
                            new_ytms_list = []
                            for info in ytms_list:
                                order_number = info[0]
                                content = info[1]
                                last_character = content[-1]
                                print(
                                    "层序："
                                    + str(order_number)
                                    + " 本层内容："
                                    + content
                                    + " 本层最后一个字符："
                                    + last_character
                                )

                                # exit()
                                first_en_colon_position = findSubStrIndex(":", content, 1)
                                first_chn_colon_position = findSubStrIndex("：", content, 1)

                                first_full_stop_position = findSubStrIndex("。", content, 1)
                                print(first_full_stop_position)
                                # print(first_en_colon_position,first_chn_colon_position,first_full_stop_position)

                                tips = ""
                                if first_chn_colon_position == None:
                                    first_chn_colon_position = 0
                                if first_en_colon_position == None:
                                    first_en_colon_position = 0
                                if first_full_stop_position == None:
                                    first_full_stop_position = len(content)
                                first_colon_position = max(
                                    first_chn_colon_position, first_en_colon_position
                                )
                                print("第一个冒号的位置：", first_colon_position)
                                print("第一个句号的位置：", first_full_stop_position)
                                if first_full_stop_position < first_colon_position:
                                    if last_character == "。":
                                        # tips = '这句话前面要调整下'
                                        front_str = content[: first_full_stop_position + 1]
                                        ladder_str = content[first_full_stop_position + 1 :]
                                        content = ladder_str
                                        str_update_list.append(
                                            (order_number - 1, front_str)
                                        )
                                        new_ytms_list.append((order_number, content))
                                    else:
                                        front_str = content[: first_full_stop_position + 1]
                                        ladder_str = content[first_full_stop_position + 1 :]
                                        content = ladder_str
                                        str_update_list.append((order_number, front_str))
                                        # print('Front:',front_str,'Ladder:',ladder_str)
                                        new_ytms_list.append((order_number, content))
                                else:
                                    new_ytms_list.append((order_number, content))
                                # print(content,tips)

                            # str_update_list顺序调整
                            current_update_order = 999
                            mod_str_update_list = []
                            for item in reversed(str_update_list):
                                order_number = item[0]
                                content = item[1]
                                if order_number != current_update_order:
                                    current_update_order = order_number
                                    # order = current_update_order
                                else:
                                    order_number = order_number - 1
                                    current_update_order = order_number
                                mod_str_update_list.append((order_number, content))

                            str_update_list = mod_str_update_list
                            for item in new_ytms_list:
                                print("新主体部分：" + str(item))
                            for item in str_update_list:
                                print("新补充部分：" + str(item))
                            # exit()
                            temp_list = []
                            for item in new_ytms_list:
                                # index = item[0]
                                order_number = item[0]
                                content = item[1]
                                for item in str_update_list:
                                    order_number_update = item[0]
                                    content_update = item[1]
                                    if order_number == order_number_update:
                                        content = content + content_update
                                temp_list.append((order_number, content))
                            # print(content)
                        # for item in temp_list:
                        #     print('新拼接结果：',item)
                        ytms_list = temp_list
                        # exit()
                        ####################################################################################
                        # for item in ytms_list:
                        #     print(item)
                        # exit()

                        for item in target_value_list:  # 分层信息表all in one
                            range_id = item[0]  # 柱状图标识
                            depth_order_number = item[1]  # 分层标识序号
                            depth = item[2]  # 层底深度值
                            time_reason = item[3]  # 时代成因
                            PM_order = item[4]  # 剖面层号
                            ytms_text = ""  # 岩土描述
                            drilling_number = range_id_drilling_number_chart_dict[
                                range_id
                            ]  # 钻孔编号
                            for ytms in ytms_list:
                                if ytms[0] == depth_order_number:
                                    ytms_text = ytms[1]
                            Stratification_infomation_list.append(
                                (
                                    range_id,
                                    depth_order_number,
                                    depth,
                                    time_reason,
                                    PM_order,
                                    ytms_text,
                                )
                            )  # 分层信息表
                            if ytms_text == "":
                                ytms_text = "&这层的描述啥也没写&当然也可能是翻车了&"
                            print(
                                "【柱状图标识】",
                                range_id,
                                "【钻孔编号】",
                                str(drilling_number),
                                "【层序(y坐标降序)】",
                                depth_order_number,
                                "【剖面层号】",
                                PM_order,
                                "【层底深度】",
                                depth,
                                "【时代成因】",
                                time_reason,
                                "【岩土描述】",
                                ytms_text,
                            )
                            txt.write(
                                "【柱状图标识】"
                                + str(range_id)
                                + " "
                                + "【钻孔编号】"
                                + str(drilling_number)
                                + " "
                                + "【层序(y坐标降序)】"
                                + str(depth_order_number)
                                + "【剖面层号】"
                                + str(PM_order)
                                + "【层底深度】"
                                + str(depth)
                                + "【时代成因】"
                                + str(time_reason)
                                + "【岩土描述】"
                                + str(ytms_text)
                                + "\n"
                            )

                Separation_info_list = (
                    Separation_info_list + Stratification_infomation_list
                )  # 合并到总表
                # print('')
                # print('是不是要重新拼接下：')
                # print('')
                # exit()
                # B2.标贯、土工信息表识别
                TG_BG_list = []
                for GG in TG_BG_range_list:
                    range_id = GG[0]
                    key_repr_field_name = GG[1]
                    field_name_in_cad = GG[2]
                    GG_range = GG[3]
                    GG_list = [
                        value
                        for value in text_list
                        if coor_inside_range(
                            value[2][0],
                            value[2][1],
                            GG_range[0],
                            GG_range[1],
                            GG_range[2],
                            GG_range[3],
                        )
                    ]  # 每个字段下方格子内的文本对象集合
                    TG_BG_list.append(
                        (
                            range_id,
                            key_repr_field_name,
                            field_name_in_cad,
                            GG_range,
                            GG_list,
                        )
                    )
                new_TG_BG_list = []
                for item in TG_BG_list:
                    range_id = item[0]
                    key_repr_field_name = item[1]
                    # print('字段名:', key_repr_field_name)
                    field_name_in_cad = item[2]
                    GG_range = item[3]
                    GG_list = item[4]
                    new_GG_List = []
                    used_row_list = []
                    for row in GG_list:
                        if row not in used_row_list:
                            used_row_list.append(row)
                            range_id = row[0]
                            content = row[1]
                            coordinates = row[2]
                            obj_id = row[3]
                            bounding_box = row[4]
                            min_y = bounding_box[0][1]
                            max_y = bounding_box[1][1]
                            new_row_list = [row]
                            for row in GG_list:
                                if row not in used_row_list:
                                    range_id_1 = row[0]
                                    content_1 = row[1]
                                    coordinates_1 = row[2]
                                    obj_id_1 = row[3]
                                    bounding_box_1 = row[4]
                                    min_y_1 = bounding_box_1[0][1]
                                    max_y_1 = bounding_box_1[1][1]
                                    if (
                                        (min_y >= min_y_1 and min_y <= max_y_1)
                                        or (min_y_1 >= min_y and min_y_1 <= max_y)
                                        or (max_y >= min_y_1 and max_y <= max_y_1)
                                        or (max_y_1 >= min_y and max_y_1 <= max_y)
                                    ):
                                        new_row_list.append(row)
                                        used_row_list.append(row)
                            new_row_list = sorted(new_row_list, key=lambda x: x[2][0])
                            new_content_list = [content[1] for content in new_row_list]
                            new_content = " @ ".join(new_content_list)
                            print("下方内容：", new_content, coordinates)
                            new_GG_List.append(
                                (
                                    range_id,
                                    new_content,
                                    coordinates,
                                    obj_id,
                                    bounding_box,
                                )
                            )
                            # print('同一行的文本对象：', new_content)
                    new_TG_BG_list.append(
                        (
                            range_id,
                            key_repr_field_name,
                            field_name_in_cad,
                            GG_range,
                            new_GG_List,
                        )
                    )

                # for item in new_TG_BG_list:
                #     key_repr_field_name = item[1]
                #     print('字段名:', key_repr_field_name)
                #     GG_list = item[4]
                #     for item in GG_list:
                #         print(item)
                TG_BG_list = new_TG_BG_list  # 同行文本合并后的结果列表
                # exit()

                if single_or_multiple_column == 0:  # 单列(多为分数形式)
                    print("")
                    print("--------土工标贯为分数形式--------")
                    txt.write("【土工标贯为分数形式】\n")
                    concatenated_text_list = []
                    for item in TG_BG_list:
                        key_repr_field_name = item[1]  # 目标字段名
                        field_name_in_cad = item[2]  # CAD中的字段名
                        content_list = item[4]  # 对应字段名下方格子内的文本对象列表
                        content_list = sorted(
                            content_list, key=lambda y: [-y[2][1]]
                        )  # y降序
                        print("cad里的字段名：", field_name_in_cad)
                        for content in content_list:
                            print(field_name_in_cad, " 下方的内容：", content)
                        used_text_list = []  # 已判断文本对象列表
                        text_obj_count = len(content_list)
                        print(text_obj_count)
                        if text_obj_count % 2 == 0:
                            print(
                                key_repr_field_name,
                                "：字段下方文本对象数为偶数，按y坐标降序两两一组连接",
                            )
                            # if text_obj_count % 2 == 0:
                            molecular_list = content_list[::2]  # 分子(坐标组奇数位)
                            denominator_list = content_list[1::2]  # 分母(坐标组偶数位)
                            # else:
                            #     molecular_list = content_list  # 分子(坐标组奇数位)
                            #     denominator_list = '这个应该是单行的字符串'
                            # for i in molecular_list:
                            #     print(i)
                            # for i in denominator_list:
                            #     print(i)
                            fraction_value_list = [
                                (
                                    range_id,
                                    key_repr_field_name,
                                    molecular[1],
                                    denominator[1],
                                )
                                for molecular, denominator in zip(
                                    molecular_list, denominator_list
                                )
                            ]
                            concatenated_text_list = (
                                concatenated_text_list + fraction_value_list
                            )
                        else:
                            print(
                                key_repr_field_name,
                                "：某些分数不完整，判断不了了，自己打吧(分数可能不是仅由两个文本对象组成)",
                            )
                            if extend_bottom_height != 0:
                                print(
                                    "*另外可能需要留意的是，当前底部边框增加了高度: "
                                    + str(extend_bottom_height)
                                )
                            continue
                        TG_BG_list = concatenated_text_list
                    print("--------------------------------------")

                elif single_or_multiple_column == 1:  # 多列(跨列关联、单行):
                    print("")
                    print("--------土工标贯为单行形式--------")
                    txt.write("--------土工标贯为单行形式--------\n")
                    concatenated_text_list = []
                    field_name_in_cad = " & ".join(
                        [item[2] for item in TG_BG_list]
                    )  # CAD中的字段名(合并)

                    all_content_list = sum(
                        [item[4] for item in TG_BG_list], []
                    )  # 所有字段名下方格子内的文本对象合并列表
                    # print(field_name_in_cad,'\n',all_content_list)

                    # print(range_id,key_repr_field_name,"下方内容有：")
                    # for value in concatenated_text_list:
                    #     print(' ',value[1])
                    # exit()
                    used_text_list = []
                    for item in all_content_list:
                        to_be_cocatenated_list = []
                        if item not in used_text_list:
                            to_be_cocatenated_list.append(item)
                            used_text_list.append(item)
                            content = item[1]
                            bottom_y = item[4][0][1]
                            top_y = item[4][1][1]
                            for item1 in all_content_list:
                                if item1 not in used_text_list:
                                    # used_text_list_1.append(item_1)
                                    content1 = item1[1]
                                    bottom_y1 = item1[4][0][1]
                                    top_y1 = item1[4][1][1]
                                    if (
                                        (bottom_y >= bottom_y1 and bottom_y <= top_y1)
                                        or (top_y >= bottom_y1 and top_y <= top_y1)
                                        or (bottom_y1 >= bottom_y and bottom_y1 <= top_y)
                                        or (top_y1 >= bottom_y and top_y1 <= top_y)
                                    ):
                                        to_be_cocatenated_list.append(item1)
                                        used_text_list.append(item1)
                        # for item in to_be_cocatenated_list:
                        #     print(item)
                        cocatenated_content = " @ ".join(
                            [
                                text_content[1]
                                for text_content in sorted(
                                    to_be_cocatenated_list, key=lambda x: [x[2][0]]
                                )
                            ]
                        )
                        try:
                            cocatenated_content_y = statistics.mean(
                                [
                                    y[2][1]
                                    for y in sorted(
                                        to_be_cocatenated_list, key=lambda x: [x[2][0]]
                                    )
                                ]
                            )
                            concatenated_text_list.append(
                                (
                                    range_id,
                                    cocatenated_content,
                                    cocatenated_content_y,
                                    "【单行多列】",
                                )
                            )
                        except:
                            continue
                        # print(cocatenated_content,cocatenated_content_y)
                        # print('[',cocatenated_content,']')
                        # print('+++++++++++++++')

                    # print('-----------------------------------------')
                    # concatenated_text_list = all_content_list
                    TG_BG_list = sorted(concatenated_text_list, key=lambda y: [-y[2]])

                for item in TG_BG_list:
                    drilling_number = range_id_drilling_number_chart_dict[
                        range_id
                    ]  # 钻孔编号
                    print(drilling_number, item)
                    txt.write(
                        str(item[0])
                        + "	"
                        + str(drilling_number)
                        + "	"
                        + str(item[1])
                        + "	"
                        + str(item[2])
                        + "	"
                        + str(item[3])
                        + "\n"
                    )
                TG_BG_info_list = TG_BG_info_list + TG_BG_list  # 合并到总表
            ##################################################################################
            ##################################################################################
            ##################################################################################
            end_time = time.time()
            end_times = get_current_time()
            time_cnsumption = format(end_time - start_time, ".2f")
            print("")
            print(
                "************************************************************************"
            )
            txt.write(
                "************************************************************************"
                + "\n"
            )
            if use_insertion_point == 1:
                print("文本对象(含块参照)用【插入点】坐标")
                txt.write("文本对象(含块参照)用【插入点】坐标" + "\n")
            else:
                print("文本对象(含块参照)用【中心点】坐标")
                txt.write("文本对象(含块参照)用【中心点】坐标" + "\n")
            print("--------------------------------------------")
            print(
                "柱状图数量："
                + str(max_frame_count)
                + "\n"
                + "开始时间："
                + start_times
                + "\n"
                + "结束时间："
                + end_times
                + "\n"
                + "耗时："
                + str(time_cnsumption)
                + "秒"
            )
            txt.write(
                "柱状图数量："
                + str(max_frame_count)
                + "\n"
                + "开始时间："
                + start_times
                + "\n"
                + "结束时间："
                + end_times
                + "\n"
                + "耗时："
                + str(time_cnsumption)
                + "秒"
                + "\n"
            )
            txt.write(
                "************************************************************************"
                + "\n"
            )
            txt.write("本文件路径：" + txt_path)
            txt.close()
            # os.startfile(txt_path)  #打开txt
            print("")
            print("====================【 汇    总】====================")
            print("")
            print("====================钻孔信息表====================")
            drill_hole_txt_path = result_path + "/" + "钻孔信息表.txt"
            txt = open(drill_hole_txt_path, "w")
            DrillHole_range_id_list = list(set([id[0] for id in DrillHole_info_list]))
            field_name_list = [name[1] for name in DrillHole_info_list]
            DrillHole_field_name_list = list(set(field_name_list))  # 所有字段名称
            DrillHole_field_name_list.sort(key=field_name_list.index)

            DrillHole_field_name_list.insert(0, "柱状图标识")
            DrillHole_field_names = "	".join(DrillHole_field_name_list)
            print(DrillHole_field_names)
            txt.write(DrillHole_field_names + "\n")
            for id in DrillHole_range_id_list:
                value_list = [value[3] for value in DrillHole_info_list if value[0] == id]
                value = "	".join(value_list)
                print(str(id) + "	" + value)
                txt.write(str(id) + "	" + value + "\n")
            txt.close()
            print("====================分层信息表====================")
            Separation_txt_path = result_path + "/" + "分层信息表.txt"
            txt = open(Separation_txt_path, "w")
            Separation_field_names = (
                "柱状图标识	钻孔编号	层序	层底深度	时代成因	剖面层号	岩土描述"
            )
            txt.write(Separation_field_names + "\n")
            print(Separation_field_names)
            for item in Separation_info_list:
                drilling_number = range_id_drilling_number_chart_dict[item[0]]
                item_list = [
                    str(item[0]),
                    str(drilling_number),
                    str(item[1]),
                    str(item[2]),
                    str(item[3]),
                    str(item[4]),
                    str(item[5]),
                ]
                value = "	".join(item_list)
                print(value)
                txt.write(value + "\n")
            txt.close()
            print("====================土工、标贯信息表====================")
            TG_BG_txt_path = result_path + "/" + "土工标贯信息表.txt"
            txt = open(TG_BG_txt_path, "w")
            if single_or_multiple_column == 0:
                TG_BG_field_names = "柱状图标识	钻孔编号	类型	分子	分母"
            else:
                TG_BG_field_names = "柱状图标识	钻孔编号	内容	y坐标	凑数的列"
            txt.write(TG_BG_field_names + "\n")
            print(TG_BG_field_names)
            for item in TG_BG_info_list:
                drilling_number = range_id_drilling_number_chart_dict[item[0]]
                item_list = [
                    str(item[0]),
                    str(drilling_number),
                    str(item[1]),
                    str(item[2]),
                    str(item[3]),
                ]
                value = "	".join(item_list)
                print(value)
                txt.write(value + "\n")
            txt.close()
            print("\n")
            print("程序运行结束")
            win32api.MessageBox(
                0,
                "完成【"
                + cad_name
                + "】的读取！"
                + "\n"
                + "开始时间："
                + start_times
                + "\n"
                + "结束时间："
                + end_times
                + "\n"
                + "耗时："
                + str(time_cnsumption)
                + "秒",
                "OK",
                win32con.MB_OK | win32con.MB_TOPMOST,
            )
            
            os.startfile(result_path)
        except Exception as e:
                import traceback
                error_msg = f"\n\n👉 😅😓🤡😭😩😟🤮🤬 👈\n程序出错：\n{str(e)}\n{traceback.format_exc()}"
                sys.stderr.write(error_msg)  # 会触发重定向到UI

app = SampleApp()
app.mainloop()
