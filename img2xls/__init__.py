# -*- coding: utf-8 -*-

# 
# Created by FUKUBAYASHI Yuichiro on 2013/05/03
# Copyright (c) 2013, FUKUBAYASHI Yuichiro
# 
# last update: <2013/06/01 16:08:05>
# 

import os.path

from PIL import Image
import xlwt



MAX_ROW = 65536
MAX_COL = 256

NUM_COLOUR = 56

BASE_COL_WIDTH = 740.5
BASE_ROW_HEIGHT = 301.3


def check_size_and_resize_image(img):
    img_w, img_h = img.size
    if img_w > MAX_COL:
        resize_scale = float(MAX_COL)/img_w
        img = img.resize((MAX_COL, int(img_h * resize_scale)))

    img_w, img_h = img.size
    if img_h > MAX_ROW:
        resize_scale = float(MAX_ROW)/img_h
        img = img.resize((int(img_w * resize_scale), MAX_ROW))

    return img


def get_rgblist_from_palette(palette):
    rgblist = []

    tmprgb = [None]*3
    i=0
    for c in palette:
        tmprgb[i] = c
        i+=1
        if i == 3:
            rgblist.append((tmprgb[0],tmprgb[1],tmprgb[2]))
            i = 0
            tmprgb = [None]*3
            
    return rgblist


def set_colour_list(wkbk, rgblist):
    for i in range(len(rgblist)):
        colour_index = i + 8
        r, g, b = rgblist[i]
        wkbk.set_colour_RGB(colour_index, r, g, b)


def set_cell_size(sheet, width, height, scale=1.0):
    for x in range(width):
        sheet.col(x).width = int(BASE_COL_WIDTH*scale)

    for y in range(height):
        sheet.row(y).height_mismatch = 1
        sheet.row(y).height = int(BASE_ROW_HEIGHT*scale)


def coloured_cell_style_generator(colour_index):
    style = xlwt.XFStyle()
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern.pattern_fore_colour = colour_index
    style.pattern = pattern

    return style


def coloured_text_style_generator(scale, fontname='Arial'):
    def __coloured_text_style_generator(colour_index):
        style = xlwt.XFStyle()

        fnt = xlwt.Font()
        fnt.colour_index = colour_index
        fnt.height = int(200*scale)
        fnt.name = fontname
        style.font = fnt

        align = xlwt.Alignment()
        align.horz = xlwt.Alignment.HORZ_CENTER
        align.vert = xlwt.Alignment.VERT_CENTER
        style.alignment = align
        
        return style

    return __coloured_text_style_generator


def generate_styles(style_generator=coloured_cell_style_generator):
    styles = []

    for i in range(NUM_COLOUR):
        styles.append(style_generator(i+8))
        
    return styles


def image_to_workbook(imgpath, 
                      scale=1.0,
                      style_generator=coloured_cell_style_generator,
                      text_for_pixel=None):

    # open image
    img = Image.open(imgpath)

    # check size (and resize)
    img = check_size_and_resize_image(img)
    img_w, img_h = img.size

    # convert image (reduce colours)
    pimg = img.convert('RGB').convert('P', palette=Image.ADAPTIVE, colors=NUM_COLOUR)

    # get colour list from image
    rgblist = get_rgblist_from_palette(pimg.getpalette())[0:NUM_COLOUR]

    # create workbook
    wkbk = xlwt.Workbook()
    sheetname = os.path.basename(imgpath)
    sheet = wkbk.add_sheet(sheetname)

    # set custom colour palette to workbook
    set_colour_list(wkbk, rgblist)

    # set width and height of cell
    set_cell_size(sheet, img_w, img_h, scale)

    # load pixels
    pix = pimg.load()
    
    # get style mapping
    styles = generate_styles(style_generator)
    txt_counter = 0
    
    for y in range(img_h):
        for x in range(img_w):
            colour_index = pix[x,y]

            char_to_write = None
            if (text_for_pixel is not None and 
                len(text_for_pixel) > 0):
                char_to_write = text_for_pixel[txt_counter%len(text_for_pixel)]
            else:
                char_to_write = ''

            sheet.write(y, x, char_to_write, styles[int(colour_index)])
            txt_counter += 1

    return wkbk

