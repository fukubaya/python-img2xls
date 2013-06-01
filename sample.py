#!/usr/bin/python
# -*- coding: utf-8 -*-

# 
# sample.py
# 
# Created by FUKUBAYASHI Yuichiro on 2013/05/03
# Copyright (c) 2013, FUKUBAYASHI Yuichiro
# 
# last update: <2013/06/01 15:47:13>
# 

import sys
import os

import xlwt

import img2xls

# custom generator
def coloured_cell_and_black_text_style_generator(scale, fontname='Arial'):
    def __coloured_cell_and_black_text_style_generator(colour_index):
        style = xlwt.XFStyle()
        
        fnt = xlwt.Font()
        fnt.colour_index = 0x00 # black
        fnt.height = int(200*scale)
        fnt.name = fontname
        style.font = fnt
        
        align = xlwt.Alignment()
        align.horz = xlwt.Alignment.HORZ_CENTER
        align.vert = xlwt.Alignment.VERT_CENTER
        style.alignment = align
        
        pattern = xlwt.Pattern()
        pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        pattern.pattern_fore_colour = colour_index
        style.pattern = pattern
        
        return style

    return __coloured_cell_and_black_text_style_generator



if __name__ == '__main__':
    argv = sys.argv
    imgpath = argv[1]

    # basename
    base, ext = os.path.splitext(imgpath)

    # default (fill cell with image colour)
    workbook = img2xls.image_to_workbook(imgpath=imgpath,
                                         scale=0.25)

    # coloured text
    #workbook = img2xls.image_to_workbook(imgpath=imgpath,
    #                                     scale=0.25,
    #                                     style_generator=img2xls.coloured_text_style_generator(0.25, u'メイリオ'),
    #                                     text_for_pixel=u"セルに色を塗るかわりに色のついた文字でピクセルを表現。")

    # custom style (coloured cell and black text)
    #workbook = img2xls.image_to_workbook(imgpath=imgpath,
    #                                    scale=1.0,
    #                                     style_generator=coloured_cell_and_black_text_style_generator(1.0, u'メイリオ'),
    #                                     text_for_pixel=u"セルを塗り潰して、黒の文字を入れる。")

    # savexls
    xlspath = "%s.xls" % (base)    
    workbook.save(xlspath)

