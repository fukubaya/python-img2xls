======================================================================
python-img2xls
======================================================================
Created by FUKUBAYASHI Yuichiro on 2013/06/01

About
======================================================================
pngやjpgなどの画像ファイルをxlsに変換する


Requirements
======================================================================
PIL
 http://www.pythonware.com/products/pil/

Python Excel
 http://www.python-excel.org/


How to use
======================================================================
img2xlsをimportして，画像ファイルのpathを渡すだけ．

::

    workbook = img2xls.image_to_workbook(imgpath=imgpath,
                                         scale=0.25,
                                         style_generator=img2xls.coloured_text_style_generator(0.25, u'メイリオ'),
                                         text_for_pixel=u"セルに色を塗るかわりに色のついた文字でピクセルを表現。")
    workbook.save("someimg.xls")


引数は以下のとおり

imgpath
 画像ファイルのファイルパス

scale (default=1.0)
 セルの大きさ．1.0でデフォルトのexcelの高さと同じになる．

style_generator (default=img2xls.coloured_cell_style_generator)
 セルにどのように色を付けるかの指定．この例では色付きのテキストを使用し，フォントはメイリオ，2.5pt (10pt * 0.25)．

text_for_pixel (default=None)
 テキストで塗り潰す場合のテキスト．日本語の場合はunicodeで渡す．""もしくはNoneを指定すると何も書かない．


sample.py
======================================================================
::

 % python sample.py someimg.jpg
 % ls someimg.xls
 someimg.xls
 % 


LICENSE
======================================================================
The MIT License

Copyright (c) 2013 FUKUBAYASHI Yuichiro

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「ソフトウェア」）の複製を取得するすべての人に対し、ソフトウェアを無制限に扱うことを無償で許可します。これには、ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、およびソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。

上記の著作権表示および本許諾表示を、ソフトウェアのすべての複製または重要な部分に記載するものとします。

ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての保証も含みますが、それに限定されるものではありません。 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、ソフトウェアに起因または関連し、あるいはソフトウェアの使用またはその他の扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わないものとします。
