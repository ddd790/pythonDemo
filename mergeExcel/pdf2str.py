from PyPDF2 import PdfFileReader, PdfFileWriter
from aip import AipOcr, AipSpeech
import pdfkit
import fitz
import os


pdfpath = 'F:\资料\其他备用资料'
pdfname = '部门晨读《六项精进》21.pdf'
path_wk = r'D:/wkhtmltopdf/bin/wkhtmltopdf.exe'

# 百度云的账户
APP_ID = '25101742'
API_KEY = 'Z5qy26GRDUdDKlBRHGT21XZt'
SECRET_KEY = 'p6BCz0xxGXSTbDR3MfWAfViBRbFilaAu'

# 百度文字转语音
APP_ID_MP3 = '25215794'
API_KEY_MP3 = '0QGcLRXp8AwSQAQD24O09EoC'
SECRET_KEY_MP3 = 'REgLKIcrC3uWHoTZ9Wk5bKnV1xzm8X9F'

# 以下为处理程序---------------------------------------------------------------------------
pdfkit_config = pdfkit.configuration(wkhtmltopdf=path_wk)
pdfkit_options = {'encoding': 'UTF-8', }


# 将每页pdf转为png格式图片
def pdf_image():
    pdf = fitz.open(pdfpath+os.sep+pdfname)
    for pg in range(0, pdf.pageCount):
        # 获得每一页的对象
        page = pdf[pg]
        trans = fitz.Matrix(1.0, 1.0).prerotate(0),
        # 获得每一页的流对象
        pm = page.get_pixmap(matrix=trans, alpha=False)
        # 保存图片
        pm.save(image_path + os.sep +
                pdfname[:-4] + '_' + '{:0>3d}.png'.format(pg + 1))
    page_range = range(pdf.pageCount)
    pdf.close()
    return page_range


def read_png_str(page_range):
    # 读取本地图片的函数
    def get_file_content(filePath):
        with open(filePath, 'rb') as fp:
            return fp.read()

    all_pngstr = []
    image_list = []
    for page_num in page_range:
        # 读取本地图片
        image = get_file_content(
            image_path + os.sep + r'{}_{}.png'.format(pdfname[:-4], '%03d' % (page_num + 1)))
        image_list.append(image)

    # 新建一个AipOcr
    client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
    options = {}
    options["language_type"] = "CHN_ENG"
    options["detect_direction"] = "false"
    options["detect_language"] = "false"
    options["probability"] = "false"
    for image in image_list:
        # 文字识别,得到一个字典
        pngjson = client.basicGeneral(image, options)
        pngstr = ''
        for x in pngjson['words_result']:
            # pngstr = pngstr + x['words'] + '</br>'  # 转成PDF时换行用</br>
            pngstr = pngstr + x['words']
        # print('正在调用百度接口：第{}个，共{}个'.format(len(all_pngstr), len(image_list)))
        all_pngstr.append(pngstr)
    return all_pngstr


def str2pdf(page_range, all_pngstr):
    # 字符串写入PDF
    for page_num in page_range:
        print('正在将字符串写入PDF：第{}个，共{}个'.format((page_num + 1), len(page_range)))
        pdfkit.from_string((all_pngstr[page_num]), disperse_pdfpath + os.sep + '%s.pdf' % (str(page_num + 1)),
                           configuration=pdfkit_config, options=pdfkit_options)


def pdf_merge(page_range):
    # 合并单页PDF
    pdf_output = PdfFileWriter()
    for page_num in page_range:
        print('正在合并单页：第{}个，共{}个'.format((page_num + 1), len(page_range)))
        pdf_input = PdfFileReader(
            open(disperse_pdfpath + os.sep + '%s.pdf' % (str(page_num + 1)), 'rb'))
        page = pdf_input.getPage(0)
        pdf_output.addPage(page)
    newPdfPath = pdfpath+os.sep + 'new_{}'.format(pdfname)
    pdf_output.write(open(newPdfPath, 'wb'))
    return newPdfPath


image_path = pdfpath + os.sep + "image"
if not os.path.exists(image_path):
    os.mkdir(image_path)

disperse_pdfpath = pdfpath + os.sep + "pdf"
if not os.path.exists(disperse_pdfpath):
    os.mkdir(disperse_pdfpath)

range_count = pdf_image()
all_th = read_png_str(range_count)
read_word = str(','.join(all_th)).replace(',M', '芜,').replace('MOTIVES', '')
read_word = read_word if len(read_word) < 512 else read_word[0:512]

client = AipSpeech(APP_ID_MP3, API_KEY_MP3, SECRET_KEY_MP3)
result = client.synthesis(read_word, 'zh', 1, {
    'vol': 5,
    'per': 1
})

# 识别正确返回语音二进制 错误则返回dict 参照下面错误码
if not isinstance(result, dict):
    with open('生成文件.mp3', 'wb') as f:
        f.write(result)

# print(read_word)
# str2pdf(range_count, all_th)
# pdf_merge(range_count)
