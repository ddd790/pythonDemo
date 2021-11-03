from PyPDF2 import PdfFileReader, PdfFileWriter
from aip import AipOcr
import pdfkit
import fitz
import os


pdfpath = 'F:\资料\其他备用资料'
pdfname = '部门晨读《六项精进》21.pdf'
path_wk = r'D:/wkhtmltopdf/bin/wkhtmltopdf.exe'


APP_ID = '25101742'
API_KEY = 'Z5qy26GRDUdDKlBRHGT21XZt'
SECRET_KEY = 'p6BCz0xxGXSTbDR3MfWAfViBRbFilaAu'

# 以下为处理程序---------------------------------------------------------------------------
pdfkit_config = pdfkit.configuration(wkhtmltopdf=path_wk)
pdfkit_options = {'encoding': 'UTF-8', }


# 将每页pdf转为png格式图片
def pdf_image():
    pdf = fitz.open(pdfpath+os.sep+pdfname)
    for pg in range(0, pdf.pageCount):
        # 获得每一页的对象
        page = pdf[pg]
        trans = fitz.Matrix(1.0, 1.0).preRotate(0),
        # 获得每一页的流对象
        pm = page.getPixmap(matrix=trans, alpha=False)
        # 保存图片
        pm.writePNG(image_path + os.sep +
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
            pngstr = pngstr + x['words'] + '</br>'
        print('正在调用百度接口：第{}个，共{}个'.format(len(all_pngstr), len(image_list)))
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
print(all_th)
# str2pdf(range_count, all_th)
# pdf_merge(range_count)
