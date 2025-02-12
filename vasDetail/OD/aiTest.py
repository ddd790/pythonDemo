import http.client
import mimetypes
from codecs import encode 

class VAS_GUI():
    def get_files(self):
        conn = http.client.HTTPSConnection("api.openai.com")
        dataList = []
        boundary = 'wL36Yn8afVp8Ag7AmP8qZ0SA4n1v9T'
        dataList.append(encode('--' + boundary))
        dataList.append(encode('Content-Disposition: form-data; name=file; filename={0}'.format('')))

        fileType = mimetypes.guess_type('')[0] or 'application/octet-stream'
        dataList.append(encode('Content-Type: {}'.format(fileType)))
        dataList.append(encode(''))

        with open('F:\\下载\\001.txt', 'rb') as f:
            dataList.append(f.read())
            dataList.append(encode('--' + boundary))
            dataList.append(encode('Content-Disposition: form-data; name=model;'))

            dataList.append(encode('Content-Type: {}'.format('text/plain')))
            dataList.append(encode(''))

            dataList.append(encode(""))
            dataList.append(encode('--' + boundary))
            dataList.append(encode('Content-Disposition: form-data; name=prompt;'))

            dataList.append(encode('Content-Type: {}'.format('text/plain')))
            dataList.append(encode(''))

            dataList.append(encode(""))
            dataList.append(encode('--' + boundary))
            dataList.append(encode('Content-Disposition: form-data; name=response_format;'))

            dataList.append(encode('Content-Type: {}'.format('text/plain')))
            dataList.append(encode(''))

            dataList.append(encode(""))
            dataList.append(encode('--' + boundary))
            dataList.append(encode('Content-Disposition: form-data; name=temperature;'))

            dataList.append(encode('Content-Type: {}'.format('text/plain')))
            dataList.append(encode(''))

            dataList.append(encode(""))
            dataList.append(encode('--'+boundary+'--'))
            dataList.append(encode(''))
            body = b'\r\n'.join(dataList)
        payload = body
        headers = {
            'User-Agent': 'Apifox/1.0.0 (https://apifox.com)',
            'Content-type': 'multipart/form-data; boundary={}'.format(boundary)
        }
        conn.request("GET", "/v1/audio/translations", payload, headers)
        res = conn.getresponse()
        data = res.read()
        print(data.decode("utf-8"))
        
def gui_start():
    VAS = VAS_GUI()
    VAS.get_files()


if __name__ == '__main__':
    gui_start()
