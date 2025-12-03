from flask import Flask, request, jsonify, render_template
import zipfile
import base64
from lxml import etree
from io import BytesIO
import os

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/extract', methods=['POST'])
def extract_power_queries():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    try:
        # 检查文件类型
        if not file.filename.endswith(('.xlsx', '.xlsm', '.xlsb')):
            return jsonify({'error': 'Unsupported file format. Please use .xlsx, .xlsm, or .xlsb files.'}), 400
        
        # 读取文件内容
        file_content = file.read()
        
        # 打开Excel文件作为ZIP归档
        with zipfile.ZipFile(BytesIO(file_content), 'r') as zip_ref:
            queries = []
            
            # 查找customXml目录下的XML文件
            custom_xml_files = [f for f in zip_ref.namelist() if f.startswith('customXml/') and f.endswith('.xml')]
            
            for xml_path in custom_xml_files:
                with zip_ref.open(xml_path) as xml_file:
                    xml_content_bytes = xml_file.read()
                    root = etree.fromstring(xml_content_bytes)
                    
                    # 定义命名空间
                    namespace = {'d': 'http://schemas.microsoft.com/DataMashup'}
                    
                    # 查找DataMashup元素
                    data_mashup_elements = root.xpath('//d:DataMashup', namespaces=namespace)
                    
                    if data_mashup_elements:
                        # 解码Base64内容
                        base64_content = data_mashup_elements[0].text
                        decoded_content = base64.b64decode(base64_content)
                        
                        # 查找嵌入式ZIP文件
                        zip_start = decoded_content.find(b'PK\x03\x04')
                        zip_end = decoded_content.find(b'PK\x05\x06')
                        
                        if zip_start != -1 and zip_end != -1:
                            # 提取ZIP数据
                            zip_data = BytesIO(decoded_content[zip_start:zip_end + 22])
                            
                            with zipfile.ZipFile(zip_data) as archive:
                                # 检查Section1.m文件
                                if 'Formulas/Section1.m' in archive.namelist():
                                    section1_content = archive.read('Formulas/Section1.m').decode('utf-8')
                                    queries.extend(parse_m_file(section1_content))
                                elif 'formulas/section1.m' in archive.namelist():
                                    section1_content = archive.read('formulas/section1.m').decode('utf-8')
                                    queries.extend(parse_m_file(section1_content))
            
            return jsonify({'queries': queries})
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

def parse_m_file(content):
    queries = []
    
    # 查找shared查询
    import re
    shared_regex = re.compile(r'shared\s+([^\s=]+)\s*=\s*([\s\S]*?)(?=\s*shared\s+|$)')
    
    matches = shared_regex.finditer(content)
    for match in matches:
        query_name = match.group(1).strip()
        query_code = match.group(2).strip()
        
        # 清理代码格式
        query_code = re.sub(r'^\s+', '', query_code, flags=re.MULTILINE)
        query_code = re.sub(r'\s+$', '', query_code, flags=re.MULTILINE)
        
        queries.append({
            'name': query_name,
            'code': query_code
        })
    
    return queries

if __name__ == '__main__':
    app.run(debug=True, port=5000)
