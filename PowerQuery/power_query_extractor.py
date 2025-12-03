import zipfile
import os
import base64
from lxml import etree
from io import BytesIO
import sys

def find_power_query_files(excel_path):
    # List to store the paths of Power Query files found within the Excel file
    power_query_files = []
    
    # Check if the provided file is an Excel file based on its extension
    if not excel_path.endswith(('.xlsx', '.xlsm', '.xlsb')):
        raise ValueError("Unsupported file format. Please use .xlsx, .xlsm, or .xlsb files.")
    
    # Open the Excel file as a zip archive
    with zipfile.ZipFile(excel_path, 'r') as zip_ref:
        print("Excel文件包含的文件列表：")
        for file_info in zip_ref.infolist():
            print(f"  - {file_info.filename}")
        
        # Iterate through each file in the zip archive
        for file_info in zip_ref.infolist():
            # Check if the file is located within the customXml directory and has an xml extension
            if file_info.filename.startswith('customXml/') and file_info.filename.endswith('.xml'):
                print(f"\n检查文件：{file_info.filename}")
                # Open and read the content of the XML file
                with zip_ref.open(file_info) as xml_file:
                    xml_content_bytes = xml_file.read()
                    try:
                        # Parse the XML content
                        root = etree.fromstring(xml_content_bytes)
                        # Define the namespace used in the DataMashup elements
                        namespace = {'d': 'http://schemas.microsoft.com/DataMashup'}
                        # Search for DataMashup elements within the XML document
                        data_mashup_elements = root.xpath('//d:DataMashup', namespaces=namespace)
                        if data_mashup_elements:
                            print("找到DataMashup元素")
                            # If found, decode the base64 content of the DataMashup element
                            base64_content = data_mashup_elements[0].text
                            print(f"Base64内容长度：{len(base64_content)}")
                            decoded_content = base64.b64decode(base64_content)
                            print(f"解码后内容长度：{len(decoded_content)}")
                            # Look for the ZIP archive signatures to find the embedded ZIP archive
                            zip_start = decoded_content.find(b'PK\x03\x04') # Start of ZIP archive
                            zip_end = decoded_content.find(b'PK\x05\x06') # End of ZIP archive (end of central directory record)
                            print(f"ZIP开始位置：{zip_start}, ZIP结束位置：{zip_end}")
                            if zip_start != -1 and zip_end != -1:
                                # Extract the ZIP archive from the decoded content
                                zip_data = BytesIO(decoded_content[zip_start:zip_end + 22]) # Include the EOCD size
                                with zipfile.ZipFile(zip_data) as archive:
                                    print("ZIP Archive Contents:", archive.namelist())
                                    # Check for the presence of the 'Formulas/Section1.m' file, which contains Power Query formulas
                                    if 'Formulas/Section1.m' in archive.namelist():
                                        # Read and print the content of 'Formulas/Section1.m'
                                        section1_m_content = archive.read('Formulas/Section1.m').decode('utf-8')
                                        print("\n=== Content of Formulas/Section1.m ===")
                                        print(section1_m_content)
                                        print("=====================================")
                                    elif 'formulas/section1.m' in archive.namelist():
                                        section1_m_content = archive.read('formulas/section1.m').decode('utf-8')
                                        print("\n=== Content of formulas/section1.m ===")
                                        print(section1_m_content)
                                        print("=====================================")
                                    else:
                                        print("未找到Section1.m文件")
                            else:
                                print("ZIP archive start or end signature not found.")
                        else:
                            print("DataMashup content not found.")
                    except etree.XMLSyntaxError as e:
                        # Handle any XML parsing errors
                        print(f"XML parsing error: {e}")
                    except Exception as e:
                        # Handle any other errors
                        print(f"Error: {e}")
    
    # Also check for xl/formulas directory
    print("\n=== 检查 xl/formulas 目录 ===")
    with zipfile.ZipFile(excel_path, 'r') as zip_ref:
        formula_files = [f for f in zip_ref.namelist() if 'formula' in f.lower()]
        print(f"找到的公式相关文件：{formula_files}")
        
        for formula_file in formula_files:
            try:
                content = zip_ref.read(formula_file).decode('utf-8')
                print(f"\n=== Content of {formula_file} ===")
                print(content)
                print("=====================================")
            except Exception as e:
                print(f"读取 {formula_file} 时出错：{e}")

    # Return the list of Power Query files found
    return power_query_files

# 处理命令行参数
if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python power_query_extractor.py <excel_file_path>")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    find_power_query_files(excel_path)
