import sys
import os
import win32com.client

# 使用 Word 读取 Unicode 正文，由 Python 写入 UTF-8 编码的 .txt 文件
def convert_doc_to_txt(doc_path):
    # 检查文件是否存在
    if not os.path.exists(doc_path):
        print(f"文件不存在: {doc_path}")
        return

    # 检查扩展名
    ext = os.path.splitext(doc_path)[1].lower()
    if ext not in ('.doc', '.docx'):
        print(f"不支持的文件类型: {ext}，仅支持 .doc 和 .docx")
        return

    doc_path = os.path.abspath(doc_path)

    # 输出路径：与源文件同目录，同名 .txt
    txt_path = os.path.splitext(doc_path)[0] + '.txt'

    word = None
    doc = None
    try:
        # 使用独立的 Word 实例，避免关闭用户正在使用的 Word
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0

        # 打开文档
        doc = word.Documents.Open(doc_path, ReadOnly=True)

        # COM 返回 Unicode 字符串，文本编码由 Python 控制
        text = doc.Content.Text
        text = text.replace('\r\n', '\n').replace('\r', '\n').replace('\x0b', '\n')
        text = text.replace('\x07', '').replace('\ufeff', '').replace('\u00ad', '')

        # UTF-8 BOM 便于 Windows 文本编辑器识别编码
        with open(txt_path, 'w', encoding='utf-8-sig', newline='\n') as f:
            f.write(text)

        print(f"转换成功: {doc_path} -> {txt_path}")

    except Exception as e:
        print(f"转换失败: {doc_path}\n错误信息: {e}")

    finally:
        try:
            if doc is not None:
                doc.Close(False)
        finally:
            if word is not None:
                word.Quit()

def main():
    # 获取拖拽到 exe 上的文件路径列表
    if len(sys.argv) < 2:
        print("请将 .doc 或 .docx 文件拖拽到此程序上，或通过命令行传入文件路径。")
        input("按回车键退出...")
        return

    for path in sys.argv[1:]:
        convert_doc_to_txt(path)

    input("全部处理完成，按回车键退出...")

if __name__ == "__main__":
    main()