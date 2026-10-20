import sys
import os
import win32com.client

# 使用 Word 将 .doc 或 .docx 文件转换为 UTF-8 编码的 .txt 文件
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

    # 输出路径：与源文件同目录，同名 .txt
    txt_path = os.path.splitext(doc_path)[0] + '.txt'

    # 启动 Word 应用程序（后台运行）
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False          # 不显示 Word 窗口
    word.DisplayAlerts = 0        # 关闭所有提示对话框

    try:
        # 打开文档
        doc = word.Documents.Open(doc_path, ReadOnly=True)

        # 另存为编码文本，编码指定为 UTF-8 (65001)
        # FileFormat=7 表示 wdFormatEncodedText（编码文本）
        doc.SaveAs2(txt_path, FileFormat=7, Encoding=65001)

        # 关闭文档
        doc.Close(False)
        print(f"转换成功: {doc_path} -> {txt_path}")

    except Exception as e:
        print(f"转换失败: {doc_path}\n错误信息: {e}")

    finally:
        # 退出 Word
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
