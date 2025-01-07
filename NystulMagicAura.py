import tkinter as tk
from tkinter import filedialog
from docx import Document
import re

def NystulMagicAura(input_file, output_file):
    doc = Document(input_file)
    with open(output_file, 'w', encoding='utf-8') as f:
        i = 0
        first_line = True
        second_line = False
        third_line = False
        forth_line = False
        finish = False
        process = "start"
        for para in doc.paragraphs:
            texts = para.text.splitlines()
            for text in texts:
                text = text.strip()
                if i == 0:  # 名字
                    f.write(f"<p><STRONG><FONT color=#806000>{text}<BR></FONT></STRONG>\n")
                elif i == 1:  # 阵营类型
                    f.write(f"<EM>{text}</EM><BR>\n")
                elif process == "start":
                    keywords = ["AC", "HP", "先攻", "速度"]
                    for kw in keywords:
                        if text.startswith(kw):
                            remainder = text[len(kw):].lstrip()
                            f.write(f"<STRONG>{kw}</STRONG>{remainder}\n")
                            break
                    if text.startswith("速度"):
                        process = "ability"
                        f.write('<p><table width=400 border=0 style="MARGIN-BOTTOM: 5px; BORDER-COLLAPSE: collapse; TEXT-ALIGN: center" cellSpacing=0 cellPadding=2>\n')
                        f.write('<tr style="FONT-SIZE: 10px; COLOR: #333333">\n')
                        f.write('<td colspan=2></td>\n<td>调整</td>\n<td>豁免</td>\n<td colspan=3></td>\n<td>调整</td>\n<td>豁免</td>\n<td colspan=3></td>\n<td>调整</td>\n<td>豁免</td></tr>\n<tr bgColor=#eeeeee>\n')
                elif process == "ability":
                    abilities = text.split()
                    for ability in abilities:
                        tesplit = ability.split('|')
                        name_value, adj1, adj2 = ability.split('|')
                        match = re.match(r"([^\d]+)(\d+)$", name_value)
                        if match:
                            name = match.group(1)   # 匹配到的能力名称（非数字部分）
                            value = match.group(2)  # 匹配到的数字部分
                        else:
                            # 如果没有匹配上，可以做一个简单降级处理
                            name = name_value
                            value = ""
                        if name == "力量" or name == "敏捷":
                            f.write(f"<td><b>{name}</b></td>\n<td>{value}</td>\n<td>{adj1}</td>\n<td>{adj2}</td>\n<td bgcolor=transparent></td>\n")
                        elif name == "体质":
                            f.write(f"<td><b>{name}</b></td>\n<td>{value}</td>\n<td>{adj1}</td>\n<td>{adj2}</td></tr>\n<tr>\n")
                        elif name == "智力" or name == "感知":
                            f.write(f"<td><b>{name}</b></td>\n<td>{value}</td>\n<td>{adj1}</td>\n<td>{adj2}</td>\n<td></td>\n")
                        elif name == "魅力":
                            f.write(f"<td><b>{name}</b></td>\n<td>{value}</td>\n<td>{adj1}</td>\n<td>{adj2}</td></tr></table></p>\n")
                            process = "otherthings"
                            first = True
                elif process == "otherthings":
                    keywords = ["抗性","易伤","免疫","感官","语言","CR"]
                    for kw in keywords:
                        if text.startswith(kw):
                            remain = kw + "："
                            remainder = text[len(remain):].lstrip()
                            if first:
                                f.write(f"<p><STRONG>{kw}</STRONG>{remainder}\n")
                                first = False
                            else:
                                f.write(f"<BR><STRONG>{kw}</STRONG>{remainder}")
                            break
                    if text.startswith("特质"):
                        f.write("<BR><STRONG>特质Traits</STRONG>\n")
                        process = "traits"
                    if text.startswith("动作"):
                        f.write("<BR><STRONG>动作Actions</STRONG>\n")
                        process = "actions"
                elif process == "traits":
                    if text.startswith("动作"):
                        f.write("<BR><STRONG>动作Actions</STRONG>\n")
                        process = "actions"
                    else:
                        if first_line:
                            f.write(f"<BR><STRONG>{text}</STRONG>")
                            first_line = False
                            second_line = True
                        elif second_line:
                            f.write(f"{text}\n")
                            first_line = True
                            second_line = False
                elif process == "actions":
                    if text.startswith("附赠动作"):
                        f.write("<BR><STRONG>附赠动作Bonus Actions</STRONG>\n")
                    if text.startswith("反应"):
                        f.write("<BR><STRONG>反应Reactions</STRONG>\n")
                    else:
                        if forth_line:
                            if "成功：" in text:
                                f.write(f"{text}\n")
                                finish = True
                            else:
                                first_line = True
                        if first_line:
                            f.write(f"<BR><STRONG>{text}</STRONG>")
                            first_line = False
                            second_line = True
                        elif second_line:
                            is_attack_saving = False
                            keywords = ["攻击：","豁免：","触发："]
                            for kw in keywords:
                                if kw in text:
                                    f.write(f"{text}")
                                    second_line = False
                                    third_line = True
                                    is_attack_saving = True
                            if not is_attack_saving:
                                f.write(f"{text}\n")
                                first_line = True
                                second_line = False
                        elif third_line:
                            if "失败：" in text:
                                third_line = False
                                forth_line = True
                            else:
                                f.write(f"{text}\n")
                                first_line = True
                                third_line = False
                        if finish == True:
                            forth_line = False
                            first_line = True      
                i+=1

def main():
    root = tk.Tk()
    root.title("NystulMagicAura")
    input_label = tk.Label(root, text="选择需要转换的docx文件：")
    input_label.pack(pady=5)
    
    input_frame = tk.Frame(root)
    input_frame.pack()
    
    input_entry = tk.Entry(input_frame, width=50)
    input_entry.pack(side=tk.LEFT, padx=5)
    
    def select_input_file():
        file_path = filedialog.askopenfilename(
            title="Select a DOCX file to convert",
            filetypes=[("Word Documents","*.docx")])
        if file_path:
            input_entry.delete(0, tk.END)
            input_entry.insert(0, file_path)

    input_button = tk.Button(input_frame, text="浏览", command=select_input_file)
    input_button.pack(side=tk.LEFT)

    # Label 和输入框 - 用于显示和获取 output_file
    output_label = tk.Label(root, text="输出文件 (TXT)：")
    output_label.pack(pady=5)
    
    output_frame = tk.Frame(root)
    output_frame.pack()
    
    output_entry = tk.Entry(output_frame, width=50)
    output_entry.pack(side=tk.LEFT, padx=5)

    def select_output_file():
        file_path = filedialog.asksaveasfilename(
            title="Save as...",
            defaultextension=".txt",
            filetypes=[("Text Files","*.txt")])
        if file_path:
            output_entry.delete(0, tk.END)
            output_entry.insert(0, file_path)

    output_button = tk.Button(output_frame, text="浏览", command=select_output_file)
    output_button.pack(side=tk.LEFT)
    
    # 提示信息
    status_label = tk.Label(root, text="", fg="blue")
    status_label.pack(pady=5)

    # “开始转换”按钮
    def do_convert():
        input_file = input_entry.get()
        output_file = output_entry.get()
        if not input_file or not output_file:
            status_label.config(text="请先选择输入文件和输出文件！", fg="red")
            return
        
        try:
            NystulMagicAura(input_file, output_file)
            status_label.config(text=f"转换完成！结果已保存到: {output_file}", fg="green")
        except Exception as e:
            status_label.config(text=f"转换失败: {e}", fg="red")

    convert_button = tk.Button(root, text="开始转换", command=do_convert)
    convert_button.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
