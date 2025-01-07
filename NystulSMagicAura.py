from docx import Document
import re

def convert_docx_to_txt(input_file, output_file):
    """
    将 Word (docx) 文档的内容按换行符拆分后写入 TXT 文件。
    """
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
                if i == 0:#名字
                    f.write(f"<p><STRONG><FONT color=#806000>{text}<BR></FONT></STRONG>\n")
                elif i == 1:#阵营类型
                    f.write(f"<EM>{text}</EM><BR>\n")
                #怎么在word格式里AC和HP后有空格先攻和速度没有空格，有点折磨了
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
                        name = name_value[:-2]
                        value = name_value[-2:]
                        match = re.match(r"([^\d]+)(\d+)$", name_value)
                        if match:
                            name = match.group(1)   # 匹配到的能力名称（非数字部分）
                            value = match.group(2) # 匹配到的数字部分
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
                            if is_attack_saving == False:
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


if __name__ == "__main__":
    input_docx = "example.docx"
    output_txt = "output.txt"
    
    convert_docx_to_txt(input_docx, output_txt)
    print(f"已完成转换，结果保存在: {output_txt}")
