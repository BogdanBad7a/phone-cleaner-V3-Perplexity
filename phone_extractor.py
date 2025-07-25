import re
import pandas as pd
import numpy as np

def convert_scientific_notation(value):
    try:
        if isinstance(value, float) or re.match(r"^\d+(\.\d+)?[eE][+-]?\d+$", str(value)):
            return str(int(float(value)))
    except Exception:
        pass
    return str(value)

def split_ranges(segment):
    base, *ranges = segment.split('/')
    nums = []
    if not ranges:
        return [segment]
    m = re.match(r'(.*?)(\d{1,2})$', base)
    if m:
        prefix = m.group(1)
        headnum = m.group(2)
        for r in ranges:
            r = r.strip(')')
            if len(r) == len(headnum):
                nums.append(prefix + r)
            else:
                nums.append(prefix + r.zfill(len(headnum)))
        nums.insert(0, base)
        return nums
    return [segment]

def clean_and_extract(text):
    numbers = []
    if pd.isna(text):
        return []
    if isinstance(text, float) or isinstance(text, int):
        text = convert_scientific_notation(text)
    text = str(text).strip(',"\' ')
    if not text or text.lower() in ('nan', 'none'):
        return []
    for sep in [',', 'AND', '+', '|', '\n']:
        parts = re.split(rf"\s*{re.escape(sep)}\s*", text)
        if len(parts) > 1:
            subparts = []
            for part in parts:
                if '/' in part:
                    for r in part.split('/'):
                        subparts.append(r)
                else:
                    subparts.append(part)
            text_parts = subparts
            break
    else:
        text_parts = [text]
    final_parts = []
    for seg in text_parts:
        if re.search(r"\d/\d", seg):
            final_parts.extend(split_ranges(seg))
        elif '/' in seg:
            final_parts.extend(seg.split('/'))
        else:
            final_parts.append(seg)
    for item in final_parts:
        s = item.strip(" \t\r\n,;:'\"()[]")
        s = str(s).replace('\t', ' ')
        s = re.sub(r'[\(\)]', '', s)
        s = re.sub(r'ext.*$', '', s, flags=re.I)
        s = re.sub(r'fax.*$', '', s, flags=re.I)
        s = re.sub(r'[\sA-Za-z]+$', '', s)
        s = re.sub(r'[A-Za-np-zA-NP-Z]', '', s)
        s = s.replace('o', '0')
        swap_match = re.match(r"(\d+)[-\s](0\d{1,2}|050)", s)
        if swap_match:
            code, main = swap_match.groups()[1], swap_match.groups()[0]
            s = code.lstrip('0') + main
            if code.startswith('05'):
                s = code[1:] + main
        s = re.sub(r'^(971)-(0[45])-', r'971\2', s)
        s = s.replace('-', '').replace(' ', '')
        s = re.sub(r'971\|00971', '971', s)
        s = re.sub(r'\|', '', s)
        s = re.sub(r'^(00+971|0{2,}971)', '+971', s)
        s = re.sub(r'^(?:\+)?971(?:\+)?971', '+971', s)
        if re.match(r'^05\d{8}$', s):
            s = '+971' + s[1:]
        if re.match(r'^971\d{8,9}$', s):
            s = '+' + s
        if re.match(r'^(4|5)\d{8}$', s):
            s = '+971' + s
        s = '+' + re.sub(r'[^\d]', '', s) if s[0] == '+' else re.sub(r'[^\d]', '', s)
        s = re.sub(r'^\+0?971', '+971', s)
        if re.match(r'^\+9715\d{8}$', s) or re.match(r'^\+971[24]\d{7,8}$', s):
            numbers.append(s)
    return numbers

def extract_uae_phone_numbers(file_path):
    xls = pd.ExcelFile(file_path)
    all_numbers = set()
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet)
        arr = df.values.flatten()
        for val in arr:
            nums = clean_and_extract(val)
            for n in nums:
                all_numbers.add(n)
    return pd.DataFrame({'Phone Number': sorted(all_numbers)})
