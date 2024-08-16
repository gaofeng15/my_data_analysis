# from docxtpl import DocxTemplate
# import pandas as pd
#
# df = pd.read_excel('练习.xlsx')
# for i,row in df.iterrows():
#     xingming = row['姓名']
#     # print(i)
#     # print(dict(row))
#     # print('------------------')
#
#     doc = DocxTemplate('练习.docx')
#     doc.render(dict(row))
#     doc.save(f'./结果/{xingming}.docx')


import streamlit as st,pandas as pd

def zongfen(df):
    df1 = df.groupby('班级').agg(**{'总分': ('成绩', 'sum')})
    return df1
upload_file = st.file_uploader(label='上传要处理的文件')
butt = st.button('显示各班总分')
if upload_file:

    df = pd.read_excel(upload_file)
if butt:
    st.write(zongfen(df))

import pandas as pd
import streamlit as st

