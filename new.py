import getpass
import streamlit as st
import pandas as pd
import pyodbc
from datetime import datetime



uername = getpass.getuser()
print(uername)
uer_list_A=['a0857869','As','sad1']
uer_list_B=['As','a0857869','123']


def download_invoice():
    upload_file = st.file_uploader("上传发票excel", type="xlsx")
    if upload_file is not None:
        df = pd.read_excel(upload_file)
        header_excel = df.columns.tolist()
        header_excel[27] = '系统DN号'
        df.columns = header_excel
        header_script = ['数电票号码', '销方识别号', '销方名称', '购方识别号', '购买方名称', '开票日期', '税收分类编码',
                         '特定业务类型',
                         '货物或应税劳务名称', '规格型号', '单位', '数量', '单价', '金额', '税率', '税额', '价税合计',
                         '发票来源', '发票票种', '发票状态', '是否正数发票', '发票风险等级', '开票人', '备注',
                         '系统DN号', 'VAT类型']
        header_excel.append('VAT类型')
        for index, row in df.iterrows():  # 遍历每一行数据,添加VAT类型
            # st.write(row)
            # st.write(row['发票票种'])
            if row['发票票种'] == '电子发票（增值税专用发票）':
                df.loc[index, 'VAT类型'] = 'N_VAT'
            else:
                df.loc[index, 'VAT类型'] = 'VAT'
        st.write(df)

def search_db():
    headers = ['ID_HEAD_ID', 'IMPORT_CODE', 'DOCUMENT_NO', 'COMPANY_ID', 'DOCUMENT_DATE', 'DOCUMENT_TYPE',
               'DOC_TYPE_CODE', 'CUST_ID', 'CUST_CODE', 'CUST_NAME_CHN', 'CUST_NAME_ENG', 'TAX_CODE', 'TAX_AMOUNT',
               'BANK_ACCOUNT', 'ADDRESS', 'TAX_CONTROL_TYPE', 'REMARK', 'AMOUNT_CUR_TOTAL', 'AMOUNT_SRC_TOTAL',
               'TOTAL_LINE_NO', 'OPERATOR', 'CHECKER', 'REMITTEE', 'DOC_STATUS', 'DOC_LOCKED', 'GOODS_LIST_NAME',
               'SELLER_NAME', 'SELLER_TAX_CODE', 'SELLER_ADDRESS', 'SELLER_BANK_ACCOUNT', 'INVOICE_CODE',
               'INVOICE_NO',
               'INVOICE_DATE', 'PRIBILLING_REF_DOC', 'PRICE_WAY', 'VOID_FLAG', 'REF01', 'REF02', 'REF03', 'REF04',
               'REF05', 'REF06', 'REF07', 'REF08', 'REF09', 'REF10', 'REF11', 'REF12', 'REF13', 'REF14', 'REF15',
               'REF16', 'REF17', 'REF18', 'REF19', 'REF20', 'UPDATE_ID', 'UPDATE_NAME', 'UPDATE_TIME', 'OUTCODE',
               'EXPRESS_TYPE', 'EXPRESS_NO', 'EXPRESS_DATE', 'DELIVERY_PERSON', 'CONTACT_ID', 'PRINT_OUTPUT_FLAG',
               'PRINTER_MARK', 'EXPRESS_BATCHNUMBER', 'EXPRESS_STATUS_FLAG', 'REAL_ET_AMOUNT_TOTAL',
               'REAL_TAX_AMOUNT',
               'REAL_IT_AMOUNT_TOTAL', 'INVMATCH_STATUS', 'MANUAL_DOC_FLAG', 'BASE_VALUE']
    server = 'apshasqlt221,13163'
    database = 'CNDBSTAXP01'
    username = 'hitpoint'
    password = 'hitpoint@77'
    conn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
    cursor = conn.cursor()

    # 处理搜索逻辑
    def search(keyword):
        cursor.execute("SELECT * FROM dbo.TAX_D_ID_HEAD where REF07 = ? and DOCUMENT_DATE between ? and ?", keyword,
                       start_date_input, end_date_input)
        results = cursor.fetchall()
        return results

    def searchs(keyword, document_no):
        cursor.execute(
            "SELECT * FROM dbo.TAX_D_ID_HEAD where REF07 = ? and (DOCUMENT_DATE between ? and ?) and DOCUMENT_NO = ?",
            keyword,
            start_date_input, end_date_input, document_no)
        result_two = cursor.fetchall()
        return result_two

    def convert_df(df):
        return df.to_excel('C:/Users/a0857869/PycharmProjects/pythonProject1/result1.xlsx', index=False)

    def convert_dfs(df):
        return df.to_excel('C:/Users/a0857869/PycharmProjects/pythonProject1/results1.xlsx', index=False)

    min_date = datetime(1949, 1, 1)
    max_date = datetime.today()
    # 获取开始日期
    start_date_input = st.date_input("请输入查询的开始日期(默认为当天日期)", min_value=min_date, max_value=max_date)
    # 获取结束日期
    end_date_input = st.date_input("请输入查询的结束日期(默认为当天日期)", min_value=min_date, max_value=max_date)
    if start_date_input <= end_date_input:
        st.write('Valid range!', f'Start Date: {start_date_input}, End Date: {end_date_input}')
    else:
        st.write('Invalid range! Please make sure the start date is before or on the end date.')

    keyword = st.text_input('请输入关键字：')

    document_no = st.text_input('请输入document_no：')

    # 添加搜索框
    # 创建 Streamlit 应用 st.title('前端搜索系统')
    if st.button('搜索'):
        if keyword and document_no:  # 搜索两个条件.
            resultss = searchs(keyword, document_no)
            if len(resultss) > 0:
                st.write(f"找到了{len(resultss)} 条相关记录：")
                df_alls = pd.DataFrame()
                for resulta in resultss:
                    dictss = dict(zip(headers, resulta))
                    dfs = pd.json_normalize(dictss)
                    # 合并df
                    df_alls = df_alls._append(dfs, ignore_index=True)
                    df_alls.to_excel('C:/Users/a0857869/PycharmProjects/pythonProject1/result1.xlsx',
                                     sheet_name='Sheet1', index=False)
                convert_dfs(df_alls)
                st.write(df_alls)
                with open('C:/Users/a0857869/PycharmProjects/pythonProject1/results1.xlsx', "rb") as template_files:
                    template_bytes = template_files.read()
                st.download_button(label="Click to Download Template File",
                                   data=template_bytes,
                                   file_name="templates.xlsx",
                                   mime='application/octet-stream')
            else:
                st.write("没有找到相关记录。")
        elif keyword:  # 只搜索一个条件
            results = search(keyword)
            if len(results) > 0:
                st.write(f"找到了{len(results)} 条相关记录：")
                df_all = pd.DataFrame()
                for result in results:
                    dicts = dict(zip(headers, result))
                    df = pd.json_normalize(dicts)
                    # 合并df
                    df_all = df_all._append(df, ignore_index=True)
                    df_all.to_excel('C:/Users/a0857869/PycharmProjects/pythonProject1/result1.xlsx',
                                    sheet_name='Sheet1', index=False)
                convert_df(df_all)
                st.write(df_all)
                with open('C:/Users/a0857869/PycharmProjects/pythonProject1/result1.xlsx', "rb") as template_file:
                    template_byte = template_file.read()
                st.download_button(label="Click to Download Template File",
                                   data=template_byte,
                                   file_name="template.xlsx",
                                   mime='application/octet-stream')
            else:
                st.write("没有找到相关记录。")
        else:
            st.write("没有找到相关记录。")


if uername in uer_list_A and uername in uer_list_B:
    session_state = st.session_state
    session_state['page'] = 'Home'
    # 导航栏
    page = st.sidebar.radio('Navigate', ['search_db', '全电发票更新'])
    if page == 'search_db':
        search_db()
    elif page =='全电发票更新':
        download_invoice()
elif uername in uer_list_B:
    search_db()
elif uername in uer_list_A:
    download_invoice()
else:
    st.write('登录失败，您没有登录权限')