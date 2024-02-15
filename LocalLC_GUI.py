import streamlit as st
import pandas as pd
import numpy as np
import pickle
import sqlite3
import json
from datetime import datetime
from datetime import timedelta
from typing import TypeVar, Tuple, List, Optional, Dict
from urllib import parse
from bs4 import BeautifulSoup
import ctypes
# 타 모듈에서 사용
import subprocess
import time
import psutil

import sys
sys.path.append('C:\\python_source')
import NERP_PI_LC


def run_as_admin():
    if ctypes.windll.shell32.IsUserAnAdmin():
        return
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)

# 기준정보
file_path_db = 'C:\\python_source\\LocalLC\\' + 'LocalLC.db'
file_path_json = 'C:\\python_source\\LocalLC\\default_opt.json'

locrct_id = {
    '전자문서번호':{'main_id':'IssueIdentifier', 'datatype':str},
    '신용장번호':{'main_id':'DocumentReferenceIdentifier', 'datatype':str},
    '세금계산서번호':{'main_id':'TaxInvoiceIdentifier', 'datatype':str},
    '발급일자':{'main_id':'IssueDate', 'datatype':datetime},
    '인수일자':{'main_id':'AcceptanceDate', 'datatype':datetime}, 
    '금액':{'main_id':'AcceptanceAmount','sub_id':['AmountBasisAmount'], 'datatype':float}, # AmountBasisAmount
    '통화':{'main_id':'AcceptanceAmount','sub_id':['AmountBasisAmount'], 'class':'currency', 'datatype':str},
    '참고사항':{'main_id':'AdditionalConditionsDescriptionText', 'datatype':str}
    }
taxinv_id = {
    '세금계산서번호':{'main_id':'NTSISSUEID', 'datatype':str},
    '발급일자':{'main_id':'TAXDATE', 'datatype':datetime},
    '통화':{'main_id':'WAERK_D', 'datatype':str}, 
    '금액':{'main_id':'TOTAMT_D', 'datatype':float},
    '빌링번호':{'main_id':'', 'datatype':str},
    '신용장번호':{'main_id':'', 'datatype':str}
    #'빌링번호':{'main_id':'VBELN', 'datatype':str},
    }
locadv_id = {
    '전자문서번호':{'main_id':'LocalLetterOfCreditApplicationIdentifier', 'datatype':str}, #
    '개설은행코드':{'main_id':'IssuingBank','sub_id':['Organization','OrganizationIdentifier'], 'datatype':str},
    '개설은행':{'main_id':'IssuingBank','sub_id':['Organization','OrganizationName'], 'datatype':str},
    '신용장번호':{'main_id':'LocalLetterOfCreditIdentifier', 'datatype':str}, #
    '개설의뢰인':{'main_id':'ApplicantParty','sub_id':['Organization','OrganizationName'], 'datatype':str},
    '개설의뢰인대표명':{'main_id':'ApplicantParty','sub_id':['Organization','OrganizationCEOName'], 'datatype':str},
    '개설의뢰인사업자번호':{'main_id':'ApplicantParty','sub_id':['Organization','OrganizationIdentifier'], 'datatype':str},
    '수익자':{'main_id':'BeneficiaryParty','sub_id':['Organization','OrganizationName'], 'datatype':str},
    # '수익자담당자ID':{'main_id':'Contact','sub_id':['ContactEmailAccountText'], 'datatype':str}, # 특정 업체에서만 추가적으로 입력하여 제외
    # '수익자담당자도메인':{'main_id':'Contact','sub_id':['ContactEmailDomainText'], 'datatype':str}, # 특정 업체에서만 추가적으로 입력하여 제외
    '개설일자':{'main_id':'LocalLetterOfCreditIssueDate', 'datatype':datetime}, #
    '통지일자':{'main_id':'NotificationDate', 'datatype':datetime}, #
    '물품인도기일':{'main_id':'DeliveryPromisedDateTime', 'datatype':datetime}, #
    '유효기일':{'main_id':'LocalLetterOfCreditEffectiveDate', 'datatype':datetime}, #
    '서류제시기한':{'main_id':'DocumentPresentationPeriodDate', 'datatype':str}, #  
    'Partial':{'main_id':'TransportPartialShipmentMethodCode', 'datatype':str}, # 9 허용
    '물품명':{'main_id':'SupplyGoodsDescriptionText', 'datatype':str}, # 
    '개설회차':{'main_id':'LocalLetterOfCreditOpenDegreeNumber', 'datatype':str}, # 
    '인수금액':{'main_id':'LocalLetterOfCreditOpenAmount','sub_id':['AmountConvertedAmount'], 'datatype':float}, #
    '인수통화':{'main_id':'LocalLetterOfCreditOpenAmount','sub_id':['AmountConvertedAmount'], 'class':'currency', 'datatype':str}, #
    '제출서류':{'main_id':'RequiredDocuments', 'datatype':str, 'show_tag':True}, #
    '참고사항':{'main_id':'AdditionalInformationDescriptionText', 'datatype':str} #
    }
registeredlc_id = {
    '신용장번호':{'main_id':'wnd[0]/usr/txtZTSDP00130-ZLC_NO', 'sub_id':'text', 'datatype':str},
    '통화':{'main_id':'wnd[0]/usr/ctxtZTSDP00130-ZCURR', 'sub_id':'text', 'datatype':str}, 
    '금액':{'main_id':'wnd[0]/usr/txtZTSDP00130-ZOP_AMT', 'sub_id':'text', 'datatype':float},
    '잔액':{'main_id':'wnd[0]/usr/txtZTSDP00130-ZLC_RAMT', 'sub_id':'text', 'datatype':float},
    'ApplicantCode':{'main_id':'wnd[0]/usr/ctxtZTSDP00130-ZBUYER', 'sub_id':'text', 'datatype':str},
    'PaymentCode':{'main_id':'wnd[0]/usr/ctxtZTSDP00130-ZTERM', 'sub_id':'text', 'datatype':str},
    'PaymentText':{'main_id':'wnd[0]/usr/txtZTSDP00130-ZTERM_DESC', 'sub_id':'text', 'datatype':str},
    'Incoterms':{'main_id':'wnd[0]/usr/ctxtZTSDP00130-ZINCO', 'sub_id':'text', 'datatype':str},
    '개설일자':{'main_id':'wnd[0]/usr/ctxtZTSDP00130-ZOP_DT', 'sub_id':'text', 'datatype':datetime},
    '물품인도기일':{'main_id':'wnd[0]/usr/ctxtZTSDP00130-ZSP_DT', 'sub_id':'text', 'datatype':datetime},
    '유효기일':{'main_id':'wnd[0]/usr/ctxtZTSDP00130-ZVAL_DT', 'sub_id':'text', 'datatype':datetime},  
    'Partial':{'main_id':'wnd[0]/usr/chkZTSDP00130-ZPS_TAG', 'sub_id':'selected', 'datatype':str}
    }
nego_history_id = {
    '빌링번호':{'main_id':'', 'datatype':str},
    '네고일자':{'main_id':'', 'datatype':datetime}
    }
customer_id = {'사업자번호':{'main_id':'', 'datatype':str},
    'Name':{'main_id':'', 'datatype':str},
    'ApplicantCode':{'main_id':'', 'datatype':str},
    'PaymentCode':{'main_id':'', 'datatype':str},
    'Incoterms':{'main_id':'','datatype':str},
    '영업담당자':{'main_id':'', 'datatype':str},
    '영업담당자Knox':{'main_id':'', 'datatype':str} #별도 크롤링하는게 없어서 여기서 관리
    }

convert_table = {'ReceiptTestimonyCopyNumber':'물품수령증 사본',
                 'TaxInvoiceCopyNumber':'세금계산서 사본',
                 'LocalLetterOfCreditCopyNumber':'내국신용장 사본',
                 'OfferSheetCopyNumber':'물품매도확약서 사본'
                 }

def db_open(db_path):
    conn_db = sqlite3.connect(db_path)
    db_cursor = conn_db.cursor()
    return (conn_db, db_cursor)

def db_to_df(conn_db, db_cursor, sql_txt:str)->pd.DataFrame:
    query = db_cursor.execute(sql_txt)
    cols = [column[0] for column in query.description]
    df = pd.DataFrame.from_records(data=query.fetchall(),columns=cols)
    conn_db.close()

    return df

def write_load_pickle(wr_type, file_path, list_for_pickle=None):
    with open(file_path, wr_type) as f:
        if 'w' in wr_type:
            if list_for_pickle is not None:
                pickle.dump(list_for_pickle, f)
            else:
                raise Exception('Dump할 list를 함수에 넣어주세요')
        elif 'r' in wr_type:
            return pickle.load(f)
        else:
            raise Exception('w, wb, r, rb 중 하나로 입력하세요')

def write_load_json(wr_type, file_path, list_object=None):
    with open(file_path, wr_type, encoding='utf-8') as f:
        if 'w' in wr_type:
            if list_object is not None:
                json.dump(list_object, f, indent=2, ensure_ascii=False)
            else:
                raise Exception('Dump할 값을 함수에 넣어주세요')
        elif 'r' in wr_type:
            return json.load(f, strict=False)
        else:
            raise Exception('w, wb, r, rb 중 하나로 입력하세요')
        
def exist_lc_ZSDP10200_C(session, lc_no:str)->bool:
    session.StartTransaction('ZSDP10200_C')
    session.findById("wnd[0]/usr/txtZTSDP00130-ZLC_NO").text = lc_no #"L12G9231000064TEST"
    session.findById("wnd[0]").sendVKey (0)
    if  "cannot be found" in session.findById("wnd[0]/sbar").Text:
        session.findById("wnd[0]").sendVKey (3)
        return False
    else:
        session.findById("wnd[0]").sendVKey (3)
        return True

def crawl_lc_ZSDP10200_C(session, registeredlc_id, lc_no):
    session.StartTransaction('ZSDP10200_C')
    session.findById("wnd[0]/usr/txtZTSDP00130-ZLC_NO").text = lc_no #"L12G9231000064TEST"
    session.findById("wnd[0]").sendVKey (0)

    temp_dict = {}
    for column_name in registeredlc_id.keys():
        temp_dict[column_name] = session.findById(registeredlc_id[column_name]['main_id'])
        if registeredlc_id[column_name]['sub_id'] == 'text':
            temp_dict[column_name] = temp_dict[column_name].text
        elif registeredlc_id[column_name]['sub_id'] == 'selected':
            temp_dict[column_name] = temp_dict[column_name].selected
    session.findById("wnd[0]").sendVKey (3)
    return temp_dict

def update_nerp_lc(registeredlc_id, input_data)->None:
    # 데이터 삽입
    conn_db, db_cursor = db_open(file_path_db)
    sql_query = f'''INSERT OR REPLACE INTO 내국신용장등록내역  
    VALUES (:{', :'.join(registeredlc_id.keys())})
    '''
    db_cursor.execute(sql_query, (input_data)) # 튜플로 넣어야 저장
    conn_db.commit()
    conn_db.close() 

def check_NegoDueDate(row):
    if row['신용장_선적기일'] is None or row['신용장_선적기일'] == '신용장정보X':
        lc_last_ship = None
    else:
        lc_last_ship = datetime.strptime(row['신용장_선적기일'],'%Y.%m.%d').date()

    if row['수령증_발급일자'] is None:
        receipt_issue_date = None
    else:
        receipt_issue_date = datetime.strptime(row['수령증_발급일자'],'%Y-%m-%d').date()
        receipt_issue_date_5_workingday = np.busday_offset(np.datetime64(receipt_issue_date, 'D'), 5).astype(datetime)

    if lc_last_ship is None:
        return '신용장정보X'
    elif receipt_issue_date is None: # 신용장정보는 있지만 수령증은 없음
        return lc_last_ship
    else: # 신용장정보와 수령증이 모두 있음 > 신용장 유효
        return min(lc_last_ship, receipt_issue_date_5_workingday)

def check_progress_localnego(row):
    if row['수령증_발급일자'] is None:
        receipt_issue_date = datetime.strptime('1900-01-01','%Y-%m-%d').date()
    else:
        receipt_issue_date = datetime.strptime(row['수령증_발급일자'],'%Y-%m-%d').date()
    if row['신용장_선적기일'] == '신용장정보X':
        lc_last_ship = datetime.strptime('1900.01.01','%Y.%m.%d').date()
    else:
        lc_last_ship = datetime.strptime(row['신용장_선적기일'],'%Y.%m.%d').date()
    receipt_issue_date_5_workingday = np.busday_offset(np.datetime64(receipt_issue_date, 'D'), 5).astype(datetime)

    if row['네고일자'] is not None:
        return '네고완료'
    if lc_last_ship <= datetime.today().date(): # 1순위)신용장 유효기간
        return '신용장 유효기간 만료'
    elif row['수령증_계산서번호'] is None or row['수령증_계산서번호'] == np.nan: # 2순위)수령증이 없는 케이스
        return '세금계산서가 발행되었으니 10일 이내 물품수령증 발행 필요\n(중소기업이 구매하는 경우는 예외)'
    elif receipt_issue_date > datetime.today().date(): # 3순위)수령증이 있는 케이스
        return '물품수령증 날짜가 오늘 이후이므로 재발행 필요'
    elif receipt_issue_date_5_workingday <= datetime.today().date():
        return '물품수령증 발급일자로부터 5일 경과, 재발급필요'
    else:
        return 'Due date이내에 네고 필요'
    
def chk_and_change_df(df_merged_3table:pd.DataFrame)->pd.DataFrame:
    for column_name in ['신용장_선적기일','신용장_통화']:
        df_merged_3table[column_name].fillna(value='신용장정보X', inplace=True)

    df_merged_3table['통화Chk'] = (df_merged_3table['계산서_통화'] == df_merged_3table['수령증_통화'])
    df_merged_3table['금액Chk'] = (df_merged_3table['계산서_금액'] == df_merged_3table['수령증_금액']) 
    df_merged_3table['계산서Chk'] = df_merged_3table['계산서_계산서번호'] == df_merged_3table['수령증_계산서번호']
    df_merged_3table['인수/발급일Chk'] = (df_merged_3table['수령증_인수일자'] == df_merged_3table['계산서_발급일자'])


    df_merged_3table['NegoDueDate'] = df_merged_3table.apply(check_NegoDueDate, axis=1) 
    df_merged_3table['참고사항'] = df_merged_3table.apply(check_progress_localnego, axis=1) 
    df_merged_3table['수령증발급'] = df_merged_3table['수령증여부'] = np.where(pd.notna(df_merged_3table['수령증_금액']), True, False)

    return df_merged_3table[['참고사항','네고일자','NegoDueDate','수령증발급','빌링번호','신용장번호','신용장_유효기일','수령증_발급일자', '계산서_발급일자', '수령증_인수일자','인수/발급일Chk','신용장_선적기일','계산서_통화','수령증_통화','통화Chk','계산서_금액','수령증_금액','금액Chk','계산서_계산서번호','수령증_계산서번호','계산서Chk','수령증_참고사항']]

def merge_for_locallc_df(df_locallc:pd.DataFrame, df_joined_taxinv_receipt:pd.DataFrame):
    # 신용장번호 별 네고금액 합계 구하기
    summed_amount = df_joined_taxinv_receipt[df_joined_taxinv_receipt['참고사항']=='네고완료'].groupby('신용장번호')['수령증_금액'].sum()
    summed_amount.name = '네고완료_합계'

    merged_local_lc = pd.merge(left=df_locallc, right=summed_amount, on='신용장번호', how='left')
    merged_local_lc['개설금액'] = merged_local_lc['인수금액']
    merged_local_lc['네고완료_합계'] = merged_local_lc['네고완료_합계'].replace(np.nan,0)
    merged_local_lc['네고필요_합계'] =  merged_local_lc['개설금액'].str.replace(',','').astype(float) - merged_local_lc['네고완료_합계']

    merged_local_lc['Partial'] = np.where(pd.notna(merged_local_lc['Partial']), False, True)

    return merged_local_lc[['개설금액','네고완료_합계', '네고필요_합계','신용장번호', '인수통화', '인수금액', '인수잔액', 'ApplicantCode', 'PaymentCode',
       'PaymentText', 'Incoterms', '개설일자', '물품인도기일', '유효기일', 'Partial']]

def input_and_search_xml_ZLLEI09020(session, companyid:str, msgid:str, date:list)->bool:
    '''
        sap에서 메뉴 조작하기위해 사용 
        xmltype은 send/receive 2가지 입력
        날짜는 list안에 입력 ['2023.01.01', '2023.01.31'] or ['2023.01.01']

        조회결과 없으면 (False, 에러메시지) 반환
    '''
    session.StartTransaction('ZLLEI09020')
    
    session.findById("wnd[0]/usr/radR_ACT_D").select() # Transaction Base
    session.findById("wnd[0]/usr/radR_EXW_X").select() # EDI
    session.findById("wnd[0]/usr/radP_ACK_LN").select() # Summary
    session.findById("wnd[0]/usr/ctxtS_BUKRS-LOW").text = companyid # "C100"

    session.findById("wnd[0]/usr/ctxtS_MSGID-LOW").text = msgid #"LOCRCT"
    # session.findById("wnd[0]/usr/txtS_RID-LOW").text = searchid # RECEIVER ID
    # session.findById("wnd[0]/usr/txtS_SID-LOW").text = searchid # SENDER ID

    # 날짜입력
    if type(date) == str: date_start, date_end = date, date
    elif type(date) == list:
        if len(date) == 1: date_start, date_end = date[0], date[0]
        elif len(date) == 2: date_start, date_end = date[0], date[1]
    else: raise ValueError('날짜는 []안에 1개 또는 2개 입력필요')

    session.findById("wnd[0]/usr/ctxtSO_AEDAT-LOW").text = date_start
    session.findById("wnd[0]/usr/ctxtSO_AEDAT-HIGH").text = date_end
    session.findById("wnd[0]").sendVKey(8)

    # 조회결과 없으면 종료
    if  'Message' in session.findById("wnd[0]/sbar").Text:  # == 'Message=>Data not found':
        return False#, session.findById("wnd[0]/sbar").Text)
    elif 'limit is greater' in session.findById("wnd[0]/sbar").Text:
        return False#, session.findById("wnd[0]/sbar").Text)
    elif 'Invalid date' in session.findById("wnd[0]/sbar").Text:
        return False#, session.findById("wnd[0]/sbar").Text)
    
    # 조회 결과에서 NORMAL건 클릭하여 진입
    session.findById("wnd[0]/usr/shell/shellcont[1]/shell").currentCellColumn = "NORMAL"
    session.findById("wnd[0]/usr/shell/shellcont[1]/shell").selectedRows = "0"
    session.findById("wnd[0]/usr/shell/shellcont[1]/shell").doubleClickCurrentCell()
    return True

def loop_get_xml_ZLLEI09020(session, id_list:dict, db_table_name:str, convert_table:dict=None)->bool:
    '''
        조회까지 완료된 상태에서 실행, 전체 xml을 조회하고 dataframe으로 반환
    '''
    for i in range(session.findById("wnd[0]/usr/shell/shellcont[1]/shell").RowCount):
        session.findById("wnd[0]/usr/shell/shellcont[1]/shell").selectedRows = i
        session.findById("wnd[0]/tbar[1]/btn[5]").press()

        # 조회결과 없으면 종료
        if  'Message' in session.findById("wnd[0]/sbar").Text:  # == 'Message=>Data not found
            print(f'(조회불가)', session.findById("wnd[0]/sbar").Text )
            return False
        
        # xml파일경로 확인 및 변환
        file_path = session.findById("wnd[0]/usr/cntlGUI_CONTAINER_X/shellcont/shell").BrowserHandle.LocationURL
        for (before, after) in (('file:///', ''), ('/', '\\'), ('\\', '\\\\')):
            file_path = file_path.replace(before, after)
        file_path = parse.unquote(file_path)

        #  xml 읽고 파싱
        with open(file_path, 'r', encoding='utf-8') as f:
            xml = f.read()
            soup = BeautifulSoup(xml, 'xml')

        temp_row = {}
        for key, tags in id_list.items():
            temp_txt = ''
            if len(tags) == 2: # main_id만 있음
                parsed_txt = soup.findAll(tags['main_id'])
                if len(parsed_txt) == 1:
                    temp_txt = soup.find(tags['main_id']).text
                else:
                    for _ in range(len(parsed_txt)):
                        temp_txt += parsed_txt[_].text + ' '
            elif 'show_tag' in tags.keys():
                for _ in soup.find(tags['main_id']).findAll():
                    temp_txt += f'{_.name}({_.text})' + ' '
            elif 'sub_id' in tags.keys() and 'class' not in tags.keys():
                element_souped = soup.find(tags['main_id'])
                for each_id in tags['sub_id']:
                    element_souped = element_souped.find(each_id)
                temp_txt = element_souped.text
                # temp_txt = soup.find(tags['main_id']).find(tags['sub_id']).text
            elif 'sub_id' in tags.keys() and 'class' in tags.keys():  
                temp_txt = soup.find(tags['main_id']).find(tags['sub_id'])[tags['class']]

            temp_row[key] = temp_txt

        # 파싱된 데이터 타입변환
        for key in temp_row:
            if type(temp_row[key]) == id_list[key]['datatype']:
                continue
            
            if id_list[key]['datatype'] == str:
                temp_row[key] = str(temp_row[key])
            elif id_list[key]['datatype'] == float:
                temp_row[key] = float(temp_row[key])
            elif id_list[key]['datatype'] == datetime:
                temp_row[key] = datetime.strptime(temp_row[key],'%y%m%d').date()
        
        # 파싱된 데이터 내용변환
        if convert_table is not None:
            for name, value in convert_table.items():
                for key in temp_row:
                    if type(temp_row[key]) == str:
                        temp_row[key] = temp_row[key].replace(name, value)

        # 데이터 삽입
        conn_db, db_cursor = db_open(file_path_db)
        sql_query = f'''INSERT OR REPLACE INTO {db_table_name} 
        VALUES (:{', :'.join(id_list.keys())})
        '''
        db_cursor.execute(sql_query, temp_row)
        conn_db.commit()
        conn_db.close()

        session.findById("wnd[0]").sendVKey(3)

def open_nerp_session():
    dict_default_opt = write_load_json('r',file_path_json)
    sap_option = dict_default_opt['sap_option']
    sap_option[3] = int(sap_option[3])

    try:
        sessions = NERP_PI_LC.check_and_open_sap(*sap_option)
        session = sessions[sap_option[3]-1] # sessions[0]
    except:
        run_as_admin()
        sessions = NERP_PI_LC.check_and_open_sap(*sap_option)
        session = sessions[sap_option[3]-1]
    return session

def first_and_end_of_month(year, month)->datetime:
    base_date = datetime.strptime(str(year)+str(month),'%Y%m').date()

    first_day = base_date.replace(day=1)
    next_first_day = (first_day + timedelta(32)).replace(day=1)
    last_day = next_first_day - timedelta(1)

    return (first_day, last_day)


def input_and_search_taxinv_ZRSDM62110(session, payer_list:list, date:str, salesorg:str )->bool:
    '''
        sap에서 메뉴 조작하기위해 사용 
        payer_list: 리스트형의 PayerCode를 입력
        date : YYYYMM 형식 (202311)
        조회결과 없으면 (False, 에러메시지) 반환
    '''
    session.StartTransaction('ZRSDM62110')
    
    session.findById("wnd[0]/usr/ctxtPA_YYMM").text = date #202308
    session.findById("wnd[0]/usr/ctxtPA_VKORG").text = salesorg # R001
    session.findById("wnd[0]/usr/btn%_SO_KUNNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/tbar[0]/btn[16]").press()
    pd.DataFrame(data={'9999999':payer_list}).to_clipboard(index=False)
    session.findById("wnd[1]/tbar[0]/btn[24]").press()
    session.findById("wnd[1]").sendVKey (8)
    session.findById("wnd[0]").sendVKey (8)

    # 조회결과 없으면 종료
    if  'No data' in session.findById("wnd[0]/sbar").Text:  # No data found. Please check your Input data again
        return False#, session.findById("wnd[0]/sbar").Text)
    elif ' Fill out' in session.findById("wnd[0]/sbar").Text: # Fill out all required entry fields
        return False#, session.findById("wnd[0]/sbar").Text)
    else:
        return True

def loop_get_taxinv_ZRSDM62110(session, taxinv_id:dict, cursorandcon)->None:
    temp_row_taxinv = {}
    for result_row in range(session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").RowCount):
        # 행 선택 후 Billing 진입
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = str(result_row)
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellRow = str(result_row)
        session.findById("wnd[0]/tbar[1]/btn[14]").press() # Billing list Disp버튼
        if  'No item has been selected' in session.findById("wnd[0]/sbar").Text:  # No item has been selected
            continue
        elif 'items have been found' in session.findById("wnd[0]/sbar").Text:
            #temp_row_taxinv['빌링번호'] = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(0, taxinv_id['빌링번호']['main_id']) # Billing No
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = "VBELN"
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectedRows = "0"
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").clickCurrentCell()
            session.findById("wnd[0]/tbar[1]/btn[16]").press()
            session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectedRows = "0"
            session.findById("wnd[1]/tbar[0]/btn[2]").press()
            session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").currentCellColumn = "SGTXT"
            session.findById("wnd[0]/usr/cntlCTRL_CONTAINERBSEG/shellcont/shell").doubleClickCurrentCell()
            temp_row_taxinv['빌링번호'] = session.findById("wnd[0]/usr/txtBSEG-ZUONR").text
            temp_row_taxinv['신용장번호'] = session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text
            session.findById("wnd[0]").sendVKey(3)
            session.findById("wnd[0]").sendVKey(3)
            session.findById("wnd[1]").close()
            session.findById("wnd[0]").sendVKey(3)

            session.findById("wnd[0]").sendVKey (3)

        # 정상진입이 된 케이스라면 데이터 수집(안되었다면 Total 또는 Sub Total 행으로 Skip한다)
        temp_row_taxinv['세금계산서번호'] = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(result_row, taxinv_id['세금계산서번호']['main_id'])# 세금계산서 번호
        temp_row_taxinv['발급일자'] = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(result_row, taxinv_id['발급일자']['main_id']) # 세금계산서 발행일자
        temp_row_taxinv['통화'] = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(result_row,taxinv_id['통화']['main_id']) # 통화종류
        temp_row_taxinv['금액'] = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").getCellValue(result_row,taxinv_id['금액']['main_id']) # 발행금액

        # 파싱된 데이터 변환
        for key in temp_row_taxinv:
            if type(temp_row_taxinv[key]) == taxinv_id[key]['datatype']:
                continue
            
            if taxinv_id[key]['datatype'] == str:
                temp_row_taxinv[key] = str(temp_row_taxinv[key])
            elif taxinv_id[key]['datatype'] == float:
                temp_row_taxinv[key] = float(temp_row_taxinv[key].replace(",",""))
            elif taxinv_id[key]['datatype'] == datetime:
                temp_row_taxinv[key] = datetime.strptime(temp_row_taxinv[key],'%Y.%m.%d').date()

        # 데이터 삽입
        conn_db, db_cursor = cursorandcon
        sql_query = f'''INSERT OR REPLACE INTO 세금계산서 
        VALUES (:{', :'.join(taxinv_id.keys())})
        '''
        
        try:
            db_cursor.execute(sql_query, temp_row_taxinv)
        except:
            conn_db, db_cursor = db_open(file_path_db)
            db_cursor.execute(sql_query, temp_row_taxinv)
        
        conn_db.commit()
        conn_db.close()   

def register_localLC(dict_locallc:dict)->None:
    #T코드진입
    session.StartTransaction('ZSDP10200_A')
    session.FindbyId("wnd[0]/usr/ctxtZTSDP00130-ZLCORG").Text = "1"
    session.FindbyId("wnd[0]").sendVKey(0)

    session.FindbyId("wnd[0]/usr/radLLCMARK_03").Select()
    session.FindbyId("wnd[0]/usr/ctxtZTSDP00130-ZBUYER").Text = dict_locallc['ApplicantCode'] #"1157966"
    session.FindbyId("wnd[0]/usr/txtZTSDP00130-ZLC_NO").Text = dict_locallc['신용장번호']#"L12G9230500089"
    session.FindbyId("wnd[0]/usr/ctxtZTSDP00200-LCOUNTRY").Text = dict_locallc['POL(5자리)'][:2]#"KR"
    session.FindbyId("wnd[0]/usr/ctxtZTSDP00200-POL").Text = dict_locallc['POL(5자리)'][2:]#"ICN"
    session.FindbyId("wnd[0]/usr/ctxtZTSDP00200-COUNTRY").Text = dict_locallc['POD(5자리)'][:2]#"KR"
    session.FindbyId("wnd[0]/usr/ctxtZTSDP00200-FINDEST").Text = dict_locallc['POD(5자리)'][2:]#"PUS"
    session.FindbyId("wnd[0]").sendVKey(0)

    session.FindbyId("wnd[0]/usr/txtZTSDP00140-SALESMAN").Text = dict_locallc['영업담당자']#"정은경P"
    session.FindbyId("wnd[0]/usr/txtZTSDP00140-MAIL_ID").Text = dict_locallc['영업담당자Knox']#"ekdms.jeong"
    session.FindbyId("wnd[0]/usr/ctxtZTSDP00130-ZCURR").Text = dict_locallc['인수통화']#"USD"
    session.FindbyId("wnd[0]/usr/txtZTSDP00130-ZOP_AMT").Text = dict_locallc['인수금액']#"293625"
    session.FindbyId("wnd[0]/usr/ctxtZTSDP00130-ZTERM").Text = dict_locallc['PaymentCode']#"CD96"
    session.FindbyId("wnd[0]/usr/ctxtZTSDP00130-ZINCO").Text = dict_locallc['Incoterms']#"FOB"
    session.FindbyId("wnd[0]/usr/ctxtZTSDP00130-ZOP_DT").Text = dict_locallc['개설일자'].replace('-','.')#"2023.05.30"
    session.FindbyId("wnd[0]/usr/ctxtZTSDP00130-ZSP_DT").Text = dict_locallc['물품인도기일'].replace('-','.')#"2023.06.30"
    session.FindbyId("wnd[0]/usr/ctxtZTSDP00130-ZVAL_DT").Text = dict_locallc['유효기일'].replace('-','.')#"2023.06.30"
    session.findById("wnd[0]/usr/chkZTSDP00130-ZPS_TAG").Selected = dict_locallc['Partial']
    session.FindbyId("wnd[0]").sendVKey(0)
    session.FindbyId("wnd[0]").sendVKey(11)
    session.FindbyId("wnd[1]/usr/btnSPOP-OPTION1").press()
    session.FindbyId("wnd[1]/tbar[0]/btn[0]").press()#체크박스 클릭
    session.FindbyId("wnd[1]/usr/btnSPOP-OPTION2").press()#CREATE ITEM NO

    applicant_name = session.FindbyId("wnd[0]/usr/txtZTSDP00200-ZBUY_NM1").Text
    session.FindbyId("wnd[0]/usr/txtZTSDP00200-ZNOTI1").Text = applicant_name
    session.FindbyId("wnd[0]/usr/txtZTSDP00200-ZCONS1").Text = applicant_name
    session.FindbyId("wnd[0]").sendVKey (11)
    session.FindbyId("wnd[1]/usr/btnSPOP-OPTION1").press()   #SAVE YES

   #입력한 포트 갯수따라 반영
    if len(dict_locallc['POL2(5자리)'])!=5 or len(dict_locallc['POD2(5자리)'])!=5:
        session.FindbyId("wnd[1]/usr/btnSPOP-OPTION2").press()   #ADDITIONAL PORT NO
    else:
        session.FindbyId("wnd[1]/usr/btnSPOP-OPTION1").press()   #ADDITIONAL PORT YES
        
        session.FindbyId("wnd[0]/usr/ctxtZTSDP00200-LCOUNTRY").Text = dict_locallc['POL2(5자리)'][:2]#"KR"
        session.FindbyId("wnd[0]/usr/ctxtZTSDP00200-POL").Text = dict_locallc['POL2(5자리)'][2:]#"ICN"
        session.FindbyId("wnd[0]/usr/ctxtZTSDP00200-COUNTRY").Text = dict_locallc['POD2(5자리)'][:2] #"KR"
        session.FindbyId("wnd[0]/usr/ctxtZTSDP00200-FINDEST").Text = dict_locallc['POD2(5자리)'][2:] #'ICN
        
        session.FindbyId("wnd[0]").sendVKey(11) #SAVE
        session.FindbyId("wnd[1]/usr/btnSPOP-OPTION1").press()   #SAVE YES
        


#dict_default_opt = write_load_pickle(wr_type ='w', file_path_db+'default_option.pickle', list_for_pickle=None)
df_receipt = db_to_df(*db_open(file_path_db), 'SELECT * FROM 물품수령증')
df_taxinv = db_to_df(*db_open(file_path_db), 'SELECT * FROM 세금계산서')
df_locallc = db_to_df(*db_open(file_path_db), 'SELECT * FROM 내국신용장')
df_nerplc = db_to_df(*db_open(file_path_db), 'SELECT * FROM 내국신용장등록내역')
df_customer_info = db_to_df(*db_open(file_path_db), 'SELECT * FROM 거래선정보')
df_joined_taxinv_receipt = chk_and_change_df(db_to_df(*db_open(file_path_db), '''
    SELECT             
        물품수령증.세금계산서번호 as 수령증_계산서번호,
        물품수령증.발급일자 as 수령증_발급일자,
        물품수령증.인수일자 as 수령증_인수일자,
        물품수령증.통화 as 수령증_통화,                          
        물품수령증.금액 as 수령증_금액,
        물품수령증.참고사항 as 수령증_참고사항, 
        세금계산서.신용장번호 as 신용장번호, 
        세금계산서.세금계산서번호 as 계산서_계산서번호,
        세금계산서.발급일자 as 계산서_발급일자,                                       
        세금계산서.통화 as 계산서_통화,     
        세금계산서.금액 as 계산서_금액,                     
        세금계산서.빌링번호 as 빌링번호,
        내국신용장등록내역.인수통화 as 신용장_통화,                    
        내국신용장등록내역.물품인도기일 as 신용장_선적기일,  
        내국신용장등록내역.유효기일 as 신용장_유효기일,                    
        내국신용장등록내역.Partial as 신용장_분할선적,                
        빌링번호별네고일자.네고일자 as 네고일자           
    FROM
        세금계산서 AS 세금계산서 
    LEFT JOIN
        물품수령증 AS 물품수령증 ON 세금계산서.세금계산서번호 = 물품수령증.세금계산서번호
    LEFT JOIN
        내국신용장등록내역 AS 내국신용장등록내역 ON 세금계산서.신용장번호 = 내국신용장등록내역.신용장번호
    LEFT JOIN
        빌링번호별네고일자 AS 빌링번호별네고일자 ON 세금계산서.빌링번호 = 빌링번호별네고일자.빌링번호
'''))
df_joined_local_negoamount = merge_for_locallc_df(df_nerplc, df_joined_taxinv_receipt)


dict_default_opt = write_load_json('r',file_path_json)

st.set_page_config(layout="wide")
tab1, tab2, tab3, tab4 = st.tabs(["L/C등록(NERP)", "L/C현황", '세금계산서/물품수령증','옵션설정'])

with tab1:
    tab1_col1, tab1_col2, tab1_col3 = st.columns([3, 0.3, 6])
    with tab1_col1 : # 화면 좌측
        st.title('Local L/C 등록')
        locadv_no = st.text_input(label="전자문서번호를 입력해주세요", value='', max_chars=30, help='30자리 텍스트만 입력가능')#, autocomplete='on')
        df_locallc_for_register = df_locallc[df_locallc['전자문서번호'].str.contains(locadv_no)]

        tab1_col1_col1, tab1_col1_col2 = st.columns([5,5])
        temp_dict_for_register = {}
        with tab1_col1_col1:
            if locadv_no != '':
                column_to_not_show = ['전자문서번호','개설은행코드','개설은행','개설의뢰인대표명','수익자','수익자담당자도메인','참고사항','','']
                for column_name in df_locallc_for_register.columns:
                    # db에 있는 정보 사용
                    if column_name not in column_to_not_show:
                        temp_dict_for_register[column_name] = st.text_input(label=column_name, value=df_locallc_for_register[column_name].values[0])
        with tab1_col1_col2:
            if locadv_no != '':
                df_customer  = df_customer_info[df_customer_info['사업자번호']==temp_dict_for_register['개설의뢰인사업자번호']]
                column_not_in_db = ['ApplicantCode', 'PaymentCode', 'Incoterms', 'POL(5자리)','POD(5자리)', 'POL2(5자리)','POD2(5자리)','영업담당자','영업담당자Knox']
                for column_add in column_not_in_db:
                    # db에 없는 정보 사용1
                    if column_add in ['ApplicantCode', 'PaymentCode', 'Incoterms','영업담당자','영업담당자Knox']:
                        value_box = df_customer[column_add].item()
                    else:
                        value_box = ''
                    temp_dict_for_register[column_add] = st.text_input(label=column_add, value=value_box)
                # #db에 없는 정보 사용2(표시할 필요없이 불러오기만 하면 되는 값)
                # temp_dict_for_register['영업담당자'] = df_customer['영업담당자'].item()
                # temp_dict_for_register['영업담당자Knox'] = df_customer['영업담당자Knox'].item()

                # print(temp_dict_for_register['영업담당자'], temp_dict_for_register['영업담당자Knox'])
                    

    with tab1_col2:
        if st.button('📌등록',): # NERP를 켠 후 등록작업 수행
            if locadv_no != '':
                # pi_name_and_applicantcode = [temp_dict_for_register['신용장번호'],temp_dict_for_register['ApplicantCode']]
                # pi_dates = [temp_dict_for_register['개설일자'],temp_dict_for_register['물품인도기일'],temp_dict_for_register['유효기일']]
                # partial = False if temp_dict_for_register['Partial'] == '9' else True
                # pi_cur_value_payment_inco_incotext_part_trans = [temp_dict_for_register['인수통화'],temp_dict_for_register['인수금액'],
                #                                                     temp_dict_for_register['PaymentCode'],
                #                                                     temp_dict_for_register['Incoterms'],'',partial, False] # Tranship없으므로 항상 True\
                
                temp_dict_for_register['Partial'] = False if temp_dict_for_register['Partial'] == '9' else True

                pi_port_and_addr = []
                proceed_register = True
                for i, each_port in  enumerate([['POL(5자리)','POD(5자리)'], ['POL2(5자리)','POD2(5자리)']]):
                    print()
                    print(i, temp_dict_for_register[each_port[0]], temp_dict_for_register[each_port[1]])
                    print()
                    if len(temp_dict_for_register[each_port[0]])!=5 or len(temp_dict_for_register[each_port[1]]) != 5:
                        if i == 0:
                            st.toast('포트코드를 입력!\n(첫번째 POL/POD 필수, 각 5자리)')
                            proceed_register = False
                            break
                        else:
                            break
                if proceed_register:
                    session = open_nerp_session()
                    register_localLC(temp_dict_for_register)

       


    with tab1_col3:
        st.title('Local L/C 수신내역')
        tab1_col3_col1, tab1_col3_col2, tab1_col3_col3, tab1_col3_col4, tab1_col3_col5 = st.columns([6,1,2,2,1.7])
        with tab1_col3_col1:
            text_box_searchlc = st.text_input(label="검색할 L/C번호를 입력하고 Enter", value='', max_chars=30, help='30자리 숫자만 입력 가능')#, autocomplete='on')
        with tab1_col3_col2:
            pass
        with tab1_col3_col3: 


                        
            # month_to_update_edi = st.date_input("기준일자(해당일자의 연/월로 업데이트)",format='YYYY-MM-DD', value=datetime.strptime('1900-01-01','%Y-%m-%d').date())
            year_to_update_edi = st.text_input(label="연(YYYY)", value=str(datetime.today().year), max_chars=4, help='4자리 숫자만 입력 가능')
        with tab1_col3_col4:
            month_to_update_edi = st.text_input(label="월(MM)", value=str(datetime.today().month), max_chars=2, help='2자리 숫자만 입력 가능')

        with tab1_col3_col5: 
            if st.button('입력한 연/월로 ↓업데이트',):
                session = open_nerp_session()
                
                first_day, last_day = first_and_end_of_month(year_to_update_edi, month_to_update_edi)
                date = [[datetime.strftime(first_day,'%Y.%m.%d'),datetime.strftime(last_day,'%Y.%m.%d')]]
                for each_date in date:
                    if input_and_search_xml_ZLLEI09020(session, companyid='C100', msgid='LOCADV', date=each_date):
                        loop_get_xml_ZLLEI09020(session, id_list=locadv_id, db_table_name='내국신용장', convert_table=convert_table)

        if text_box_searchlc or text_box_searchlc == '':
            df_locallc['Partial'] = df_locallc['Partial'].str.replace('9','허용')
            df_locallc_viewer = df_locallc[df_locallc['신용장번호'].str.contains(text_box_searchlc)]
        st.dataframe(df_locallc_viewer, width=1500, hide_index=True)

with tab2:
    tab2_col1, tab2_col2 = st.columns([8,1.38])
    with tab2_col1:
        st.title('NERP Local L/C 등록내역')
        tab2_col1_col1, tab2_col1_col2 = st.columns([2,8])
        with tab2_col1_col1:
            filter_df_joined_local_negoamount = st.selectbox('전체/잔액O 보기', ['전체', '잔액O'])
        with tab2_col1_col2:
            text_box_searchNERPlc = st.text_input(label="검색할 L/C번호를 입력하고 Enter ", value='', max_chars=30, help='30자리 숫자만 입력 가능')#, autocomplete='on')
        if text_box_searchNERPlc or text_box_searchNERPlc == '':
            df_nerplc_viewer = df_joined_local_negoamount[df_joined_local_negoamount['신용장번호'].str.contains(text_box_searchNERPlc)].sort_values('유효기일')
        if filter_df_joined_local_negoamount == '잔액O':
            df_nerplc_viewer = df_nerplc_viewer[df_nerplc_viewer['네고필요_합계']>0]
        st.dataframe(df_nerplc_viewer, width=1500, hide_index=True)



    with tab2_col2:
        st.title('수동Update')
        text_box_lc = st.text_input(label="추가할 L/C번호를 입력하세요", value='', max_chars=30, help='30자리 숫자만 입력 가능')#, autocomplete='on')   
        if st.button('입력한 Local L/C 추가/업데이트'):
            session = open_nerp_session()
            if exist_lc_ZSDP10200_C(session,text_box_lc):
                input_data = crawl_lc_ZSDP10200_C(session, registeredlc_id, text_box_lc)
                update_nerp_lc(registeredlc_id, input_data) 

        if st.button('Local L/C EDI 수신 내역으로 \n불러오기'):
            session = open_nerp_session()
            need_chk_all = set(db_to_df(*db_open(file_path_db), 'SELECT * FROM 내국신용장')['신용장번호'].tolist())
            need_not_chk = set(db_to_df(*db_open(file_path_db), 'SELECT * FROM 내국신용장등록내역')['신용장번호'].tolist())
            need_chk_final =need_chk_all - need_not_chk

            for lc in need_chk_final:
                if exist_lc_ZSDP10200_C(session,lc):
                    input_data = crawl_lc_ZSDP10200_C(session, registeredlc_id, lc)
                    update_nerp_lc(registeredlc_id, input_data) 

        if st.button('새로고침'):
            pass


with tab3:
    tab3_col1, tab3_col2 = st.columns([8,1.38])
    with tab3_col1:
        st.title('세금계산서/물품수령증')
        tab3_col1_col1, tab3_col1_col2, tab3_col1_col3, tab3_col1_col4, tab3_col1_col5, tab3_col1_col6 = st.columns([6,1,2,2,1.5,1.5])
        with tab3_col1_col1:
            text_box_search_taxinv = st.text_input(label="검색할 세금계산서번호를 입력하세요", value='', max_chars=30, help='30자리 숫자만 입력 가능')#, autocomplete='on')
            if text_box_search_taxinv or text_box_search_taxinv == '':
                df_taxinv_viewer = df_joined_taxinv_receipt[df_taxinv['세금계산서번호'].str.contains(text_box_search_taxinv)]
        with tab3_col1_col3:
            year_to_update_tax_receipt = st.text_input(label="연(YYYY) ", value=str(datetime.today().year), max_chars=4, help='4자리 숫자만 입력 가능')
        with tab3_col1_col4:
            month_to_update_tax_receipt = st.text_input(label="월(MM) ", value=str(datetime.today().month), max_chars=2, help='2자리 숫자만 입력 가능')
            if len(month_to_update_tax_receipt) == 1:
                month_to_update_tax_receipt = '0' + month_to_update_tax_receipt
        with tab3_col1_col5:
            if st.button('입력한 연/월로 세금계산서'):
                date = [year_to_update_tax_receipt + month_to_update_tax_receipt]
                #payer_list = [info["ApplicantCode"] for info in dict_default_opt["applicant_info"].values()]
                payer_list = df_customer_info['ApplicantCode'].tolist()

                session = open_nerp_session()
                for each_date in date:
                    if input_and_search_taxinv_ZRSDM62110(session, payer_list=payer_list, date=each_date, salesorg='R001'):
                        loop_get_taxinv_ZRSDM62110(session, taxinv_id, db_open(file_path_db))

        with tab3_col1_col6:
            if st.button('입력한 연/월로 물품수령증',):
                first_day, last_day = first_and_end_of_month(year_to_update_tax_receipt, month_to_update_tax_receipt)
                date = [[datetime.strftime(first_day,'%Y.%m.%d'),datetime.strftime(last_day,'%Y.%m.%d')]]

                session = open_nerp_session()
                for each_date in date:
                    if input_and_search_xml_ZLLEI09020(session, companyid='C100', msgid='LOCRCT', date=each_date):
                        loop_get_xml_ZLLEI09020(session, id_list=locrct_id, db_table_name='물품수령증')



        st.dataframe(df_taxinv_viewer, width=1500, hide_index=True)

    with tab3_col2:
        st.title('네고일 등록')
        date_nego = st.date_input("네고일자 선택",format='YYYY-MM-DD')
        text_box_billingno = st.text_input(label="빌링번호 입력", value='', max_chars=12, help='12자리 숫자만 입력 가능')#, autocomplete='on')
        if st.button('📌등록 '):
            input_data = {'빌링번호':text_box_billingno, '네고일자':date_nego}
            conn_db, db_cursor = db_open(file_path_db)
            sql_query = f'''INSERT OR REPLACE INTO 빌링번호별네고일자  
            VALUES (:{', :'.join(nego_history_id.keys())})
            '''
            db_cursor.execute(sql_query, (input_data)) # 튜플로 넣어야 저장
            conn_db.commit()
            conn_db.close() 
        if st.button('🛒삭제 '):
            conn_db, db_cursor = db_open(file_path_db)
            sql_query = f'''DELETE FROM 빌링번호별네고일자 WHERE 빌링번호 == {text_box_billingno}'''
            db_cursor.execute(sql_query) # 튜플로 넣어야 저장
            conn_db.commit()
            conn_db.close()
        if st.button('새로고침 '):
            pass
        if st.button('수령증 미발급건 클립보드 복사'):
            df_taxinv_viewer[df_taxinv_viewer['수령증발급']==False][['빌링번호','신용장번호','신용장_유효기일','계산서_발급일자','계산서_계산서번호']].to_clipboard(sep='\t', index=False)


with tab4:
    tab4_col1, tab4_col2, tab4_col3 = st.columns([2,4,3])
    with tab4_col1:
        st.title('로그인정보')
        idx_txtbox = {0:'SEP/DEP',1:'ID',2:'PW',3:'n번째 NERP창 사용(최대6)'}
        data_txtbox = {}
        for i, each_data in enumerate(dict_default_opt['sap_option']):
            data_txtbox[i] = st.text_input(label=idx_txtbox[i], value=dict_default_opt['sap_option'][i], type=('password' if i == 2 else 'default'))
        if st.button('저장'):
            dict_default_opt['sap_option'] = [data_txtbox[0], data_txtbox[1], data_txtbox[2], data_txtbox[3]]
            write_load_json('w', file_path_json, list_object=dict_default_opt)

    with tab4_col2:
        st.title('거래선정보')
        
        st.dataframe(df_customer_info, hide_index=True)

        # 거래선정보를 dict로 관리할때만 사용
        # for each_data in dict_default_opt['applicant_info']:
        #     st.write(each_data, str(dict_default_opt['applicant_info'][each_data]))
    with tab4_col3:
        st.title('거래선정보 수정')
        tab4_col3_col1, tab4_col3_col2 = st.columns([9,1])
        with tab4_col3_col1:
            text_box_customer = st.text_input(label=':red[수정할 거래선의 사업자번호 입력하세요]', value='', max_chars=30, help='')#, autocomplete='on')
            temp_dict_for_customer={}
            if text_box_customer == '':
                iterater = df_customer_info
                for column_name in iterater.columns:
                    temp_dict_for_customer[column_name] = st.text_input(label=f'{column_name}_Customer', value='')
            else:
                try:
                    iterater = df_customer_info[df_customer_info['사업자번호'].str.contains(text_box_customer)]
                    for column_name in iterater.columns:
                        temp_dict_for_customer[column_name] = st.text_input(label=f'{column_name}_Customer', value=iterater[column_name].iloc[0])
                except:
                    iterater = df_customer_info
                    for column_name in iterater.columns:
                        temp_dict_for_customer[column_name] = st.text_input(label=f'{column_name}_Customer', value='')

        with tab4_col3_col2:
            if st.button('추가/수정 '):
                conn_db, db_cursor = db_open(file_path_db)
                sql_query = f"INSERT OR REPLACE INTO 거래선정보 VALUES (:{', :'.join(customer_id.keys())})"
                db_cursor.execute(sql_query, temp_dict_for_customer)
                conn_db.commit()
                conn_db.close()
            if st.button('삭제 '):
                conn_db, db_cursor = db_open(file_path_db)
                db_cursor.execute(f'DELETE FROM 거래선정보 WHERE 사업자번호 = "{text_box_customer}"')
                conn_db.commit()
                conn_db.close()
            if st.button('새로고침  '):
                pass


                    
            

        
# streamlit run LocalLC_GUI.py
# streamlit run .\LocalLC\LocalLC_GUI.py
