import os
from io import BytesIO
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv()
APPLICATION_ID = os.getenv("APPLICATION_ID")

FILE_PATH = Path(__file__).parent / "法人番号API.xlsx"


def get_corporation_list(file_path):
    """
    法人番号APIから法人番号を取得するためのリストを作成する

    :param file_path: 法人番号API.xlsxのファイルパス
    :return:法人情報のリスト
    """
    # 取引先一覧シートをデータフレームとして取得
    df_info = pd.read_excel(file_path, sheet_name='取引先一覧', dtype=str)

    # 住所の郵便番号のデータフレームを取得
    url = 'https://www.post.japanpost.jp/zipcode/dl/oogaki/zip/ken_all.zip'
    df_ken_all = get_city_code_df(url, 2)

    # 事業所の個別郵便番号のデータフレームを取得
    url = 'https://www.post.japanpost.jp/zipcode/dl/jigyosyo/zip/jigyosyo.zip'
    df_jigyosyo = get_city_code_df(url, 7)

    # 住所の郵便番号と事業所の個別郵便番号のデータフレームを結合　pd.concat([df_ken_all, df_jigyosyo]
    # 取引先一覧と結合したデータフレームを郵便番号をキーとしてマージ
    df_info = pd.merge(df_info, pd.concat([df_ken_all, df_jigyosyo]), on='郵便番号', how='left')

    # 市区町村コードが存在しない行があるときはエラー.csvに出力
    if len(df_info[df_info['市区町村コード'].isnull()]):
        df_info[df_info['市区町村コード'].isnull()].to_csv('エラー.csv', index=False, encoding='cp932')
        print('市区町村コードが存在しない法人がありました。エラー.csvを確認してください。')

    # 市区町村コードが存在しない行を削除
    df_info = df_info.dropna(subset=['市区町村コード'])

    # 法人情報のリストを作成、戻り値として返す
    return df_info.to_dict('records')


def get_city_code_df(url, col):
    """
    市区町村コード・郵便番号のデータフレームを取得する

    :param url:リクエストURL
    :param col:郵便番号の列番号
    :return:市区町村コード・郵便番号のデータフレーム
    """
    response = requests.get(url)
    if response.status_code != 200:
        print({response.status_code})
        exit(1)

    # 市区町村コード・郵便番号のみをデータフレームとして取得
    return pd.read_csv(BytesIO(response.content), compression='zip', header=None, encoding='cp932',
                       usecols=[0, col], names=['市区町村コード', '郵便番号'], dtype=str)


def get_corporation_info_from_api(corporation_list, application_id):
    """
    法人番号APIから法人番号等を取得して新たにリストを作成する

    :param corporation_list: 法人情報のリスト
    :param application_id: アプリケーションID
    :return: 新たに作成した法人情報のリスト
    """
    result_list = []
    base_url = 'https://api.houjin-bangou.nta.go.jp/4/name?'

    for corporation in corporation_list:
        param_dict = {
            'id': application_id, 'type': 12, 'history': 0,
            'name': corporation['法人名'], 'address': corporation['市区町村コード'],
        }

        response = requests.get(base_url, params=param_dict)
        if response.status_code != 200:
            print(f'法人番号APIがエラーです:{response.status_code}')
            exit(1)

        soup = BeautifulSoup(response.text, 'xml')
        for result in soup.find_all('corporation'):
            result_list.append({
                'ID': corporation['ID'],
                '法人番号': result.find('corporateNumber').text,
                '法人名': result.find('name').text,
                '住所': result.find('prefectureName').text + result.find('cityName').text + result.find(
                    'streetNumber').text,
            })
    return result_list


def get_invoice_list_from_api(corporation_list, application_id):
    """
    法人情報リストにインボイスの登録番号と登録年月日を追加する

    法人情報リストの法人番号を利用してインボイスAPIから登録番号と登録年月日を取得してリストに追加する
    :param corporation_list:法人情報のリスト
    :param application_id:アプリケーションID
    :return:法人情報のリスト
    """
    base_url = 'https://web-api.invoice-kohyo.nta.go.jp/1/num?'

    for corporation in corporation_list:
        param_dict = {'id': application_id, 'number': 'T' + corporation['法人番号'], 'type': 21, }

        response = requests.get(base_url, params=param_dict)
        if response.status_code != 200:
            print(f'インボイスAPIがエラーです:{response.status_code}')
            exit(1)

        if int(response.json()['count']):
            corporation['登録番号'] = response.json()['announcement'][0]['registratedNumber']
            corporation['登録年月日'] = response.json()['announcement'][0]['registrationDate']

    return corporation_list


if __name__ == '__main__':
    corporation_list = get_corporation_list(FILE_PATH)
    corporation_list = get_corporation_info_from_api(corporation_list, APPLICATION_ID)
    corporation_list = get_invoice_list_from_api(corporation_list, APPLICATION_ID)

    # 法人情報のリストをデータフレームに変換してインボイス.xlsxに出力
    df = pd.DataFrame(corporation_list)
    df.to_excel('インボイス.xlsx', index=False)
    print('完了しました')
