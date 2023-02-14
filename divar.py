import json
import string
import os.path
import urllib3
import requests
import xlsxwriter
from requests.exceptions import ConnectTimeout, ReadTimeout, ConnectionError


def _request(url):
    while True:
        # Request
        print(f'Requesting for {url}')
        try:
            response = requests.get(url, timeout=20)
            res = response.json()
        except (ConnectTimeout, ReadTimeout, ConnectionError, urllib3.exceptions.ReadTimeoutError):
            continue

        return res


def _write_file(name: str, data: list):
    file_name = f'{name}.json'

    # Create File If Not Exists
    if not os.path.exists(file_name):
        open(file_name, 'w').close()

    # Read Old File
    with open(file_name) as file:
        try:
            list_obj = json.load(file)
        except json.decoder.JSONDecodeError:
            list_obj = list()  # Initial

    # Append New Data
    list_obj.extend(data)

    # Write New File
    with open(file_name, 'w') as file:
        json.dump(list_obj, file,
                  indent=4,
                  separators=(',', ': '))


def label_to_int(label: str) -> int:
    """
    :param label: ۱۷ آگهی
    :return: 17
    """
    label = (
        label[:-5]
        .replace('\u06f0', '0')
        .replace('\u06f1', '1')
        .replace('\u06f2', '2')
        .replace('\u06f3', '3')
        .replace('\u06f4', '4')
        .replace('\u06f5', '5')
        .replace('\u06f6', '6')
        .replace('\u06f7', '7')
        .replace('\u06f8', '8')
        .replace('\u06f9', '9')
    )

    return int(label)


def create_excel(rows: list[dict], name: str) -> str:
    file_name = f'{name}.xlsx'

    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})

    header = rows[0].keys()
    worksheet.set_column('A:A', 5)
    for i, h in enumerate(header):
        char = string.ascii_uppercase[i + 1]
        worksheet.set_column(f'{char}:{char}', 15)
        worksheet.write(f'{char}1', h, bold)

    row_number = 1
    for row in rows:
        column_number = 0
        worksheet.write(row_number, column_number, row_number)
        column_number += 1
        for _, v in row.items():
            worksheet.write(row_number, column_number, v)
            column_number += 1
        row_number += 1
    workbook.close()
    return file_name


def collect_stores(category: str, last_item_id: str = ''):
    has_next = True

    while has_next:
        # Request
        url = f'https://api.divar.ir/v8/marketplace/stores-list/tehran/{category}?last_item_identifier={last_item_id}'
        res = _request(url)

        data = res.get('widget_list', [])
        _write_file(name=category, data=data)

        # Continue If It Has Next
        has_next = res['infinite_scroll_response'].get('has_next', False)
        last_item_id = res['infinite_scroll_response']['last_item_identifier']


def collect_store_products(store_slug: str):
    has_next = True
    last_item_id = ''

    while has_next:
        # Request
        url = f'https://api.divar.ir/v8/marketplace/w/landing2/{store_slug}?last_item_identifier={last_item_id}'
        res = _request(url)

        data = res.get('widget_list', [])
        _write_file(name=store_slug, data=data)

        # Continue If It Has Next
        has_next = res['infinite_scroll_response']['has_next']
        last_item_id = res['infinite_scroll_response']['last_item_identifier']


def collect_store_contact(store_slug: str) -> dict:
    """
    phone_number	"09121491846"
    is_good_time	true
    """
    # Request
    url = f'https://api.divar.ir/v8/marketplace/{store_slug}/contact'
    return _request(url)


def clean_stores_before_excel(category: str, data: dict):
    _write_file(name=f'{category}-cleaned', data=[data])


def task_1(categories):
    """Collect Data"""
    for c in categories:
        collect_stores(c)


def task_2(categories):
    """Create Data For Excel"""
    for c in categories:
        all_contents = list()
        with open(f'{c}.json') as file:
            contents = json.loads(file.read())
            for content in contents:
                if content['widget_type'] == 'EVENT_ROW':
                    content_data = content.get('data')
                    slug = content_data.get('action', {}).get('payload', {}).get('slug', '')

                    if slug:
                        contact = collect_store_contact(slug)
                        phone_number = contact.get('contact', {}).get('phone_number', '')
                    else:
                        phone_number = ''

                    data = {
                        'title': content_data.get('title', ''),
                        'slug': slug,
                        'subtitle': content_data.get('subtitle', ''),
                        'phone_number': phone_number,
                        'image_url': content_data.get('image_url', ''),
                        'label': content_data.get('label', ''),
                    }
                    clean_stores_before_excel(c, data)
            all_contents.append(data)


def task_3(categories):
    """Create Excel"""
    for c in categories:
        with open(f'{c}-cleaned.json') as file:
            contents = json.loads(file.read())
            contents = sorted(contents, key=lambda x: label_to_int(x['label']), reverse=True)
            create_excel(contents, name=c)


if __name__ == '__main__':
    _categories = ['electronic-devices', 'personal', 'home-kitchen']

    # task_1(categories)  # Done
    # task_2(_categories)  # Done
    task_3(_categories)  # Done
