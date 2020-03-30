# coding=utf-8
import better_exceptions

better_exceptions.hook()


def get_birth_day(card_id):
    card_id = str(card_id)
    return card_id[6:10] + '年' + card_id[10:12] + '月' + card_id[12:14] + '日'


def get_power(df):
    s = ""
    for c in df.index:
        if c == '集体资产管理权' or c=='集体收益分配权':
            s+='\n'
        if df[c] == '√':
            s += '☑'
        else:
            s += '□'
        s += c
    return s


def get_state(p):
    if p == '户在人在':
        s = '☑户在人在\n□户在人不在□人在户不在□人户均不在□其它______'
    elif p == '户在人不在':
        s = '□户在人在\n☑户在人不在□人在户不在□人户均不在□其它______'
    elif p == '人在户不在':
        s = '□户在人在\n□户在人不在☑人在户不在□人户均不在□其它______'
    elif p == '人户均不在':
        s = '□户在人在\n□户在人不在□人在户不在☑人户均不在□其它______'
    else:
        s = '□户在人在\n□户在人不在□人在户不在□人户均不在☑其它______'

    return s


def show_excel(doc):
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = ""
                for p in cell.paragraphs:  ##如果cell中有多段，即有回车符
                    text += p.text
                print(text, end=' ')
            print()

def print_options(args, parser):
    message = ''
    message += '----------------- Options ---------------\n'
    for k, v in sorted(vars(args).items()):
        comment = ''
        default = parser.get_default(k)
        if v != default:
            comment = '\t[default: %s]' % str(default)
        message += '{:>25}: {:<30}{}\n'.format(str(k), str(v), comment)
    message += '----------------- End -------------------'
    print(message)
# phone_numbers = {
#     '葛晓华': '13979268600',
#     '葛四宝': '15350223863',
#     '葛水龙': '18296295661',
#     '李初水': '13979268600',
#     '葛緑喜': '13968089575',
#     '葛瑞仁': '15180691973',
#     '葛是平': '18370248958',
#     '葛海波': '13715267785',
#     '李爱妹': '13479216872',
#     '葛庚喜': '18397921798',
#     '葛小泉': '18458146582',
#     '葛连员': '15158989152',
#     '葛小波': '15180654630',
#     '吴艳霞': '13576222185',
#     '柳先荣': '15676185179',
#     '葛国喜': '13979295794',
#     '葛俊良': '18017536702',
#     '葛胜钦': '13622325964',
#     '葛明': '13519921025',
#     '葛友龙': '13667026093',
#     '葛三喜': '15979294674',
#     '刘明钗': '15079245746',
#     '葛喜钦': '13576222185',
#     '葛松彬': '',
#     '葛小峰': '13295757898',
#     '葛小初': '18629196881',
#     '葛玉岩': '18146608327',
#     '葛小元': '15057974336',
#
# }

# nb_hu = {
#     '葛晓华': '06025',
#     '葛四宝': '06027',
#     '葛水龙': '06026',
#     '李初水': '06024',
#     '葛緑喜': '06022',
#     '葛瑞仁': '06021',
#     '葛是平': '06018',
#     '葛海波': '06017',
#     '李爱妹': '06016',
#     '葛庚喜': '06014',
#     '葛小泉': '06013',
#     '葛连员': '06002',
#     '葛小波': '06007',
#     '吴艳霞': '12012',
#     '柳先荣': '06023',
#     '葛国喜': '06020',
#     '葛俊良': '06019',
#     '葛胜钦': '06075',
#     '葛明': '06011',
#     '葛友龙': '06010',
#     '葛三喜': '06009',
#     '刘明钗': '06008',
#     '葛喜钦': '06006',
#     '葛松彬': '06005',
#     '葛小峰': '06004',
#     '葛小初': '06003',
#     '葛玉岩': '06001',
#     '葛小元': '06012',
# }
import pandas as pd


def get_huzhu_dict(excel, huzhu_index, name_index):
    huzhu_dict = {}
    huzhu_name = None

    columns = excel.columns.tolist()
    for row in excel.values:
        if row[huzhu_index].__contains__('户主'):
            huzhu_name = row[name_index]
            huzhu_dict[huzhu_name] = pd.DataFrame(columns=excel.columns, dtype=object)
        new_df = pd.DataFrame(data=dict({(columns[i], row[i]) for i in range(len(columns))}), index=[0],
                              columns=excel.columns, dtype=object)
        huzhu_dict[huzhu_name] = pd.concat([huzhu_dict[huzhu_name], new_df], ignore_index=True)

    return huzhu_dict