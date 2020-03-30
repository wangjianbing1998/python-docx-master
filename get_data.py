# coding=gbk
import pandas as pd

pd.set_option('expand_frame_repr', False)
pd.set_option('display.max_rows', 20)
pd.set_option('precision', 2)


def get_huzhu(excel, huzhu_index, name_index):
    huzhu_dict = {}
    huzhu_name = None

    columns = excel.columns.tolist()
    for row in excel.values:
        if row[huzhu_index] == row[name_index]:
            huzhu_name = row[name_index]
            huzhu_dict[huzhu_name] = pd.DataFrame(columns=excel.columns, dtype=object)
        new_df = pd.DataFrame(data=dict({(columns[i], row[i]) for i in range(len(columns))}), index=[0],
                              columns=excel.columns, dtype=object)
        huzhu_dict[huzhu_name] = pd.concat([huzhu_dict[huzhu_name], new_df], ignore_index=True)

    return huzhu_dict


def get_df(huzhu_dict,columns):
    df=pd.DataFrame(columns=columns,dtype=object)

    for huzhu_name, excel in huzhu_dict.items():
        df=df.append(excel,ignore_index=True)

    return df



def get_df_huzhu(zuhao):
    df = pd.read_excel(f"C:/Users/25536/Desktop/六组/所需手工导入数据/{zuhao}.xlsx", dtype=object)
    df.fillna('', inplace=True)
    df=df.applymap(lambda x:str(x).strip())
    huzhu_dict = get_huzhu(df, huzhu_index=0, name_index=1)

    for huzhu_name, excel in huzhu_dict.items():
        excel.loc[excel['户籍号'] == '', '户籍号'] = excel.ix[0, '户籍号']
        excel.loc[excel['户序号'] == '', '户序号'] = excel.ix[0, '户序号']
        excel.loc[excel['存在状态'] == '', '存在状态'] = excel.ix[0, '存在状态']
        excel.loc[excel['户籍地址'] == '', '户籍地址'] = '江西湖口'
        excel.loc[excel['现住地址'] == '', '现住地址'] = f'石山村{zuhao}'

    df=get_df(huzhu_dict,df.columns)


    return df,huzhu_dict


if __name__ == '__main__':
    df,huzhu_dict=get_df_huzhu('十三组')

