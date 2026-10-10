import pandas as pd
import re


def check_phone(phone):
    """校验手机号是否合法"""
    if pd.isna(phone):
        return False, "手机号为空"
    phone_str = str(phone).strip()
    pattern = r'^1[3-9]\d{9}$'
    if re.match(pattern, phone_str):
        return True, "正常"
    else:
        return False, "手机号格式错误"


def clean_waybill(file_name):
    # 读取原始运单表格
    df = pd.read_excel(file_name, sheet_name="运单")
    normal_data = []
    error_data = []

    for index, row in df.iterrows():
        waybill_id = row["运单号"]
        phone = row["手机号"]

        # 判断运单号是否为空
        if pd.isna(waybill_id) or str(waybill_id).strip() == "":
            error_data.append({**row, "异常原因": "运单号为空"})
            continue

        # 判断手机号
        phone_ok, msg = check_phone(phone)
        if not phone_ok:
            error_data.append({**row, "异常原因": msg})
            continue

        # 无异常，加入正常数据集
        normal_data.append(row)

    # 导出清洗完成的正常数据
    df_normal = pd.DataFrame(normal_data)
    df_normal.to_excel("清洗后运单数据.xlsx", index=False)
    # 导出异常数据日志
    df_error = pd.DataFrame(error_data)
    df_error.to_excel("运单异常记录.xlsx", index=False)

    print(f"处理完成！正常数据：{len(normal_data)}条，异常数据：{len(error_data)}条")


if __name__ == "__main__":
    clean_waybill("运单原始数据.xlsx")
