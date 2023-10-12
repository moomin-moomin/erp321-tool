from os import chdir, path
from sys import executable, platform
from typing import cast

from erp321_tool.utils import (
    load_xlsx_or_xls_workbook,
    sheet_to_dict,
    workbook_to_dicts,
)

if platform == "darwin":
    mac_app_inner_path = "聚水潭导单回单工具.app/Contents/MacOS/聚水潭导单回单工具"
    mac_bin_inner_path = "聚水潭导单回单工具/聚水潭导单回单工具"
    if executable.endswith(mac_app_inner_path):
        app_path = executable.removesuffix(mac_app_inner_path)
        chdir(app_path)
    elif executable.endswith(mac_bin_inner_path):
        app_path = executable.removesuffix(mac_bin_inner_path)
        chdir(app_path)

option_workbook = load_xlsx_or_xls_workbook("工具设置/字段匹配.xlsx")
price_workbook = load_xlsx_or_xls_workbook("工具设置/渠道商品价格.xlsx")


_full_column_name_map = sheet_to_dict(option_workbook["表名称对应关系"])

column_name_map = {
    shop_name: {key: value for (key, value) in shop_column_name_map.items()}
    for (shop_name, shop_column_name_map) in _full_column_name_map.items()
}


def get_normalized_name(name: str):
    """
    获取标准渠道名称方便统一使用
    标准名称以“工具设置/字段匹配.xlsx”中“表名称对应关系”表格中的列名称为准
    目前可能的名称为：
    蜜淘、童品会、BM、群接龙、快团团、任冉、零元生活、qtools
    后续可能会根据实际情况增减
    """
    keyword_map = {key: key for key in _full_column_name_map}
    file_basename = path.basename(name)
    for keyword, normalized_name in keyword_map.items():
        if keyword in file_basename:
            return normalized_name
    return ""


def _is_valid_price_row(row: dict):
    for column in ["合作渠道", "商品编码", "供货价"]:
        if row[column] is None:
            return False
    return True


# 价格对照表，对应“工具设置/字段匹配.xlsx”中的内容
# 参考格式如下：
# {
#     "蜜淘": {"ZHZ0447": "104", "ZHZ1250": "59.92", "ZHZ1251": "79.92"},
#     "童品会": {"ZHZ0371": "80", "ZHZ0389": "118", "ZHZ1286": "79.9"},
#     "BM": {"ZHZ1169": "57.9", "ZHZ1250": "59.92", "ZHZ1251": "79.92"},
#     "任冉": {"ZHZ0540": "104", "ZHZ0542": "79.9"},
# }
price_map: dict[str, dict[str, str]] = {}


def update_price_map():
    price_dicts = [
        row for row in workbook_to_dicts(price_workbook) if _is_valid_price_row(row)
    ]
    for price_dict in price_dicts:
        channel = cast(str, price_dict["合作渠道"])
        shop = get_normalized_name(channel)
        price_map[shop] = price_map.get(shop, {})
        goods_id = str(price_dict["商品编码"])
        goods_price = str(price_dict["供货价"])
        price_map[shop][goods_id] = goods_price


update_price_map()

import_shop_select_map = {
    shop_name: shop_column_name_map["聚水潭导单店铺名称"]
    for (shop_name, shop_column_name_map) in _full_column_name_map.items()
}

guess_goods_id_map = sheet_to_dict(option_workbook["商品编码对照"])

possible_express_company_column_names = {"物流公司", "快递公司", "物流公司（必填）"}
possible_express_number_column_names = {"运单号", "物流单号", "快递单号", "物流单号（必填）"}


option_workbook.close()
price_workbook.close()
