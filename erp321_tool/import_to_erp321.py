import json
import re
from os import path

from erp321_tool.browser import close_browser, launch_browser
from erp321_tool.constants import (
    column_name_map,
    get_normalized_name,
    guess_goods_id_map,
    import_shop_select_map,
    price_map,
)
from erp321_tool.utils import (
    dicts_to_workbook,
    get_input_xlsl_file_paths,
    load_xlsx_or_xls_workbook,
    syncify,
    workbook_to_dicts,
)


def parse_xlsx(file_path: str):
    name = get_normalized_name(file_path)
    if not name:
        raise RuntimeError(f"不能识别“{file_path}”这个xlsx文件名")
    workbook = load_xlsx_or_xls_workbook(file_path)
    rows = workbook_to_dicts(workbook)
    workbook.close()
    return rows


def guess_goods_id(row: dict, name: str):
    """
    根据商品名称推测商品编码
    商品名称和商品编码的对应关系在“工具设置/字段匹配.xlsx”中的“商品编码对照”表格
    """
    goods_id = ""
    if name in ("BM", "零元生活"):
        goods_id = guess_goods_id_map[name][f"{row['商品名称']}{row['规格名称']}"]
        goods_id = str(goods_id)
    return goods_id


def check_row_error(row: dict, name: str):
    if not row["商品编码"]:
        raise RuntimeError(
            f"请在工具设置中设置来自{name}的订单中以下商品的商品编码：\n{json.dumps(row, ensure_ascii=False)}"
        )
    if row["商品单价"] is None:
        raise RuntimeError(
            f"请在工具设置中设置来自{name}的订单中以下商品的商品单价：\n{json.dumps(row, ensure_ascii=False)}"
        )


def transform_row(row: dict, file_path: str):
    """
    根据文件“工具设置/字段匹配.xlsx”中的“表名称对应关系”表格，对数据的列进行重命名或转换
    """
    new_row = {
        "线上订单号": "",
        "付款时间": "",
        "卖家备注": "",
        "买家留言": "",
        "收货人姓名": "",
        "手机": "",
        "地址": "",
        "商品编码": "",
        "数量": "",
        "商品单价": "",
        "业务员": "无名氏",
    }

    name = get_normalized_name(file_path)
    for key in column_name_map[name]:
        if key == "业务员":
            new_row[key] = str(column_name_map[name][key])
        elif key in new_row:
            column_name = column_name_map[name][key]

            # 如果列名称是list，说明输出的列信息需要由输入的多列聚合组成
            # 例如有的“地址”列，需要由输入的“省”“市”“区”“地址”这样的多列拼接成
            if isinstance(column_name, list):
                for column_name_item in column_name:
                    new_row[key] += row[column_name_item]
            else:
                new_row[key] = row[column_name]

    new_row["卖家备注"] = (
        new_row["卖家备注"] or f"{name}手工单" if not new_row["卖家备注"] else new_row["卖家备注"]
    )

    # 手机号仅保留数字和“-”，否则聚水潭导单会失败
    new_row["手机"] = re.sub("[^0-9-]", "", str(new_row["手机"]))

    # 有些渠道的推单xlsx文件里面没有包含商品编码，目前有BM、零元生活渠道，需要我们通过商品名称来推测
    # 商品名称和商品编码的对应关系在“工具设置/字段匹配.xlsx”中的“商品编码对照”表格
    new_row["商品编码"] = new_row["商品编码"] or guess_goods_id(row, name)

    if new_row["商品编码"]:
        # 部分渠道文件中的商品编码格式有问题，需要去掉分隔符后面的多余内容，例如“ZH1234-56”实际应为“ZH1234”
        new_row["商品编码"] = str(new_row["商品编码"]).split("-", maxsplit=1)[0]

        if not new_row["商品单价"]:
            new_row["商品单价"] = str(price_map.get(name, {}).get(new_row["商品编码"], 0))

    check_row_error(new_row, name)

    return new_row


def is_normal_row(row):
    """
    有的表格最后会多出一行，例如群接龙表格的最后一行是“合计”，不包含有效信息，需要检测行是否是这种无效行
    | 快递公司 | 快递单号 | 商品名称 | 数量 | 价格 |
    | xxxx     | xxxx    | xxxx    | xxxx |     |
    | 合计     |                                 |
    """
    if row.get("快递公司") == "合计":
        return False
    return True


def transform_rows(rows, file_path):
    return [transform_row(row, file_path) for row in rows if is_normal_row(row)]


def generate_xlsx_file(input_file_path):
    """
    根据传入的路径xlsx文件路径，生成对应的导单文件
    生成的文件可以在聚水潭系统的订单导入系统上传导入
    """
    original_rows = parse_xlsx(input_file_path)
    rows = transform_rows(original_rows, input_file_path)
    workbook = dicts_to_workbook(rows)

    input_file_name = path.basename(input_file_path)
    [input_file_basename, _] = path.splitext(input_file_name)
    output_file_name = f"【导单】{input_file_basename}.xlsx".replace("【推单】", "")
    new_file_dir = path.dirname(input_file_path)
    new_file_path: str = path.join(new_file_dir, output_file_name)

    workbook.save(new_file_path)

    return new_file_path


def generate_xlsx_files():
    """
    根据程序同文件夹下的“【推单】xxx.xlsx”文件，生成对应的“【导单】xxx.xlsx”文件
    生成的文件可以在聚水潭系统的订单导入系统上传导入
    """
    file_paths = get_input_xlsl_file_paths()
    if not file_paths:
        raise RuntimeError("没有找到推单xlsx文件")

    generated_file_paths: list[str] = []

    for file_path in file_paths:
        generated_file_path = generate_xlsx_file(file_path)
        generated_file_paths.append(generated_file_path)

    return generated_file_paths


async def upload_xlsx_file(file_path: str):
    """
    将文件导入聚水潭系统，通过playwright自动打开浏览器进行点击和选择文件操作，最后实际导入需要人工确认
    """
    order_import_url = "https://www.erp321.com/app/order/import/import.aspx"
    name = get_normalized_name(file_path)
    page = await launch_browser(order_import_url)

    await page.locator("#order").set_input_files(file_path)

    shop_name = import_shop_select_map[name]

    if shop_name:
        # 对应导单页面左上角的“请选择店铺”选择框下拉选项，根据文件名匹配到对应店铺进行自动点击
        # 文件名和店铺对应关系在“工具设置/字段匹配.xlsx”中的“表名称对应关系”表格中“聚水潭导单店铺名称”这一行
        shop_option = page.locator(f"._cbb_label:has-text('{shop_name}')")
        await shop_option.dispatch_event("click")

        # 对应页面左边“确认导入订单”按钮
        await page.click("#btnImport")

        # 等待用户操作，导入订单完成或放弃导入后，用户手动关闭浏览器，程序才会继续执行
        await page.wait_for_event("close", timeout=0)

    await close_browser()


async def import_to_erp321():
    xlsx_files = generate_xlsx_files()
    for xlsx_file in xlsx_files:
        await upload_xlsx_file(xlsx_file)


if __name__ == "__main__":
    syncify(import_to_erp321)()
