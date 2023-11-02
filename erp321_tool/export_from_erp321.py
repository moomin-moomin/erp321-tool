import json
from os import path

from erp321_tool.browser import close_browser, launch_browser
from erp321_tool.constants import (
    column_name_map,
    get_normalized_name,
    possible_express_company_column_names,
    possible_express_number_column_names,
)
from erp321_tool.utils import (
    get_input_xlsl_file_paths,
    load_xlsx_or_xls_workbook,
    syncify,
    workbook_to_dicts,
)


def get_order_logistics_info(order):
    # 被拆单的数据行，没有实际物流数据
    if order["src_status"] == "Split":
        return None

    express_company = order["logistics_company"]
    express_number = order["l_id"]

    return {
        "order_id": order["so_id"],
        "express_company": express_company,
        "express_number": express_number,
    }


async def get_orders_logistics_info(order_ids: list[str]):
    """
    从聚水潭系统中导出订单的物流信息
    通过playwright自动打开浏览器，自动在网页输入要查询的订单号，然后解析拦截到的浏览器请求接口内容，返回订单相关的物流信息
    """
    order_list_url = "https://www.erp321.com/app/order/order/list.aspx"
    page = await launch_browser(order_list_url)

    page_size = page.locator("#_jt_page_size")
    page_size_value = await page_size.input_value()
    if not page_size_value == "500":
        await page_size.select_option("500")
    await page.wait_for_load_state("networkidle")

    # 在页面左侧的“线上订单号”输入框中输入订单号，然后按回车键查询
    order_id_input = page.frame_locator("#s_filter_frame").locator("#so_id")
    await order_id_input.focus()
    await order_id_input.fill(",".join(order_ids))
    await order_id_input.press("Enter")

    # 拦截读取订单接口返回的信息并进行解析，接口返回不是标准的JSON，需要处理一下才能得到结果
    async with page.expect_response(
        lambda response: response.url.startswith(order_list_url)
    ) as response_info:
        order_api_response = await response_info.value
    response_text = await order_api_response.text()
    response_json_text = response_text[2:]
    response_json = json.loads(response_json_text)
    return_value = json.loads(response_json["ReturnValue"])

    orders_logistics_info: dict[str, dict[str, str]] = {}
    for order in return_value["datas"]:
        order_logistics_info = get_order_logistics_info(order)
        if order_logistics_info:
            order_id = order_logistics_info["order_id"]
            original_orders_logistics_info = orders_logistics_info.get(order_id)

            # 一个订单号查询到多个物流信息
            if original_orders_logistics_info:
                original_express_company = original_orders_logistics_info[
                    "express_company"
                ]
                original_express_number = original_orders_logistics_info[
                    "express_number"
                ]
                express_company = order_logistics_info["express_company"]
                express_number = order_logistics_info["express_number"]
                if express_company and express_number:
                    express_company = f"{original_express_company};{express_company}"
                    express_number = f"{original_express_number};{express_number}"
                    original_orders_logistics_info["express_company"] = express_company
                    original_orders_logistics_info["express_number"] = express_number

            # 首次记录这个订单号对应的物流信息
            else:
                express_company = order_logistics_info["express_company"]
                express_number = order_logistics_info["express_number"]
                if express_company and express_number:
                    orders_logistics_info[order_id] = order_logistics_info

    await close_browser()

    return orders_logistics_info


def update_row_logistics_info(row, order_logistics_info, file_path):
    """
    将物流信息更新填写到到表格行中
    """
    normalized_name = get_normalized_name(file_path)

    express_company = order_logistics_info["express_company"]
    express_number = order_logistics_info["express_number"]

    express_company_name = column_name_map[normalized_name].get("物流公司")
    if express_company_name and express_company_name in row:
        row[express_company_name] = express_company
    else:
        for (
            possible_express_company_column_name
        ) in possible_express_company_column_names:
            if possible_express_company_column_name in row:
                row[possible_express_company_column_name] = express_company

    express_number_name = column_name_map[normalized_name].get("物流单号")
    if express_number_name and express_number_name in row:
        row[express_number_name] = express_number
    else:
        for possible_express_number_column_name in possible_express_number_column_names:
            if possible_express_number_column_name in row:
                row[possible_express_number_column_name] = express_number


async def update_rows_logistics_info(rows, file_path):
    """
    将物流信息更新填写到到多个表格行中
    """
    order_id_name = column_name_map[get_normalized_name(file_path)]["线上订单号"]
    order_ids = [row[order_id_name] for row in rows]

    # 根据订单号，批量查询物流信息
    orders_logistics_info = await get_orders_logistics_info(order_ids)

    for row in rows:
        order_id = row[order_id_name]
        order_logistics_info = orders_logistics_info.get(order_id)
        if order_logistics_info:
            update_row_logistics_info(row, order_logistics_info, file_path)


async def generate_xlsx_file(input_file_path):
    """
    根据传入的路径xlsx文件路径，生成对应的回单文件，即在表格中填写物流信息
    """
    workbook = load_xlsx_or_xls_workbook(input_file_path)

    sheet = workbook[workbook.sheetnames[0]]
    column_names = [sheet.cell(1, column[0].column).value for column in sheet.columns]  # type: ignore

    # 如果表格没有物流公司或物流单号信息，手动在表格最右边追加两列
    if not possible_express_company_column_names.intersection(column_names):
        express_company_column = len(column_names) + 1
        sheet.insert_cols(express_company_column + 1)
        sheet.cell(row=1, column=express_company_column).value = "物流公司"
    if not possible_express_number_column_names.intersection(column_names):
        express_number_column = len(column_names) + 2
        sheet.insert_cols(express_number_column + 1)
        sheet.cell(row=1, column=express_number_column).value = "物流单号"

    rows = workbook_to_dicts(workbook)
    await update_rows_logistics_info(rows, input_file_path)

    for sheet_row in sheet.rows:
        for cell in sheet_row:
            if cell.row > 1:
                column_name = sheet.cell(row=1, column=cell.column).value
                column_name = str(column_name)
                row = rows[cell.row - 2]
                if row[column_name] and not cell.value:
                    cell.value = row[column_name]

    input_file_name = path.basename(input_file_path)
    [input_file_basename, _] = path.splitext(input_file_name)
    output_file_name = f"【回单】{input_file_basename}.xlsx".replace("【推单】", "")
    output_file_dir = path.dirname(input_file_path)
    output_file_path = path.join(output_file_dir, output_file_name)

    workbook.save(output_file_path)
    workbook.close()


async def export_from_erp321():
    """
    根据程序同文件夹下的“【推单】xxx.xlsx”文件，生成对应的“【回单】xxx.xlsx”文件，即包含物流信息的表格
    生成的文件，用来回复给渠道，告知订单对应的物流信息
    """
    file_paths = get_input_xlsl_file_paths()
    if not file_paths:
        raise RuntimeError("没有找到推单xlsx文件")

    for file_path in file_paths:
        await generate_xlsx_file(file_path)


if __name__ == "__main__":
    syncify(export_from_erp321)()
