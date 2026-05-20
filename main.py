BK1_FILEPATH = "./sample/aginprod.bk1"

# This dict holds the mappings for each product number for B01
file_layout_dict_b01 = {
    "type": {"offset": 0, "length": 1},
    "code": {"offset": 1, "length": 2},
    "product_number": {"offset": 3, "length": 10},
    "product_description": {"offset": 13, "length": 35},
    "pack": {"offset": 48, "length": 4},
    "size": {"offset": 52, "length": 4},
    "measure": {"offset": 56, "length": 2},
    "scan_description": {"offset": 58, "length": 12},
    "retail_by_weight": {"offset": 70, "length": 1},
    "charge_tax_1_pst_hst": {"offset": 71, "length": 1},
    "charge_tax_2_unit_gst": {"offset": 72, "length": 1},
    "charge_tax_3_case_gst": {"offset": 73, "length": 1},
    "bottle_deposit_amount": {"offset": 74, "length": 4},
    "item_status": {"offset": 78, "length": 1},
    "dsd_product": {"offset": 79, "length": 1},
    "food_service_or_retail": {"offset": 80, "length": 1},
    "stock_status_id": {"offset": 81, "length": 1},
    "restriction_01_prescription_drugs": {"offset": 82, "length": 1},
    "restriction_02_YTTobTax": {"offset": 83, "length": 1},
    "restriction_03_SKHBC": {"offset": 84, "length": 1},
    "restriction_04_HdwCoop": {"offset": 85, "length": 1},
    "restriction_05_HdwSelCoop": {"offset": 86, "length": 1},
    "restriction_06_BCTobTax": {"offset": 87, "length": 1},
    "restriction_07_BCHBC": {"offset": 88, "length": 1},
    "restriction_08_FoodCoop": {"offset": 89, "length": 1},
    "restriction_09_BCNeomycin": {"offset": 90, "length": 1},
    "restriction_10_ABLSDisease": {"offset": 91, "length": 1},
    "restriction_11_FrozenFood": {"offset": 92, "length": 1},
    "restriction_12_CropSupp": {"offset": 93, "length": 1},
    "restriction_13_WestBest": {"offset": 94, "length": 1},
    "restriction_14_ONTobTax": {"offset": 95, "length": 1},
    "restriction_15_FeedRestr": {"offset": 96, "length": 1},
    "restriction_16_HarmonCtry": {"offset": 97, "length": 1},
}


def product_number(product_line):
    print(product_line[file_layout_dict_b01["product_number"]["offset"]:file_layout_dict_b01["product_number"]["offset"] + file_layout_dict_b01["product_number"]["length"]])

def product_description(product_line):
    print(product_line[file_layout_dict_b01["product_description"]["offset"]:file_layout_dict_b01["product_description"]["offset"] + file_layout_dict_b01["product_description"]["length"]])

def main():
    products = []
    with open(BK1_FILEPATH, "r") as f:
        for line in f:
            products.append(line.strip())

    product_number(products[0])
    product_description(products[0])




main()