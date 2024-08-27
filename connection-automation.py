import re
import win32com.client
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import os


def extract_purchase_order_details(body):
    order_number = re.search(r'order number is: (\d+)', body).group(1)
    po_number = re.search(r'Purchase Order #: (\w+)', body).group(1)

    product_details = []
    pattern = re.compile(r'Product Description\s+Availability\s+Qty\s+Unit Price\s+Total\s+([\s\S]+?)Subtotal')
    product_section = pattern.search(body)

    if product_section:
        product_lines = product_section.group(1).strip().split('\n')
        for line in product_lines:
            details = re.split(r'\s{2,}', line)
            if len(details) >= 4:
                description = details[0].strip()
                qty = details[-3].strip()
                unit_price = details[-2].strip()
                total = details[-1].strip()
                product_details.append([order_number, po_number, description, qty, unit_price, total])

    return product_details

def extract_shipping_confirmation_details(body):
    order_number = re.search(r'order # (\d+)', body).group(1)
    po_number = re.search(r'Purchase Order #: : (\w+)', body).group(1)
    tracking_number = re.search(r'Tracking Number:\s*(\w+)', body).group(1)

    product_details = []
    pattern = re.compile(r'Item #\s+Product Description\s+Qty\s+Shipping Information\s+([\s\S]+?)(?:CUSTOMER CARE|Thank you for placing your order)')
    product_section = pattern.search(body)

    if product_section:
        product_lines = product_section.group(1).strip().split('\n')
        for line in product_lines:
            details = re.split(r'\s{2,}', line.strip())
            if len(details) >= 3:
                item_number = details[0].strip()
                description = details[1].strip()
                qty = details[2].strip()
                product_details.append([order_number, po_number, item_number, description, qty, tracking_number])

    return product_details

def extract_backorder_details(body):
    order_number = re.search(r'order (\d+)', body).group(1)
    po_number = re.search(r'P\.O\. Number: (\w+)', body).group(1)

    product_details = []
    pattern = re.compile(r'Qty\s+Item #\s+Product\s+Shipping Information\s+([\s\S]+?)(?=\n\n|$)')

