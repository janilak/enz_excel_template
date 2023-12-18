# -*- coding: utf-8 -*-
{
    'name': "Enzapps Default Excel Template",
    'author':
        'Enzapps Private Limited',
    'summary': """
This module will help create Default Excel Report
""",

    'description': """
This module will help create Default Excel Report
    """,
    "live_test_url": '',
    "website": "https://www.enzapps.com",
    'category': 'base',
    'version': '15.0',
    'depends':['base', 'account', 'stock', 'product','sale', 'sale_management', 'purchase','contacts','report_xlsx'],
    "images": ['static/description/icon.png'],
    'data': [
        'security/ir.model.access.csv',
        'reports/reports.xml',
        'views/excel_formate.xml',
        'data/data.xml',
    ],
    'demo': [],
    'installable': True,
    'application': True,
}
