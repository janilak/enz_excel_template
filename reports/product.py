from odoo import fields, models
from odoo.exceptions import UserError


class ProductXlsx(models.AbstractModel):
    _name = "report.enz_excel_template.productxlsx"
    _inherit = "report.report_xlsx.abstract"
    _description = "Product XLSX Report"

    def generate_xlsx_report(self, workbook, data, product):
        sheet = workbook.add_worksheet("Report")
        bold = workbook.add_format({"bold": True})
        date = workbook.add_format({'text_wrap': True, 'num_format': 'yyyy-mm-dd'})
        product_field = self.env['excel.formate.list'].search([('name', '=', 'enz_excel_template.productxlsx')])
        label_list = ['Internal Reference', 'Name', 'Price', 'Customer Taxes', 'Vendor Taxes', 'Can be Sold',
                      'Can be Purchased', 'Product Type', 'Sales Price', 'Cost', 'Invoicing Policy',
                      'Control Policy', 'Re-Invoice Expenses', 'Product Category']
        length = 14
        if product_field:
            length = product_field.list_length
            label_list = label_list + product_field.field_lines.mapped('excel_name')
        for line in range(0, length):
            sheet.write(0, line, label_list[line], bold)
        i = 1
        for obj in product:
            if obj.default_code:
                sheet.write(i, 0, obj.default_code)
            if obj.name:
                sheet.write(i, 1, obj.name)
            if obj.price:
                sheet.write(i, 2, obj.price)
            if obj.taxes_id:
                sheet.write(i, 3, obj.taxes_id.name)
            if obj.supplier_taxes_id:
                sheet.write(i, 4, obj.supplier_taxes_id.name)
            if obj.sale_ok:
                sheet.write(i, 5, obj.sale_ok)
            if obj.purchase_ok:
                sheet.write(i, 6, obj.purchase_ok)
            if obj.type:
                sheet.write(i, 7, obj.type)
            if obj.list_price:
                sheet.write(i, 8, obj.list_price)
            if obj.standard_price:
                sheet.write(i, 9, obj.standard_price)
            if obj.invoice_policy:
                sheet.write(i, 10, obj.invoice_policy)
            if obj.purchase_method:
                sheet.write(i, 11, obj.purchase_method)
            if obj.expense_policy:
                sheet.write(i, 12, obj.expense_policy)
            if obj.categ_id:
                sheet.write(i, 13, obj.categ_id.name)

            ##CALLING START
            total_list_no = 14
            if product_field:
                if product_field.field_lines:
                    for listline in product_field.field_lines:
                        model_id = self.env['ir.model'].sudo().search([('model', '=', 'product.template')]).id
                        field_id = self.env['ir.model.fields'].sudo().search(
                            [('name', '=', listline.field_id), ('model_id', '=', model_id)])
                        if field_id:
                            if getattr(obj, listline.field_id):
                                if listline.field_type == 'normal':
                                    sheet.write(i, total_list_no, getattr(obj, listline.field_id))
                                elif listline.field_type == 'many2one':
                                    manyone = getattr(obj, listline.field_id)
                                    if manyone:
                                        sheet.write(i, total_list_no, getattr(manyone, listline.manyonename))
                                else:
                                    sheet.write(i, total_list_no, getattr(obj, listline.field_id), date)
                            total_list_no = total_list_no + 1
                        else:
                            raise UserError("No Field Found in the Name " + listline.field_id)
            ##CALLING END
            i = i + 1
