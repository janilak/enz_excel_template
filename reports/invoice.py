from odoo import fields, models
from odoo.exceptions import UserError


class InvoiceXlsx(models.AbstractModel):
    _name = "report.enz_excel_template.invoicexlsx"
    _inherit = "report.report_xlsx.abstract"
    _description = "Invoice XLSX Report"

    def generate_xlsx_report(self, workbook, data, invoice):
        sheet = workbook.add_worksheet("Report")
        bold = workbook.add_format({"bold": True})
        date = workbook.add_format({'text_wrap':True,'num_format':'yyyy-mm-dd'})
        invoice_field = self.env['excel.formate.list'].search([('name','=','enz_excel_template.invoicexlsx')])
        label_list = ['Old Ref Number','Partner','Payment Reference','Recipient Bank','Invoice/Bill Date',
                      'Due Date','Payment Terms','Journal',
                      'Invoice lines/Product','Invoice lines/Label','Invoice lines/Account',
                      'Invoice lines/Quantity','Invoice lines/Unit Price','Invoice lines/Taxes','Invoice lines/Subtotal']
        length = 15
        if invoice_field:
            # length = invoice_field.list_length
            length = length + len(invoice_field.field_lines)
            label_list = label_list + invoice_field.field_lines.mapped('excel_name')
        for line in range(0,length):
            sheet.write(0, line ,label_list[line], bold)
        i = 1
        for obj in invoice:
            if obj.name:
                sheet.write(i, 0,obj.name)
            if obj.partner_id:
                sheet.write(i, 1, obj.partner_id.name)
            if obj.payment_reference:
                sheet.write(i, 2, obj.payment_reference)
            if obj.partner_bank_id:
                sheet.write(i, 3, obj.partner_bank_id.acc_number)
            if obj.invoice_date:
                sheet.write(i, 4, obj.invoice_date,date)
            if obj.invoice_date_due:
                sheet.write(i, 5, obj.invoice_date_due,date)
            if obj.invoice_payment_term_id:
                sheet.write(i, 6, obj.invoice_payment_term_id.name)
            if obj.journal_id:
                sheet.write(i, 7, obj.journal_id.name)
            ##CALLING START
            total_list_no = 15
            if invoice_field:
                if invoice_field.field_lines:
                    for listline in invoice_field.field_lines:
                        model_id = self.env['ir.model'].sudo().search([('model', '=', 'account.move')]).id
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
                                    sheet.write(i, total_list_no, getattr(obj, listline.field_id),date)
                            total_list_no = total_list_no + 1
                        else:
                            raise UserError("No Field Found in the Name " + listline.field_id)
            ##CALLING END
            for inv in obj.invoice_line_ids:
                if inv.product_id:
                    sheet.write(i, 8, inv.product_id.name)
                if inv.name:
                    sheet.write(i, 9, inv.name)
                if inv.account_id:
                    sheet.write(i, 10, inv.account_id.name)
                if inv.quantity:
                    sheet.write(i, 11, inv.quantity)
                if inv.price_unit:
                    sheet.write(i, 12, inv.price_unit)
                if inv.tax_ids:
                    sheet.write(i, 13, inv.tax_ids.name)
                if inv.price_subtotal:
                    sheet.write(i, 14, inv.price_subtotal)

                i = i+1
