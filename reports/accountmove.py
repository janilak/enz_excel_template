from odoo import models
from odoo.exceptions import UserError

class AccountMoveXlsx(models.AbstractModel):
    _name = "report.enz_excel_template.accountmovexlsx"
    _inherit = "report.report_xlsx.abstract"
    _description = "Account Move XLSX Report"

    def generate_xlsx_report(self, workbook, data, invoice):
        sheet = workbook.add_worksheet("Report")
        bold = workbook.add_format({"bold": True})
        date = workbook.add_format({'text_wrap': True, 'num_format': 'yyyy-mm-dd'})
        invoice_field = self.env['excel.formate.list'].search([('name', '=', 'enz_excel_template.accountmovexlsx')])
        label_list = ['Old Ref Number','Type','Partner', 'Payment Reference', 'Recipient Bank', 'Reference', 'Date',
                      'Invoice/Bill Date',
                      'Due Date', 'Payment Terms', 'Journal', 'Journal Items/Account', 'Journal Items/Partner',
                      'Journal Items/Label', 'Journal Items/Debit', 'Journal Items/Credit']
        length = 16
        if invoice_field:
            # length = invoice_field.list_length
            length = length + len(invoice_field.field_lines)
            label_list = label_list + invoice_field.field_lines.mapped('excel_name')
        for line in range(0, length):
            sheet.write(0, line, label_list[line], bold)
        i = 1
        for obj in invoice:
            if obj.name:
                sheet.write(i, 0, obj.name)
            if obj.move_type:
                sheet.write(i, 1, obj.move_type)
            if obj.partner_id:
                sheet.write(i, 2, obj.partner_id.name)
            if obj.payment_reference:
                sheet.write(i, 3, obj.payment_reference)
            if obj.partner_bank_id:
                sheet.write(i, 4, obj.partner_bank_id.acc_number)
            if obj.ref:
                sheet.write(i, 5, obj.ref)
            if obj.date:
                sheet.write(i, 6, obj.date, date)
            if obj.invoice_date:
                sheet.write(i, 7, obj.invoice_date, date)
            if obj.invoice_date_due:
                sheet.write(i, 8, obj.invoice_date_due, date)
            if obj.invoice_payment_term_id:
                sheet.write(i, 9, obj.invoice_payment_term_id.name)
            if obj.journal_id:
                sheet.write(i, 10, obj.journal_id.name)
            ##CALLING START
            total_list_no = 23
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
                                    sheet.write(i, total_list_no, getattr(obj, listline.field_id), date)
                            total_list_no = total_list_no + 1
                        else:
                            raise UserError("No Field Found in the Name " + listline.field_id)
            ##CALLING END
            # j = i
            # for inv in obj.invoice_line_ids:
            #     if inv.product_id:
            #         sheet.write(j, 11, inv.product_id.name)
            #     if inv.name:
            #         sheet.write(j, 12, inv.name)
            #     if inv.account_id:
            #         sheet.write(j, 13, inv.account_id.code + " " +inv.account_id.name)
            #     if inv.quantity:
            #         sheet.write(j, 14, inv.quantity)
            #     if inv.price_unit:
            #         sheet.write(j, 15, inv.price_unit)
            #     if inv.tax_ids:
            #         sheet.write(j, 16, inv.tax_ids.name)
            #     if inv.price_subtotal:
            #         sheet.write(j, 17, inv.price_subtotal)
            #     j = j + 1
            for journalline in obj.line_ids:
                if journalline.account_id:
                    sheet.write(i, 11, journalline.account_id.code +" "+journalline.account_id.name)
                if journalline.partner_id:
                    sheet.write(i, 12, journalline.partner_id.name)
                if journalline.name:
                    sheet.write(i, 13, journalline.name)
                if journalline.debit:
                    sheet.write(i, 14, journalline.debit)
                if journalline.credit:
                    sheet.write(i, 15, journalline.credit)
                i = i + 1
